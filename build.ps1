<#
.SYNOPSIS
Builds the single self-contained distributable from the sources under src\ and vendor\.

.DESCRIPTION
Concatenates the module files, inlines both XAML documents, and embeds each vendored
community script as a base64 payload, producing one .ps1 that can be copied to a USB stick
and run during OOBE with no other files present.

The vendored scripts are embedded as base64 of their exact bytes, never as text: any
re-encoding or line-ending change would invalidate Andrew Taylor's Authenticode signature.
Verify-embedding round-trips the checksum during the build and the result is checked again
at runtime by Test-ApVendorScript.

.PARAMETER OutputPath
Where to write the built script. Defaults to dist\Get-WindowsAutopilotImportGUICommunity.ps1.

.PARAMETER UpdateVendorManifest
Recompute vendor\VERSION.json from the files currently in vendor\ instead of building.
Use after refreshing the vendored scripts from upstream.

.PARAMETER SkipTests
Skip the Pester run that normally gates the build.

.EXAMPLE
.\build.ps1

.EXAMPLE
.\build.ps1 -UpdateVendorManifest
#>
[CmdletBinding()]
param(
    [string]$OutputPath,
    [switch]$UpdateVendorManifest,
    [switch]$SkipTests
)

$ErrorActionPreference = 'Stop'

$repoRoot = $PSScriptRoot
$vendorDir = Join-Path $repoRoot 'vendor'
$distDir = Join-Path $repoRoot 'dist'
if (-not $OutputPath) { $OutputPath = Join-Path $distDir 'Get-WindowsAutopilotImportGUICommunity.ps1' }

# Same order the development entry point uses: Logging first, then Config, then the rest.
$moduleFiles = @(
    'src\Private\Logging.ps1'
    'src\Private\Config.ps1'
    'src\Private\Elevation.ps1'
    'src\Private\DeviceInfo.ps1'
    'src\Private\VendorScript.ps1'
    'src\Private\ArgumentBuilder.ps1'
    'src\Private\ProgressParser.ps1'
    'src\Private\ScriptRunner.ps1'
    'src\Private\Connectivity.ps1'
    'src\Private\Preflight.ps1'
    'src\Private\Background.ps1'
    'src\Private\XamlLoader.ps1'
    'src\Private\Dialogs.ps1'
    'src\Public\Show-AutopilotImportGui.ps1'
)

function Write-Step { param([string]$Text) Write-Host "==> $Text" -ForegroundColor Cyan }
function Write-Ok { param([string]$Text) Write-Host "    $Text" -ForegroundColor DarkGray }

# ---------------------------------------------------------------- vendor manifest
function Update-VendorManifest {
    Write-Step 'Rebuilding vendor\VERSION.json'

    $manifestPath = Join-Path $vendorDir 'VERSION.json'
    $existing = if (Test-Path $manifestPath) { Get-Content $manifestPath -Raw | ConvertFrom-Json } else { $null }

    $scripts = @()
    foreach ($file in (Get-ChildItem $vendorDir -Filter *.ps1 | Sort-Object Name)) {
        $version = ''
        foreach ($line in (Get-Content $file.FullName -TotalCount 20)) {
            if ($line -match '^\s*\.VERSION\s+(\S+)') { $version = $Matches[1]; break }
        }

        $signature = Get-AuthenticodeSignature -LiteralPath $file.FullName
        $previous = if ($existing) { $existing.scripts | Where-Object { $_.file -eq $file.Name } } else { $null }

        $scripts += [ordered]@{
            file          = $file.Name
            version       = $version
            sha256        = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
            bytes         = $file.Length
            signer        = "$($signature.SignerCertificate.Subject)"
            role          = if ($previous) { $previous.role } elseif ($file.Name -like '*diagnostic*') { 'diagnostics' } else { 'engine' }
            psGalleryName = if ($previous) { $previous.psGalleryName } else { [System.IO.Path]::GetFileNameWithoutExtension($file.Name) }
        }

        Write-Ok "$($file.Name)  v$version  $($signature.Status)"
    }

    $manifest = [ordered]@{
        _comment = 'Third-party scripts vendored byte-exact so their Authenticode signatures remain valid. Do not reformat or re-save these files. Regenerate this manifest with build.ps1 -UpdateVendorManifest.'
        source   = if ($existing) { $existing.source } else { @{} }
        scripts  = $scripts
    }

    $manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
    Write-Ok "wrote $manifestPath"
}

if ($UpdateVendorManifest) { Update-VendorManifest; return }

# ---------------------------------------------------------------- tests
if (-not $SkipTests) {
    Write-Step 'Running Pester tests'
    $pester = Get-Module -ListAvailable Pester |
                Where-Object { $_.Version.Major -ge 5 } |
                Sort-Object Version -Descending | Select-Object -First 1

    if (-not $pester) {
        Write-Warning 'Pester 5+ not found; skipping tests. Install-Module Pester -Scope CurrentUser'
    }
    else {
        Import-Module $pester.Path -Force
        $config = New-PesterConfiguration
        $config.Run.Path = Join-Path $repoRoot 'tests'
        $config.Run.PassThru = $true
        $config.Output.Verbosity = 'Normal'
        $result = Invoke-Pester -Configuration $config

        if ($result.FailedCount -gt 0) {
            throw "$($result.FailedCount) test(s) failed. Fix them or pass -SkipTests."
        }
        Write-Ok "$($result.PassedCount) tests passed"
    }
}

# ---------------------------------------------------------------- verify sources
Write-Step 'Verifying sources'
foreach ($relative in $moduleFiles) {
    $path = Join-Path $repoRoot $relative
    if (-not (Test-Path $path)) { throw "Missing module file: $relative" }

    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$null, [ref]$errors)
    if ($errors.Count -gt 0) {
        throw "$relative has $($errors.Count) parse error(s), first: $($errors[0].Message)"
    }
}
Write-Ok "$($moduleFiles.Count) module files parse cleanly"

# ---------------------------------------------------------------- embed payloads
Write-Step 'Embedding XAML'
$xamlSources = @{
    MainWindow = Join-Path $repoRoot 'src\Views\MainWindow.xaml'
    Dark       = Join-Path $repoRoot 'src\Themes\Dark.xaml'
}

$xamlBlock = New-Object System.Text.StringBuilder
[void]$xamlBlock.AppendLine('$script:ApEmbeddedXaml = @{}')
foreach ($key in @('MainWindow', 'Dark')) {
    $path = $xamlSources[$key]
    if (-not (Test-Path $path)) { throw "Missing XAML: $path" }

    $content = Get-Content -LiteralPath $path -Raw

    # Validate before embedding: a XAML error is far cheaper to find here than at runtime
    # on a technician's bench.
    $probe = New-Object System.Xml.XmlDocument
    $probe.LoadXml($content)

    # A single-quoted here-string needs no escaping except doubling its own terminator.
    $escaped = $content -replace "'@", "''@"
    [void]$xamlBlock.AppendLine("`$script:ApEmbeddedXaml['$key'] = @'")
    [void]$xamlBlock.AppendLine($escaped)
    [void]$xamlBlock.AppendLine("'@")
    Write-Ok "$key ($([math]::Round($content.Length / 1KB, 1)) KB)"
}

Write-Step 'Embedding vendored scripts'
$manifestPath = Join-Path $vendorDir 'VERSION.json'
if (-not (Test-Path $manifestPath)) { throw "Missing vendor manifest: $manifestPath" }
$manifestJson = Get-Content -LiteralPath $manifestPath -Raw
$manifest = $manifestJson | ConvertFrom-Json

$vendorBlock = New-Object System.Text.StringBuilder
[void]$vendorBlock.AppendLine('$script:ApEmbeddedScripts = @{}')

foreach ($entry in $manifest.scripts) {
    $path = Join-Path $vendorDir $entry.file
    if (-not (Test-Path $path)) { throw "Vendored script listed in the manifest is missing: $($entry.file)" }

    $bytes = [System.IO.File]::ReadAllBytes($path)
    $actualHash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    if ($actualHash -ne $entry.sha256) {
        throw "Checksum mismatch for $($entry.file). The vendored file changed; re-run build.ps1 -UpdateVendorManifest after confirming why."
    }

    $base64 = [Convert]::ToBase64String($bytes)

    # Prove the round-trip before shipping it, so a corrupt payload can never reach a device.
    $roundTrip = [Convert]::FromBase64String($base64)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $roundTripHash = ($sha.ComputeHash($roundTrip) | ForEach-Object { $_.ToString('x2') }) -join ''
    }
    finally {
        $sha.Dispose()
    }
    if ($roundTripHash -ne $actualHash.ToLower()) { throw "Base64 round-trip failed for $($entry.file)." }

    # Wrap so the generated file stays readable in an editor and in a diff.
    $wrapped = ($base64 -split '(.{120})' | Where-Object { $_ }) -join [Environment]::NewLine

    [void]$vendorBlock.AppendLine("`$script:ApEmbeddedScripts['$($entry.file)'] = @'")
    [void]$vendorBlock.AppendLine($wrapped)
    [void]$vendorBlock.AppendLine("'@ -replace '\s',''")
    Write-Ok "$($entry.file) v$($entry.version) ($([math]::Round($bytes.Length / 1KB, 1)) KB source, $([math]::Round($base64.Length / 1KB, 1)) KB base64)"
}

[void]$vendorBlock.AppendLine("`$script:ApEmbeddedVersionJson = @'")
[void]$vendorBlock.AppendLine(($manifestJson -replace "'@", "''@"))
[void]$vendorBlock.AppendLine("'@")

# ---------------------------------------------------------------- assemble
Write-Step 'Assembling'

$engineEntry = $manifest.scripts | Where-Object { $_.role -eq 'engine' } | Select-Object -First 1
$buildStamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

# Single source of truth for the version: the value the GUI itself displays.
$guiSource = Get-Content -LiteralPath (Join-Path $repoRoot 'src\Public\Show-AutopilotImportGui.ps1') -Raw
if ($guiSource -match "\`$script:ApAppVersion\s*=\s*'([^']+)'") {
    $appVersion = $Matches[1]
}
else {
    throw 'Could not read $script:ApAppVersion from src\Public\Show-AutopilotImportGui.ps1.'
}
Write-Ok "version $appVersion"

$header = @"
<#PSScriptInfo
.VERSION $appVersion
.GUID 6f2b9c14-8d3e-4a71-9c5f-1b0e7a4d2c98
.AUTHOR Mark Orr
.COMPANYNAME orr365.tools
.COPYRIGHT (c) 2026 Mark Orr. MIT License.
.TAGS Windows Autopilot Intune EntraID DevicePreparation GUI WPF PowerShell OOBE
.LICENSEURI https://github.com/markorr321/Get-WindowsAutopilotImportCommunity/blob/main/LICENSE
.PROJECTURI https://github.com/markorr321/Get-WindowsAutopilotImportCommunity
.RELEASENOTES
$appVersion Single-file build. Autopilot v1 and v2 (Device Preparation) support.
#>

<#
.SYNOPSIS
A graphical front end for Windows Autopilot device registration, supporting both Autopilot v1
(hardware hash) and Autopilot v2 (Device Preparation identifiers).

.DESCRIPTION
Registers a Windows device for Autopilot from a resizable dark-themed window, driving the
Windows Autopilot Community script by Andrew S Taylor.

Supports both registration modes: Autopilot v1 uploads the 4K hardware hash, and Autopilot v2
imports the Manufacturer,Model,Serial device identifier used by Device Preparation policies,
which needs no hardware hash and therefore works on virtual machines.

Runs from a single self-contained file with nothing to download first. The window, the theme
and the Autopilot engine are all embedded, so it works during OOBE on a restricted network.

Includes live staged progress with a working cancel button, an offline CSV export of both the
hardware hash and the device identifier, batch import from CSV, a concurrent network
prerequisite check across the documented Autopilot and Intune endpoints, Autopilot
diagnostics, and a full session log.

Requires Windows PowerShell 5.1 and administrator rights.

.PARAMETER GroupTag
Pre-fills the group tag field.

.PARAMETER AssignedUser
Pre-fills the assigned user UPN field.

.PARAMETER Mode
Pre-selects the registration mode: v1 (hardware hash) or v2 (device preparation).

.PARAMETER NoElevate
Skip the automatic elevation prompt. The hardware hash cannot be read without
administrator rights, so Autopilot v1 will not work in this state.

.EXAMPLE
.\Get-WindowsAutopilotImportGUICommunity.ps1

.EXAMPLE
.\Get-WindowsAutopilotImportGUICommunity.ps1 -GroupTag FINANCE -Mode v2

.NOTES
Author  : Mark Orr (@markorr321)
Website : https://orr365.tools
License : MIT

GENERATED FILE. Do not edit by hand: change the sources under src\ and re-run build.ps1.
Built $buildStamp with engine v$($engineEntry.version).

Autopilot engine: get-windowsautopilotinfocommunity.ps1 (c) Andrew S Taylor, MIT, embedded
unmodified with its Authenticode signature intact.
Inspired by AutoPilot_Import_GUI (c) 2023 Ugur Koc, MIT.

.LINK
https://orr365.tools

.LINK
https://github.com/andrew-s-taylor/WindowsAutopilotInfo
#>
[CmdletBinding()]
param(
    [string]`$GroupTag = '',
    [string]`$AssignedUser = '',
    [ValidateSet('v1', 'v2')][string]`$Mode = '',
    [switch]`$NoElevate
)

`$ErrorActionPreference = 'Stop'

# In the single-file build there is no src\ tree; the app root is only used for the
# optional config.json sitting next to this script.
`$script:ApAppRoot = `$PSScriptRoot
if (-not `$script:ApAppRoot) {
    `$script:ApAppRoot = Split-Path -Parent `$MyInvocation.MyCommand.Definition
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
"@

$footer = @'

Initialize-ApLog | Out-Null
Import-ApConfig | Out-Null
Write-ApLog "Starting single-file build from $script:ApAppRoot"

# Must happen before anything auto-loads a module, and before any child process is started:
# a PowerShell 7 parent leaks its module path into this Windows PowerShell process, which
# shadows in-box modules with Core-only copies and breaks the engine before sign-in.
Repair-ApModulePath | Out-Null

# Elevate once for the whole session so child engine runs inherit the token.
if (-not $NoElevate -and -not (Test-ApElevated)) {
    Write-ApLog 'Not elevated; requesting elevation.'
    if (Invoke-ApSelfElevate -ScriptPath $MyInvocation.MyCommand.Path -BoundParameters $PSBoundParameters) {
        exit 0
    }
    Write-ApLog 'Continuing without elevation. The hardware hash will not be readable.' -Level WARN
}

try {
    # -Mode has a ValidateSet that rejects the empty-string default, so only forward
    # parameters that were actually supplied.
    $guiArgs = @{}
    if ($GroupTag) { $guiArgs.GroupTag = $GroupTag }
    if ($AssignedUser) { $guiArgs.AssignedUser = $AssignedUser }
    if ($Mode) { $guiArgs.Mode = $Mode }

    Show-AutopilotImportGui @guiArgs
}
catch {
    $message = $_.Exception.Message
    Write-ApLog "Fatal error: $message" -Level ERROR
    if ($_.ScriptStackTrace) { Write-ApLog $_.ScriptStackTrace -Level DEBUG }

    try {
        [System.Windows.MessageBox]::Show(
            "$message`r`n`r`nLog: $(Get-ApLogPath)",
            'Autopilot Import GUI', 'OK', 'Error') | Out-Null
    }
    catch {
        Write-Host "Autopilot Import GUI failed: $message" -ForegroundColor Red
    }
    exit 1
}
'@

$out = New-Object System.Text.StringBuilder
[void]$out.AppendLine($header)
[void]$out.AppendLine()
[void]$out.AppendLine('#region embedded resources')
[void]$out.AppendLine($xamlBlock.ToString())
[void]$out.AppendLine($vendorBlock.ToString())
[void]$out.AppendLine('#endregion embedded resources')
[void]$out.AppendLine()

foreach ($relative in $moduleFiles) {
    $path = Join-Path $repoRoot $relative
    [void]$out.AppendLine("#region $relative")
    [void]$out.AppendLine((Get-Content -LiteralPath $path -Raw).TrimEnd())
    [void]$out.AppendLine("#endregion $relative")
    [void]$out.AppendLine()
}

[void]$out.Append($footer)

if (-not (Test-Path $distDir)) { New-Item -ItemType Directory -Path $distDir -Force | Out-Null }

# UTF-8 with BOM: Windows PowerShell 5.1 assumes ANSI for a BOM-less file, which would
# corrupt any non-ASCII character in the XAML or the UI strings.
$encoding = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($OutputPath, $out.ToString(), $encoding)

# ---------------------------------------------------------------- verify output
Write-Step 'Verifying the built script'
$errors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile($OutputPath, [ref]$null, [ref]$errors)
if ($errors.Count -gt 0) {
    throw "The built script has $($errors.Count) parse error(s), first at line $($errors[0].Extent.StartLineNumber): $($errors[0].Message)"
}

$built = Get-Item $OutputPath
$lineCount = (Get-Content -LiteralPath $OutputPath).Count
Write-Ok "parses cleanly: $lineCount lines, $([math]::Round($built.Length / 1KB, 1)) KB"

Write-Host ''
Write-Host "Built $OutputPath" -ForegroundColor Green
Write-Host "Engine v$($engineEntry.version) embedded." -ForegroundColor DarkGray
Write-Host 'Validate before shipping:' -ForegroundColor DarkGray
Write-Host '    powershell.exe -STA -ExecutionPolicy Bypass -File tests\Test-Distribution.ps1' -ForegroundColor DarkGray

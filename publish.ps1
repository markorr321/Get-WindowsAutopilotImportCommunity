<#
.SYNOPSIS
Publishes the single-file build in dist\ to the PowerShell Gallery.

.DESCRIPTION
Validates the built script, checks the Gallery for a name or version clash, and publishes it.
Every check runs before anything is uploaded, because a published version cannot be replaced:
the Gallery only allows unlisting, so a bad publish means burning a version number.

Run build.ps1 first. This script never builds, so what gets published is exactly what was
tested.

.PARAMETER ApiKey
PowerShell Gallery API key. Get one from https://www.powershellgallery.com/account/apikeys.
If omitted, the POWERSHELL_GALLERY_API_KEY environment variable is used.

.PARAMETER WhatIf
Run every validation and report what would be published, without uploading.

.PARAMETER Path
Script to publish. Defaults to dist\Get-WindowsAutopilotImportGUICommunity.ps1.

.PARAMETER Repository
Target repository. Defaults to PSGallery.

.EXAMPLE
.\publish.ps1 -WhatIf

Validate everything and show what would happen. Always do this first.

.EXAMPLE
.\publish.ps1 -ApiKey 'oy2...'

.NOTES
Author  : Mark Orr (@markorr321)
Website : https://orr365.tools

Must run under Windows PowerShell 5.1: PowerShellGet's Publish-Script is the classic v2
implementation, and the build targets 5.1 anyway.
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [string]$ApiKey,
    [string]$Path,
    [string]$Repository = 'PSGallery'
)

$ErrorActionPreference = 'Stop'

$repoRoot = $PSScriptRoot
if (-not $Path) { $Path = Join-Path $repoRoot 'dist\Get-WindowsAutopilotImportGUICommunity.ps1' }

function Write-Step { param([string]$Text) Write-Host "==> $Text" -ForegroundColor Cyan }
function Write-Ok { param([string]$Text) Write-Host "    $Text" -ForegroundColor DarkGray }
function Write-Warn { param([string]$Text) Write-Host "    $Text" -ForegroundColor Yellow }

# ---------------------------------------------------------------- host
if ($PSVersionTable.PSEdition -eq 'Core') {
    throw 'Run this under Windows PowerShell 5.1 (powershell.exe), not PowerShell 7. Publish-Script comes from PowerShellGet v2.'
}

# The module path on this developer's machine can arrive polluted from a PowerShell 7 parent,
# which stops PowerShellGet resolving at all. Reuse the app's own repair.
$preflight = Join-Path $repoRoot 'src\Private\Preflight.ps1'
$logging = Join-Path $repoRoot 'src\Private\Logging.ps1'
if ((Test-Path $logging) -and (Test-Path $preflight)) {
    . $logging
    . $preflight
    $fixed = Get-ApWindowsPowerShellModulePath
    if ($fixed -ne $env:PSModulePath) {
        $env:PSModulePath = $fixed
        Write-Ok 'Normalised PSModulePath for Windows PowerShell.'
    }
}

# Load PowerShellGet up front with WhatIf suppressed. Its module body calls Set-Alias, and
# under -WhatIf an auto-load emits half a dozen "What if: Performing the operation Set Alias"
# lines that bury the actual validation output.
$savedWhatIf = $WhatIfPreference
try {
    $WhatIfPreference = $false
    Import-Module PowerShellGet -ErrorAction SilentlyContinue
}
finally {
    $WhatIfPreference = $savedWhatIf
}

# ---------------------------------------------------------------- validate the artefact
Write-Step "Validating $Path"
if (-not (Test-Path -LiteralPath $Path)) {
    throw "Not found: $Path. Run .\build.ps1 first."
}

$file = Get-Item -LiteralPath $Path
$info = Test-ScriptFileInfo -Path $Path   # throws with a specific reason if the metadata is bad

Write-Ok "Name        $($info.Name)"
Write-Ok "Version     $($info.Version)"
Write-Ok "Author      $($info.Author)"
Write-Ok "Size        $([math]::Round($file.Length / 1KB, 1)) KB"

if ($info.Name -ne $file.BaseName) {
    # Publish-Script derives the package name from the file name, not from PSScriptInfo.
    throw "The file name ($($file.BaseName)) must match the script name ($($info.Name)), because the Gallery package is named after the file."
}
foreach ($required in 'Author', 'Description', 'Version', 'Guid') {
    if (-not $info.$required) { throw "PSScriptInfo is missing $required." }
}

# The build is generated, so a stale dist is a real risk: publishing a version that does not
# match the sources is worse than failing here.
$guiSource = Join-Path $repoRoot 'src\Public\Show-AutopilotImportGui.ps1'
if (Test-Path $guiSource) {
    $raw = Get-Content -LiteralPath $guiSource -Raw
    if ($raw -match "\`$script:ApAppVersion\s*=\s*'([^']+)'") {
        $sourceVersion = $Matches[1]
        if ($sourceVersion -ne $info.Version.ToString()) {
            throw "dist is stale: sources say $sourceVersion but the built script says $($info.Version). Re-run .\build.ps1."
        }
        Write-Ok "Version matches the sources ($sourceVersion)."
    }
}

$newerSource = Get-ChildItem (Join-Path $repoRoot 'src'), (Join-Path $repoRoot 'vendor') -Recurse -File |
                Where-Object { $_.LastWriteTime -gt $file.LastWriteTime } |
                Select-Object -First 3
if ($newerSource) {
    Write-Warn 'These sources are newer than the build:'
    $newerSource | ForEach-Object { Write-Warn "  $($_.Name)" }
    throw 'dist is older than the sources. Re-run .\build.ps1.'
}

Write-Step 'Confirming the build still parses and self-verifies'
$errors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$errors)
if ($errors.Count -gt 0) { throw "The built script has $($errors.Count) parse error(s)." }
Write-Ok 'Parses cleanly.'

$smoke = Join-Path $repoRoot 'tests\Test-Distribution.ps1'
if (Test-Path $smoke) {
    Write-Ok 'Running the distribution smoke test...'
    $psExe = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $output = & $psExe -NoProfile -STA -ExecutionPolicy Bypass -File $smoke 2>&1
    if ($LASTEXITCODE -ne 0) {
        $output | Select-Object -Last 15 | ForEach-Object { Write-Warn "$_" }
        throw 'The distribution smoke test failed. Not publishing.'
    }
    Write-Ok ($output | Select-String -Pattern '^PASS:' | Select-Object -Last 1)
}

# ---------------------------------------------------------------- Gallery state
Write-Step "Checking $Repository for an existing package"
$published = $null
try {
    $published = Find-Script -Name $info.Name -Repository $Repository -ErrorAction Stop
}
catch {
    Write-Ok 'Name is not yet published; this will be the first version.'
}

if ($published) {
    Write-Ok "Currently published: $($published.Version) by $($published.Author)"

    if ($published.Author -ne $info.Author) {
        # Someone else owns the name: publishing would fail, but say so clearly first.
        throw "The name '$($info.Name)' on $Repository belongs to '$($published.Author)'. Choose a different script name."
    }
    if ([version]$published.Version -ge [version]$info.Version) {
        throw "Version $($info.Version) is not newer than the published $($published.Version). The Gallery cannot overwrite a version; bump `$script:ApAppVersion in src\Public\Show-AutopilotImportGui.ps1 and rebuild."
    }
}

# ---------------------------------------------------------------- publish
if (-not $ApiKey) { $ApiKey = $env:POWERSHELL_GALLERY_API_KEY }

if ($WhatIfPreference) {
    Write-Host ''
    Write-Host "WHATIF: would publish $($info.Name) $($info.Version) to $Repository" -ForegroundColor Yellow
    Write-Host '        All validation passed. Re-run with -ApiKey to publish for real.' -ForegroundColor DarkGray
    return
}

if (-not $ApiKey) {
    throw 'No API key. Pass -ApiKey, or set POWERSHELL_GALLERY_API_KEY. Keys come from https://www.powershellgallery.com/account/apikeys'
}

if (-not $PSCmdlet.ShouldProcess("$($info.Name) $($info.Version) -> $Repository", 'Publish')) { return }

Write-Step "Publishing $($info.Name) $($info.Version) to $Repository"
Publish-Script -Path $Path -NuGetApiKey $ApiKey -Repository $Repository -ErrorAction Stop

Write-Host ''
Write-Host "Published $($info.Name) $($info.Version)." -ForegroundColor Green
Write-Host 'The Gallery takes a few minutes to index it. Then, in a test VM:' -ForegroundColor DarkGray
Write-Host ''
Write-Host "    Install-Script -Name $($info.Name) -Scope AllUsers -Force" -ForegroundColor DarkGray
Write-Host "    $($info.Name).ps1" -ForegroundColor DarkGray

<#
.SYNOPSIS
Smoke-tests the single-file build in dist\.

.DESCRIPTION
Not a Pester test: it needs an STA apartment for WPF, which Pester does not guarantee under
Windows PowerShell. Run it after build.ps1 and before shipping.

Checks that matter for a build that will be copied to a USB stick and run in OOBE:
  * the file loads at all (AMSI is happy, no parse errors)
  * the embedded XAML produces a real window with all its named controls
  * the embedded engine extracts byte-exact, with its Authenticode signature intact
  * the argument builder still produces correct commands from the packed copy

.EXAMPLE
powershell.exe -STA -ExecutionPolicy Bypass -File tests\Test-Distribution.ps1
#>
[CmdletBinding()]
param(
    [string]$DistPath
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $DistPath) { $DistPath = Join-Path $repoRoot 'dist\Get-WindowsAutopilotImportGUICommunity.ps1' }

if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    throw 'Run this with -STA; WPF cannot create a Window from an MTA thread.'
}

$script:Pass = 0
$script:Fail = 0
function Check {
    param([string]$Label, $Condition, [string]$Detail = '')
    if ($Condition) { $script:Pass++; Write-Host "  PASS  $Label" -ForegroundColor Green }
    else { $script:Fail++; Write-Host "  FAIL  $Label $Detail" -ForegroundColor Red }
}

Write-Host "Testing $DistPath"
Check 'dist file exists' (Test-Path -LiteralPath $DistPath)
if (-not (Test-Path -LiteralPath $DistPath)) { exit 1 }

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml

# ---- load the definitions without running the entry point ----
# The built file ends with a param block plus a launch sequence. Take everything up to the
# final #endregion so we get the embedded resources and all functions, but no ShowDialog.
Write-Host "`n== load =="
$text = Get-Content -LiteralPath $DistPath -Raw
$cut = $text.LastIndexOf('#endregion src\Public')
Check 'found the end of the module regions' ($cut -gt 0)

$body = $text.Substring(0, $text.IndexOf("`n", $cut) + 1)

# Drop the param() block: dot-sourcing a script that declares parameters would bind this
# test's own arguments to it. Locate it via the AST rather than a regex, because a
# non-greedy match stops at the first ')' inside [ValidateSet('v1','v2')].
$bodyAst = [System.Management.Automation.Language.Parser]::ParseInput($body, [ref]$null, [ref]$null)
if ($bodyAst.ParamBlock) {
    $extent = $bodyAst.ParamBlock.Extent
    $body = $body.Remove($extent.StartOffset, $extent.EndOffset - $extent.StartOffset)
}
Check 'param block removed' ($null -eq ([System.Management.Automation.Language.Parser]::ParseInput($body, [ref]$null, [ref]$null)).ParamBlock)

$loader = Join-Path $env:TEMP ('ap-dist-probe-{0}.ps1' -f ([guid]::NewGuid().ToString('N')))
Set-Content -LiteralPath $loader -Value $body -Encoding UTF8
try {
    $script:ApAppRoot = Split-Path -Parent $DistPath
    . $loader
    Check 'definitions loaded (no AMSI block, no parse errors)' $true
}
catch {
    Check 'definitions loaded (no AMSI block, no parse errors)' $false $_.Exception.Message
    Remove-Item -LiteralPath $loader -Force -ErrorAction SilentlyContinue
    exit 1
}
finally {
    Remove-Item -LiteralPath $loader -Force -ErrorAction SilentlyContinue
}

Initialize-ApLog | Out-Null
Import-ApConfig | Out-Null

# ---- embedded resources ----
Write-Host "`n== embedded resources =="
Check 'XAML payloads embedded' ($script:ApEmbeddedXaml.Count -eq 2) "count $($script:ApEmbeddedXaml.Count)"
Check 'vendor payloads embedded' ($script:ApEmbeddedScripts.Count -ge 1) "count $($script:ApEmbeddedScripts.Count)"
Check 'vendor manifest embedded' ([bool]$script:ApEmbeddedVersionJson)

$manifest = Get-ApVendorManifest
Check 'manifest parses' ($null -ne $manifest -and $manifest.scripts.Count -ge 1)

# ---- engine extraction and integrity ----
Write-Host "`n== engine extraction =="
$engine = Resolve-ApEngineScript -Name Engine -AllowInstall:$false
Check 'engine extracted' (Test-Path -LiteralPath $engine) $engine

$entry = $manifest.scripts | Where-Object { $_.role -eq 'engine' } | Select-Object -First 1
$extractedBytes = (Get-Item -LiteralPath $engine).Length
Check 'extracted size matches manifest' ($extractedBytes -eq $entry.bytes) "expected $($entry.bytes), got $extractedBytes"

$verify = Test-ApVendorScript -Path $engine
Check 'checksum matches after base64 round-trip' ($verify.HashMatches -eq $true) "expected $($verify.ExpectedHash) got $($verify.ActualHash)"
# The whole point of embedding raw bytes rather than text: re-encoding would break this.
Check 'Authenticode signature still Valid' ($verify.SignatureStatus -eq 'Valid') $verify.SignatureStatus
Check 'signer is the expected publisher' ($verify.Signer -like '*ANDREWSTAYLOR*') $verify.Signer

$diag = Resolve-ApEngineScript -Name Diagnostics -AllowInstall:$false
$diagVerify = Test-ApVendorScript -Path $diag
Check 'diagnostics script extracts and verifies' ($diagVerify.HashMatches -eq $true -and $diagVerify.SignatureStatus -eq 'Valid')

# ---- window ----
Write-Host "`n== window =="
$ui = New-ApMainWindow
Check 'window builds from embedded XAML' ($null -ne $ui.Window)
Check 'named controls resolved' ($ui.Elements.Count -ge 90) "count $($ui.Elements.Count)"
foreach ($required in @('RegisterButton', 'GroupTagCombo', 'ModeV1', 'ModeV2', 'NetworkGrid', 'DeviceGrid', 'StatusProgress', 'RegisterOutput', 'AdvDiagnosticsButton', 'AdvDiagnosticsOnlineCheck')) {
    Check "control '$required' present" ($ui.Elements.ContainsKey($required))
}

# ---- argument building ----
Write-Host "`n== argument building =="
$req = New-ApRegistrationRequest
$req.GroupTag = 'FINANCE'
$req.Reboot = $true
$p = (Build-ApEngineArguments $req).Parameters
Check 'v1 emits -Online and -GroupTag' ($p['Online'] -and $p['GroupTag'] -eq 'FINANCE')
Check 'reboot still coerces -Assign' ($p['Assign'] -eq $true)

$req2 = New-ApRegistrationRequest
$req2.Mode = 'v2'
$p2 = (Build-ApEngineArguments $req2).Parameters
Check 'v2 emits -identifier only' ($p2['identifier'] -and -not $p2.Contains('GroupTag'))

# ---- diagnostics launcher ----
# The Advanced page decides between a local read and a Graph lookup purely by whether it puts
# an 'Online' key in the parameter hashtable, which is also what gates the sign-in prep block.
Write-Host "`n== diagnostics launcher =="
$diagTmp = Join-Path ([System.IO.Path]::GetTempPath()) ('ap-diag-{0}' -f ([guid]::NewGuid().ToString('N')))
New-Item -ItemType Directory -Path $diagTmp -Force | Out-Null
try {
    $launcher = Join-Path $diagTmp 'launcher.ps1'

    New-ApLauncherScript -EnginePath $diag -Parameters ([ordered]@{ Online = $true }) `
                         -LogPath (Join-Path $diagTmp 'run.log') -LauncherPath $launcher | Out-Null
    $onlineText = Get-Content -LiteralPath $launcher -Raw
    Check 'online diagnostics splats -Online' ($onlineText -match 'Online\s*=\s*\$true')
    Check 'online diagnostics prepares the sign-in module' ($onlineText -match 'Microsoft\.Graph\.Authentication')

    New-ApLauncherScript -EnginePath $diag -Parameters ([ordered]@{}) `
                         -LogPath (Join-Path $diagTmp 'run.log') -LauncherPath $launcher | Out-Null
    $localText = Get-Content -LiteralPath $launcher -Raw
    Check 'local diagnostics never touches the gallery' ($localText -notmatch 'Install-Module')
}
finally {
    Remove-Item -LiteralPath $diagTmp -Recurse -Force -ErrorAction SilentlyContinue
}

# ---- device info ----
Write-Host "`n== device =="
$info = Get-ApDeviceInfo
Check 'serial read' ([bool]$info.SerialNumber)
Check 'device identifier is a three-part triple' (($info.DeviceIdentifier -split ',').Count -eq 3) $info.DeviceIdentifier

Write-Host ''
if ($script:Fail -gt 0) {
    Write-Host "PASS: $script:Pass   FAIL: $script:Fail" -ForegroundColor Red
    exit 1
}
Write-Host "PASS: $script:Pass   FAIL: $script:Fail" -ForegroundColor Green
exit 0

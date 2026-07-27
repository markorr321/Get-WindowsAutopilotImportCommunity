<#PSScriptInfo
.VERSION 1.0.0
.GUID 6f2b9c14-8d3e-4a71-9c5f-1b0e7a4d2c98
.AUTHOR Mark Orr
.COMPANYNAME orr365.tools
.COPYRIGHT (c) 2026 Mark Orr. MIT License.
.TAGS Windows Autopilot Intune EntraID DevicePreparation GUI WPF PowerShell OOBE
.LICENSEURI https://github.com/markorr321/Get-WindowsAutopilotImportCommunity/blob/main/LICENSE
.PROJECTURI https://github.com/markorr321/Get-WindowsAutopilotImportCommunity
.RELEASENOTES
1.0.0 Initial release. Autopilot v1 and v2 (Device Preparation) support, staged progress,
      working network prerequisite check, offline export.
#>

<#
.SYNOPSIS
A graphical front end for the Windows Autopilot Community import script, supporting both
Autopilot v1 (hardware hash) and Autopilot v2 (Device Preparation identifiers).

.DESCRIPTION
Drives get-windowsautopilotinfocommunity.ps1 by Andrew S Taylor, streaming its output into
a resizable dark-themed window with real staged progress, a cancel button, an offline CSV
export, a working network prerequisite check and a full session log.

This is the development entry point: it dot-sources the modules under src\. Run build.ps1
to produce the single self-contained script in dist\ for deployment to a USB stick or OOBE.

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
.\Get-WindowsAutopilotImportGUICommunity.ps1 -GroupTag FINANCE -Mode v1

.NOTES
Author  : Mark Orr (@markorr321)
Website : https://orr365.tools
License : MIT

Requires Windows PowerShell 5.1 and administrator rights.
Autopilot engine: get-windowsautopilotinfocommunity.ps1 (c) Andrew S Taylor, MIT.
Inspired by AutoPilot_Import_GUI (c) 2023 Ugur Koc, MIT.

.LINK
https://orr365.tools

.LINK
https://github.com/andrew-s-taylor/WindowsAutopilotInfo
#>
[CmdletBinding()]
param(
    [string]$GroupTag = '',
    [string]$AssignedUser = '',
    [ValidateSet('v1', 'v2')][string]$Mode = '',
    [switch]$NoElevate
)

$ErrorActionPreference = 'Stop'

# Resolve the repo root before anything else: every module path and the vendor folder
# lookup hang off it. $PSScriptRoot is empty when a script is piped into iex, so fall back
# to the invocation path.
$script:ApAppRoot = $PSScriptRoot
if (-not $script:ApAppRoot) {
    $script:ApAppRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
}

# WPF must be available before any XAML is touched.
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml

# Load order matters: Logging first because everything else logs, Config next because the
# rest read preferences from it.
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

foreach ($relative in $moduleFiles) {
    $path = Join-Path $script:ApAppRoot $relative
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Required module not found: $path. Run this from a full checkout, or use the built script in dist\."
    }
    . $path
}

Initialize-ApLog | Out-Null
Import-ApConfig | Out-Null
Write-ApLog "Starting from $script:ApAppRoot"

# Must happen before anything auto-loads a module, and before any child process is started:
# a PowerShell 7 parent leaks its module path into this Windows PowerShell process, which
# shadows in-box modules with Core-only copies and breaks the engine before sign-in.
Repair-ApModulePath | Out-Null

# Elevate once for the whole session so child engine runs inherit the token, rather than
# prompting per action the way the original GUI did.
if (-not $NoElevate -and -not (Test-ApElevated)) {
    Write-ApLog 'Not elevated; requesting elevation.'
    if (Invoke-ApSelfElevate -ScriptPath $MyInvocation.MyCommand.Path -BoundParameters $PSBoundParameters) {
        exit 0
    }
    Write-ApLog 'Continuing without elevation. The hardware hash will not be readable.' -Level WARN
}

try {
    # Only forward parameters that were actually supplied: -Mode has a ValidateSet that
    # rejects the empty-string default.
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

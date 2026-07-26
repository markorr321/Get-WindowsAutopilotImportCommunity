# Elevation.ps1 -- administrator detection and one-shot self-relaunch.
#
# Reading the Autopilot hardware hash requires administrator rights: the
# MDM_DevDetail_Ext01 class in root/cimv2/mdm/dmmap is not readable as a standard
# user (see the community script, get-windowsautopilotinfocommunity.ps1:2121).
#
# Elevating the GUI once means every child process it launches inherits the token,
# so the tech sees a single UAC prompt for the whole session instead of one per
# action -- the original GUI used -Verb RunAs on each button.
#
# In OOBE (Shift+F10) the shell already runs as SYSTEM, so this is a no-op there.

function Test-ApElevated {
    <#
    .SYNOPSIS
    True when the current process holds the local Administrators role.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    try {
        $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Invoke-ApSelfElevate {
    <#
    .SYNOPSIS
    Relaunches the entry script elevated and reports whether the caller should exit.

    .DESCRIPTION
    Returns $true when a new elevated process was started (the caller must exit).
    Returns $false when already elevated, or when the user declined the UAC prompt
    and chose to continue with reduced functionality.

    .PARAMETER ScriptPath
    Path of the entry script to relaunch.

    .PARAMETER BoundParameters
    The entry script's $PSBoundParameters, forwarded to the elevated instance.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [hashtable]$BoundParameters = @{}
    )

    if (Test-ApElevated) { return $false }

    $argList = New-Object System.Collections.Generic.List[string]
    $argList.Add('-NoProfile')
    $argList.Add('-ExecutionPolicy'); $argList.Add('Bypass')
    $argList.Add('-STA')
    $argList.Add('-File'); $argList.Add('"{0}"' -f $ScriptPath)

    foreach ($key in $BoundParameters.Keys) {
        $value = $BoundParameters[$key]
        if ($value -is [switch]) {
            if ($value.IsPresent) { $argList.Add("-$key") }
        }
        elseif ($null -ne $value -and "$value" -ne '') {
            $argList.Add("-$key")
            $argList.Add('"{0}"' -f ($value -replace '"', '""'))
        }
    }

    try {
        Start-Process -FilePath (Get-ApPowerShellPath) `
                      -ArgumentList $argList.ToArray() `
                      -Verb RunAs -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        # 1223 == ERROR_CANCELLED, i.e. the user dismissed the UAC prompt.
        Write-ApLog "Elevation declined or failed: $($_.Exception.Message)" -Level WARN
        return $false
    }
}

function Get-ApPowerShellPath {
    <#
    .SYNOPSIS
    Full path to the Windows PowerShell 5.1 host.

    .DESCRIPTION
    Always Windows PowerShell, never pwsh: the community script pins
    microsoft.graph.authentication to <= 2.9.1 and relies on Windows PowerShell
    behaviour, and WPF hosting is most reliable there. Uses the native-architecture
    path so a 32-bit host does not get redirected into SysWOW64.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()

    $sysNative = Join-Path $env:WINDIR 'sysnative\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path -LiteralPath $sysNative) { return $sysNative }

    $system32 = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path -LiteralPath $system32) { return $system32 }

    return 'powershell.exe'
}

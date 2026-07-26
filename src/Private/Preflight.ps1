# Preflight.ps1 -- checks that must pass before an online run can possibly succeed.
#
# Interactive sign-in happens inside the engine, which does:
#
#     $module = Import-Module microsoft.graph.authentication -PassThru -ErrorAction Ignore
#     if (-not $module) { Install-Module microsoft.graph.authentication -MaximumVersion 2.9.1 }
#     ...
#     Import-Module Microsoft.Graph.Authentication     # unversioned: highest wins
#
# Two consequences worth pre-empting:
#
#   * -ErrorAction Ignore means a *broken* install looks the same as a missing one, so the
#     engine tries to install 2.9.1 and then imports unversioned anyway, which picks the
#     highest installed version. A broken newer version therefore defeats both the engine's
#     version pin and its own repair attempt.
#   * A module under a OneDrive-redirected Documents folder is frequently incomplete: the
#     manifest is present but files it references (.format.ps1xml, assemblies) are still
#     cloud-only placeholders. The import fails with "the member 'FormatsToProcess' in the
#     module manifest is not valid".
#
# Without this check the tech sees only a non-zero exit code and no browser prompt.

$script:ApGraphModuleName = 'Microsoft.Graph.Authentication'

# The version the engine pins with -MaximumVersion; what a repair should install.
$script:ApGraphPinnedVersion = '2.9.1'

function Get-ApWindowsPowerShellModulePath {
    <#
    .SYNOPSIS
    A PSModulePath containing only module locations valid for Windows PowerShell 5.1.

    .DESCRIPTION
    Launching this tool from a PowerShell 7 terminal leaks PS7's PSModulePath into the 5.1
    process, and children inherit it. Observed on a real machine:

        C:\Users\<u>\OneDrive\Documents\PowerShell\Modules
        C:\Program Files\PowerShell\Modules
        C:\Program Files\WindowsApps\Microsoft.PowerShell_7.6.4.0_x64__...\Modules
        C:\Program Files\WindowsPowerShell\Modules
        C:\WINDOWS\system32\WindowsPowerShell\v1.0\Modules

    The PS7 entries sort first and shadow in-box modules with Core-only copies, so 5.1 either
    refuses to load them or cannot see them at all. That produced two different symptoms from
    one cause: "the 'Set-ExecutionPolicy' command was found in the module
    'Microsoft.PowerShell.Security', but the module could not be loaded", and
    "the term 'Get-PackageProvider' is not recognized". Both aborted the engine before any
    sign-in prompt could appear.

    Filtering rule: a genuine Windows PowerShell module path always contains
    "WindowsPowerShell". Anything else mentioning "PowerShell" belongs to PS7 and is dropped.
    Unrelated custom paths (a management agent's module directory, say) mention neither and
    are preserved.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()

    # Documents may be redirected (OneDrive), so ask Windows rather than assuming.
    $documents = [Environment]::GetFolderPath('MyDocuments')
    if (-not $documents) { $documents = Join-Path $env:USERPROFILE 'Documents' }

    $canonical = @(
        (Join-Path $documents 'WindowsPowerShell\Modules')
        (Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules')
        (Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\Modules')
    )

    $preserved = @()
    if ($env:PSModulePath) {
        $preserved = @(
            $env:PSModulePath -split ';' |
                Where-Object { $_ -and $_.Trim() } |
                ForEach-Object { $_.TrimEnd('\') } |
                Where-Object {
                    # Drop PowerShell 7 locations; keep everything unrelated.
                    -not ($_ -like '*powershell*' -and $_ -notlike '*windowspowershell*')
                }
        )
    }

    # Case-insensitive dedupe: Windows paths differ only in casing between sources
    # (System32 vs system32), and a List[string].Contains would keep both.
    $ordered = New-Object System.Collections.Generic.List[string]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($path in @($canonical + $preserved)) {
        $trimmed = "$path".TrimEnd('\')
        if (-not $trimmed) { continue }
        if ($seen.Add($trimmed)) { $ordered.Add($trimmed) }
    }

    return ($ordered -join ';')
}

function Repair-ApModulePath {
    <#
    .SYNOPSIS
    Rewrites $env:PSModulePath for this process so Windows PowerShell modules resolve.

    .DESCRIPTION
    Called once at startup. Child processes inherit the corrected value, so the engine
    launcher, the Windows Update console and the Graph repair console all benefit.
    #>
    [CmdletBinding()]
    param()

    $before = "$env:PSModulePath"
    $after = Get-ApWindowsPowerShellModulePath
    $env:PSModulePath = $after

    if ($before -ne $after) {
        $removed = @(
            $before -split ';' | Where-Object {
                $_ -and ($_ -like '*powershell*' -and $_ -notlike '*windowspowershell*')
            }
        )
        if ($removed.Count -gt 0) {
            Write-ApLog "Removed $($removed.Count) PowerShell 7 module path(s) that would shadow Windows PowerShell modules: $($removed -join '; ')" -Level WARN
        }
        Write-ApLog "PSModulePath set to: $after"
    }

    return $after
}

function Test-ApGraphModule {
    <#
    .SYNOPSIS
    Reports whether the Graph authentication module can actually be imported.

    .DESCRIPTION
    Self-contained so it can be copied into a bare background runspace: no logging, no
    config. Imports for real rather than merely checking that a folder exists, because the
    failure mode being detected is precisely a present-but-incomplete install.

    Import it the same way the engine does -- unversioned, so the highest installed version
    wins -- and only then look for a lower version that does work, which is what makes the
    remedy actionable.

    .OUTPUTS
    PSCustomObject: Available, Version, Path, Error, InstalledVersions, WorkingVersion.
    #>
    [CmdletBinding()]
    param()

    $result = [ordered]@{
        Available         = $false
        Version           = ''
        Path              = ''
        Error             = ''
        InstalledVersions = @()
        WorkingVersion    = ''
    }

    $name = 'Microsoft.Graph.Authentication'

    $available = @(Get-Module -ListAvailable -Name $name -ErrorAction SilentlyContinue |
                    Sort-Object Version -Descending)
    $result.InstalledVersions = @($available | ForEach-Object { "$($_.Version)" })

    if ($available.Count -eq 0) {
        $result.Error = 'The module is not installed.'
        return [pscustomobject]$result
    }

    # Exactly what the engine does.
    try {
        $imported = Import-Module -Name $name -PassThru -ErrorAction Stop
        $result.Available = $true
        $result.Version = "$($imported.Version)"
        $result.Path = "$($imported.ModuleBase)"
        return [pscustomobject]$result
    }
    catch {
        $result.Error = $_.Exception.Message
        $result.Version = "$($available[0].Version)"
        $result.Path = "$($available[0].ModuleBase)"
    }

    # Is there an older version that does load? If so, the fix is to remove the broken one.
    foreach ($candidate in ($available | Select-Object -Skip 1)) {
        try {
            Import-Module -Name $candidate.Path -PassThru -ErrorAction Stop | Out-Null
            $result.WorkingVersion = "$($candidate.Version)"
            break
        }
        catch {
            continue
        }
    }

    return [pscustomobject]$result
}

function Get-ApGraphModuleAdvice {
    <#
    .SYNOPSIS
    Turns a Test-ApGraphModule result into text a technician can act on.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Check)

    if ($Check.Available) { return '' }

    $lines = New-Object System.Collections.Generic.List[string]

    if ($Check.InstalledVersions.Count -eq 0) {
        $lines.Add("The $script:ApGraphModuleName module is not installed, so signing in is not possible.")
        $lines.Add('The engine will try to install it automatically, which needs access to the PowerShell Gallery.')
        return ($lines -join [Environment]::NewLine)
    }

    $lines.Add("$script:ApGraphModuleName $($Check.Version) is installed but cannot be loaded, so sign-in will fail before a browser prompt appears.")
    $lines.Add('')
    $lines.Add("Location: $($Check.Path)")
    $lines.Add("Reported: $($Check.Error)")

    if ($Check.Path -like '*OneDrive*') {
        $lines.Add('')
        $lines.Add('That path is inside OneDrive. Redirected Documents folders often leave module files as cloud-only placeholders, so the manifest is present but the files it references are not.')
    }

    if ($Check.WorkingVersion) {
        $lines.Add('')
        $lines.Add("Version $($Check.WorkingVersion) is also installed and does load, but PowerShell imports the highest version, so the broken one wins.")
    }

    return ($lines -join [Environment]::NewLine)
}

function Get-ApGraphRepairScript {
    <#
    .SYNOPSIS
    A script that removes every installed copy of the module and reinstalls the pinned version.

    .DESCRIPTION
    Installs to AllUsers (under Program Files) deliberately: that is outside any redirected
    Documents folder, which is what caused the broken install in the first place.
    #>
    [CmdletBinding()]
    param()

    return @"
`$Host.UI.RawUI.WindowTitle = 'Autopilot Import GUI - Repair Graph module'
`$ErrorActionPreference = 'Continue'

function Write-Step { param([string]`$Text) Write-Host "==> `$Text" -ForegroundColor Cyan }

Write-Host ''
Write-Host ' Repairing $script:ApGraphModuleName' -ForegroundColor White
Write-Host ''

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    Write-Step 'Currently installed versions:'
    Get-Module -ListAvailable -Name $script:ApGraphModuleName |
        Select-Object Version, ModuleBase | Format-Table -AutoSize | Out-String | Write-Host

    Write-Step 'Removing all installed versions...'
    try {
        Uninstall-Module -Name $script:ApGraphModuleName -AllVersions -Force -ErrorAction Stop
        Write-Host '    Removed via Uninstall-Module.' -ForegroundColor Gray
    }
    catch {
        Write-Host "    Uninstall-Module could not remove it: `$(`$_.Exception.Message)" -ForegroundColor Yellow
        # A hand-copied or partially synced module is not registered with PowerShellGet, so
        # Uninstall-Module cannot see it. Delete the folders directly.
        Write-Step 'Deleting module folders directly...'
        foreach (`$found in (Get-Module -ListAvailable -Name $script:ApGraphModuleName)) {
            try {
                Remove-Item -LiteralPath `$found.ModuleBase -Recurse -Force -ErrorAction Stop
                Write-Host "    Deleted `$(`$found.ModuleBase)" -ForegroundColor Gray
            }
            catch {
                Write-Host "    Could not delete `$(`$found.ModuleBase): `$(`$_.Exception.Message)" -ForegroundColor Red
            }
        }
    }

    Write-Step 'Installing version $script:ApGraphPinnedVersion for all users...'
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:`$false -ErrorAction SilentlyContinue | Out-Null
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
    Install-Module -Name $script:ApGraphModuleName -RequiredVersion $script:ApGraphPinnedVersion ``
                   -Scope AllUsers -Force -AllowClobber -Confirm:`$false -ErrorAction Stop

    Write-Step 'Verifying the module imports...'
    Import-Module -Name $script:ApGraphModuleName -Force -ErrorAction Stop
    `$loaded = Get-Module -Name $script:ApGraphModuleName

    Write-Host ''
    Write-Host " Repaired. Loaded version `$(`$loaded.Version) from `$(`$loaded.ModuleBase)" -ForegroundColor Green
    Write-Host ' Close and reopen Autopilot Import GUI, then try signing in again.' -ForegroundColor Gray
}
catch {
    Write-Host ''
    Write-Host " Repair failed: `$(`$_.Exception.Message)" -ForegroundColor Red
    Write-Host ' You can repair it manually with:' -ForegroundColor Gray
    Write-Host '   Uninstall-Module $script:ApGraphModuleName -AllVersions -Force' -ForegroundColor Gray
    Write-Host '   Install-Module $script:ApGraphModuleName -RequiredVersion $script:ApGraphPinnedVersion -Scope AllUsers -Force' -ForegroundColor Gray
}

Write-Host ''
Read-Host 'Press Enter to close'
"@
}

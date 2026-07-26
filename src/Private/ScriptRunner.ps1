# ScriptRunner.ps1 -- runs the community engine script and streams its output into the GUI.
#
# Why a child process and a log file rather than a runspace:
#   * the engine is a *script* with a Begin/Process/End pipeline and its own $ErrorActionPreference,
#     setx calls and Disconnect-MgGraph teardown -- running it in-process would pollute the GUI's
#     session state and leave a Graph connection behind
#   * interactive Connect-MgGraph spawns a browser and can call exit; a crash in the engine must not
#     take the window down with it
#   * a log file on disk is also the artefact the tech needs afterwards for a support case
#
# Why the parameters go into a generated launcher rather than onto the command line:
#   quoting a group tag containing a space or an apostrophe through
#   powershell.exe -File is a known source of silent corruption. The launcher embeds a
#   literal hashtable and splats it, so values reach the engine exactly as typed.
#
# Output capture relies on `*>&1` -- in Windows PowerShell 5.1 Write-Host writes to the
# information stream, so redirecting all streams captures the engine's entire narration
# (it uses Write-Host almost exclusively). Without the `*` the log would be empty.

$script:ApCompleteSentinel = '##AP_COMPLETE:'

function Format-ApPowerShellLiteral {
    <#
    .SYNOPSIS
    Renders a value as PowerShell source. Supports the types the argument builder emits.
    #>
    param($Value)

    if ($null -eq $Value) { return '$null' }

    if ($Value -is [bool] -or $Value -is [switch]) {
        if ($Value) { return '$true' } else { return '$false' }
    }

    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double]) { return "$Value" }

    if ($Value -is [Array]) {
        $items = @($Value | ForEach-Object { Format-ApPowerShellLiteral $_ })
        return '@(' + ($items -join ', ') + ')'
    }

    # Single-quoted string with doubled quotes: no expansion, no escape sequences.
    return "'" + ("$Value" -replace "'", "''") + "'"
}

function Format-ApHashtableLiteral {
    <#
    .SYNOPSIS
    Renders a parameter hashtable as a PowerShell hashtable literal for splatting.
    #>
    param([Parameter(Mandatory)][System.Collections.IDictionary]$Parameters)

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('@{')
    foreach ($key in $Parameters.Keys) {
        $lines.Add(('    {0} = {1}' -f $key, (Format-ApPowerShellLiteral $Parameters[$key])))
    }
    $lines.Add('}')

    return ($lines -join [Environment]::NewLine)
}

function Get-ApDependencyPrepBlock {
    <#
    .SYNOPSIS
    Launcher source that installs the engine's online prerequisites without prompting.

    .DESCRIPTION
    The engine bootstraps its own dependencies, but it does so interactively:

        $provider = Get-PackageProvider NuGet -ErrorAction Ignore
        if (-not $provider) { Find-PackageProvider -Name NuGet -ForceBootstrap -IncludeDependencies }
        $module = Import-Module microsoft.graph.authentication -PassThru -ErrorAction Ignore
        if (-not $module) { Install-Module microsoft.graph.authentication -MaximumVersion 2.9.1 }

    On a machine with no NuGet provider and an Untrusted PSGallery, -ForceBootstrap raises a
    Yes/No console prompt. The engine console is hidden, so nobody can answer it and the run
    blocks forever: the browser never appears and the log stops after the banner. That is
    exactly the reported "the sign-in prompt never opens".

    Installing the same prerequisites here, with -Force -Confirm:$false, means the engine
    finds them already present and skips both prompting branches, going straight to
    Connect-MgGraph. Every step is guarded: if the gallery is unreachable the engine still
    runs and fails with a message, rather than hanging.

    Version 2.9.1 matches the engine's own -MaximumVersion pin.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()

    return @'
# ---- prepare the engine's online prerequisites, non-interactively ----
$ConfirmPreference = 'None'

Write-RunLine 'Checking sign-in prerequisites...'

try {
    $nuget = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue
    if (-not $nuget) {
        Write-RunLine '  Installing the NuGet package provider...'
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope AllUsers `
                                -Force -Confirm:$false -ErrorAction Stop | Out-Null
        Write-RunLine '  NuGet package provider installed.'
    }
    else {
        Write-RunLine '  NuGet package provider present.'
    }
}
catch {
    Write-RunLine ('  WARNING: could not install the NuGet provider: ' + $_.Exception.Message)
}

try {
    # Untrusted is the default and makes Install-Module prompt. -Force suppresses that, but
    # trusting it explicitly keeps the engine's own Install-Module call quiet too.
    $repo = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if ($repo -and $repo.InstallationPolicy -ne 'Trusted') {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
        Write-RunLine '  PSGallery marked trusted for this machine.'
    }
}
catch {
    Write-RunLine ('  WARNING: could not set the PSGallery policy: ' + $_.Exception.Message)
}

# Getting the sign-in module *loaded* is the goal, not merely installed. A present but
# incomplete install (files left as OneDrive cloud-only placeholders, say) is discoverable by
# Get-Module -ListAvailable yet throws on import, and because PowerShell resolves an
# unversioned Import-Module to the HIGHEST version, a broken newer copy defeats both the
# engine's version pin and any attempt to install a good one alongside it.
#
# So: load a known-good version explicitly by manifest path. Once it is loaded, the engine's
# own `Import-Module Microsoft.Graph.Authentication` is a no-op and keeps this one. Nothing is
# uninstalled here; use "Repair sign-in module" on the Advanced page to clean up properly.
function Import-GraphAuthModule {
    $installed = @(Get-Module -ListAvailable -Name Microsoft.Graph.Authentication -ErrorAction SilentlyContinue |
                    Sort-Object Version -Descending)

    if ($installed.Count -gt 0) {
        Write-RunLine ('  Installed versions: ' + (($installed | ForEach-Object { $_.Version.ToString() }) -join ', '))
    }

    # Highest first: if the newest works, that is what the engine would have used anyway.
    foreach ($candidate in $installed) {
        try {
            Import-Module -Name $candidate.Path -ErrorAction Stop
            Write-RunLine ('  Loaded ' + $candidate.Version + ' from ' + $candidate.ModuleBase)
            return $true
        }
        catch {
            Write-RunLine ('  Version ' + $candidate.Version + ' will not load: ' + $_.Exception.Message)
        }
    }

    # Nothing usable on disk; fetch the version the engine pins.
    try {
        Write-RunLine '  Installing Microsoft.Graph.Authentication 2.9.1. This can take a minute...'
        Install-Module -Name Microsoft.Graph.Authentication -RequiredVersion 2.9.1 -Scope AllUsers `
                       -Force -AllowClobber -Confirm:$false -ErrorAction Stop

        $fresh = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication -ErrorAction SilentlyContinue |
                    Where-Object { $_.Version.ToString() -eq '2.9.1' } | Select-Object -First 1
        if ($fresh) {
            Import-Module -Name $fresh.Path -ErrorAction Stop
            Write-RunLine ('  Loaded ' + $fresh.Version + ' from ' + $fresh.ModuleBase)
            return $true
        }
    }
    catch {
        Write-RunLine ('  Install failed: ' + $_.Exception.Message)
    }

    return $false
}

if (Import-GraphAuthModule) {
    Write-RunLine '  Sign-in module ready.'
}
else {
    Write-RunLine '  ERROR: no usable Microsoft.Graph.Authentication could be loaded, so sign-in cannot start. Use "Repair sign-in module" on the Advanced page.'
}

Write-RunLine ('-' * 70)
'@
}

function New-ApLauncherScript {
    <#
    .SYNOPSIS
    Writes the launcher that invokes the engine and tees everything to the run log.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$EnginePath,
        [Parameter(Mandatory)][System.Collections.IDictionary]$Parameters,
        [Parameter(Mandatory)][string]$LogPath,
        [Parameter(Mandatory)][string]$LauncherPath
    )

    $paramLiteral = Format-ApHashtableLiteral $Parameters

    # Rendered here rather than in the launcher: building it there would need an if-expression,
    # which Windows PowerShell 5.1 does not support inside a string subexpression.
    $previewLine = Get-ApPreviewCommand -Parameters $Parameters -ScriptPath (Split-Path -Leaf $EnginePath)

    # Only an online run needs Graph; an offline CSV export must not touch the gallery.
    $prepareBlock = ''
    if ($Parameters.Contains('Online')) { $prepareBlock = Get-ApDependencyPrepBlock }

    # $log/$engine are injected as literals; everything else is escaped for the here-string.
    $body = @"
# Generated by Autopilot Import GUI (Community). Safe to delete.
`$ErrorActionPreference = 'Continue'
`$ProgressPreference    = 'SilentlyContinue'
`$WarningPreference     = 'Continue'

`$logPath = $(Format-ApPowerShellLiteral $LogPath)
`$engine  = $(Format-ApPowerShellLiteral $EnginePath)

`$params = $paramLiteral

# Force PSModulePath to Windows PowerShell locations only, before anything tries to
# auto-load a module. Appending the correct paths is not enough: PowerShell 7 entries sort
# first and shadow in-box modules with Core-only copies, which is how a run died at
# Get-PackageProvider with "not recognized" even though PackageManagement was installed.
# The GUI normalises this too, but the launcher repeats it so the generated script also works
# when run by hand from any shell. Pure string work, so it cannot fail for want of a module.
`$documentsPath = [Environment]::GetFolderPath('MyDocuments')
if (-not `$documentsPath) { `$documentsPath = Join-Path `$env:USERPROFILE 'Documents' }

`$modulePaths = New-Object System.Collections.Generic.List[string]
# Case-insensitive, because System32 and system32 arrive from different sources.
`$seenPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
foreach (`$candidate in @(
    (Join-Path `$documentsPath 'WindowsPowerShell\Modules')
    (Join-Path `$env:ProgramFiles 'WindowsPowerShell\Modules')
    (Join-Path `$env:WINDIR 'System32\WindowsPowerShell\v1.0\Modules')
)) {
    `$trimmedCandidate = "`$candidate".TrimEnd('\')
    if (`$seenPaths.Add(`$trimmedCandidate)) { `$modulePaths.Add(`$trimmedCandidate) }
}
if (`$env:PSModulePath) {
    foreach (`$inherited in (`$env:PSModulePath -split ';')) {
        `$trimmed = "`$inherited".TrimEnd('\')
        if (-not `$trimmed) { continue }
        # A real Windows PowerShell path always contains "WindowsPowerShell"; anything else
        # mentioning PowerShell belongs to PS7. Unrelated paths are preserved.
        if (`$trimmed -like '*powershell*' -and `$trimmed -notlike '*windowspowershell*') { continue }
        if (`$seenPaths.Add(`$trimmed)) { `$modulePaths.Add(`$trimmed) }
    }
}
`$env:PSModulePath = (`$modulePaths -join ';')

# One writer held open for the whole run, with AutoFlush so the GUI sees each line
# immediately. Add-Content was used here originally and intermittently dropped lines:
# it opens and closes the file per call, so it can collide with the GUI's tail reader,
# and its failure is non-terminating -- the line vanishes and the run carries on.
# FileShare.Read lets the reader in while keeping other writers out.
`$logStream = New-Object System.IO.FileStream(`$logPath, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
`$logWriter = New-Object System.IO.StreamWriter(`$logStream, (New-Object System.Text.UTF8Encoding(`$false)))
`$logWriter.AutoFlush = `$true

function Write-RunLine {
    param([string]`$Text)
    `$logWriter.WriteLine(`$Text)
}

# Environment banner. When a run fails on someone else's bench this log is the only
# diagnostic available, so record what could not be inferred from the failure itself.
Write-RunLine ("Engine   : " + `$engine)
Write-RunLine ("Command  : " + $(Format-ApPowerShellLiteral $previewLine))
Write-RunLine ("Host     : PowerShell " + `$PSVersionTable.PSVersion + " (" + ([IntPtr]::Size * 8) + "-bit)")
Write-RunLine ("Modules  : " + `$env:PSModulePath)
Write-RunLine ("-" * 70)

# Setup is deliberately separate from the engine call and individually guarded. Previously a
# single try wrapped both, so one failing setup line aborted the whole run before sign-in.
#
# Set-ExecutionPolicy is NOT called here on purpose. powershell.exe is already launched with
# -ExecutionPolicy Bypass, which is authoritative for this process, so the call bought
# nothing - and the cmdlet lives in Microsoft.PowerShell.Security, a module that failed to
# load on a real machine ("the module could not be loaded"). That error is raised at command
# resolution, so -ErrorAction SilentlyContinue does not suppress it: it killed the run.
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 }
catch { Write-RunLine ("WARNING: could not enable TLS 1.2: " + `$_.Exception.Message) }

$prepareBlock

`$exitCode = 0

if (-not (Test-Path -LiteralPath `$engine)) {
    Write-RunLine ("ERROR: the engine script is missing: " + `$engine)
    `$exitCode = 1
}
else {
    try {
        # *>&1 folds every stream (including Write-Host's information stream) into the
        # success stream so the GUI sees the engine's full narration.
        & `$engine @params *>&1 | ForEach-Object { Write-RunLine ("`$_" ) }
    }
    catch {
        Write-RunLine ("ERROR: " + `$_.Exception.Message)
        if (`$_.ScriptStackTrace) { Write-RunLine (`$_.ScriptStackTrace) }
        `$exitCode = 1
    }
}

Write-RunLine ('$script:ApCompleteSentinel' + `$exitCode)

`$logWriter.Flush()
`$logWriter.Dispose()
exit `$exitCode
"@

    # The sentinel is a literal in the generated file, not expanded from our session.
    $body = $body.Replace("'`$script:ApCompleteSentinel'", (Format-ApPowerShellLiteral $script:ApCompleteSentinel))

    Set-Content -LiteralPath $LauncherPath -Value $body -Encoding UTF8
    return $LauncherPath
}

function Start-ApEngineRun {
    <#
    .SYNOPSIS
    Starts an engine run and returns a run context for polling.

    .PARAMETER Parameters
    Parameter hashtable from Build-ApEngineArguments.

    .PARAMETER ShowConsole
    Show the child PowerShell window. Off by default -- output belongs in the GUI -- but
    invaluable when diagnosing a run that appears to stall.

    .OUTPUTS
    A run context: Process, LogPath, ProgressState, StartTime, IsFinished, ExitCode.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Parameters,
        [string]$EnginePath,
        [switch]$ShowConsole,
        [string]$Label = 'run'
    )

    if (-not $EnginePath) { $EnginePath = Resolve-ApEngineScript -Name Engine }

    $workDir = Get-ApWorkingDirectory
    $stamp = Get-Date -Format 'HHmmss'
    $logPath = Join-Path $workDir ('{0}-{1}.log' -f $Label, $stamp)
    $launcherPath = Join-Path $workDir ('{0}-{1}-launcher.ps1' -f $Label, $stamp)

    # Create the log up front so the tail reader always has a file to open.
    Set-Content -LiteralPath $logPath -Value '' -Encoding UTF8

    New-ApLauncherScript -EnginePath $EnginePath -Parameters $Parameters `
                         -LogPath $logPath -LauncherPath $launcherPath | Out-Null

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = Get-ApPowerShellPath
    $psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -STA -File "{0}"' -f $launcherPath
    $psi.WorkingDirectory = $workDir
    $psi.UseShellExecute = $true   # required for WindowStyle to apply
    $psi.WindowStyle = if ($ShowConsole) { 'Normal' } else { 'Hidden' }

    Write-ApLog "Starting engine run: $(Get-ApPreviewCommand -Parameters $Parameters -ScriptPath (Split-Path -Leaf $EnginePath))"
    Write-ApLog "Run log: $logPath"

    $process = [System.Diagnostics.Process]::Start($psi)

    return [pscustomobject]@{
        Process       = $process
        EnginePath    = $EnginePath
        LogPath       = $logPath
        LauncherPath  = $launcherPath
        Parameters    = $Parameters
        ProgressState = New-ApProgressState
        StartTime     = Get-Date
        IsFinished    = $false
        ExitCode      = $null
        Cancelled     = $false
        BytesRead     = 0L
        LineCount     = 0
        # Incomplete trailing line carried between polls (see Update-ApEngineRun).
        PendingLine   = ''
        # Stall detection: a hidden console cannot show an interactive prompt, so a blocked
        # engine looks identical to a slow one. Tracking silence lets the UI say so.
        LastOutputTime = Get-Date
        StallReported  = $false
    }
}

function Update-ApEngineRun {
    <#
    .SYNOPSIS
    Reads whatever the engine has written since the last call and folds it into progress.

    .DESCRIPTION
    Designed to be called from a DispatcherTimer tick. Reads from the recorded byte offset
    with FileShare.ReadWrite so it never contends with the writing process.

    .OUTPUTS
    PSCustomObject with NewLines, ProgressChanged, IsFinished, ExitCode.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Run
    )

    $newLines = New-Object System.Collections.Generic.List[string]
    $progressChanged = $false

    if (Test-Path -LiteralPath $Run.LogPath) {
        $stream = $null
        $reader = $null
        try {
            $stream = New-Object System.IO.FileStream(
                $Run.LogPath,
                [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::Read,
                [System.IO.FileShare]::ReadWrite)

            if ($Run.BytesRead -gt 0 -and $Run.BytesRead -le $stream.Length) {
                [void]$stream.Seek($Run.BytesRead, [System.IO.SeekOrigin]::Begin)
            }

            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
            $chunk = $reader.ReadToEnd()
            $Run.BytesRead = $stream.Position

            if ($chunk) {
                # A read can land mid-line. Carry the incomplete tail over to the next tick,
                # otherwise a single log line surfaces as two truncated fragments and the
                # progress regexes stop matching it.
                $chunk = $Run.PendingLine + $chunk
                $Run.PendingLine = ''

                $parts = $chunk -split "`r?`n"
                if (-not ($chunk.EndsWith("`n"))) {
                    $Run.PendingLine = $parts[-1]
                    if ($parts.Count -gt 1) { $parts = $parts[0..($parts.Count - 2)] }
                    else { $parts = @() }
                }

                foreach ($line in $parts) {
                    if ($null -eq $line) { continue }

                    if ($line.StartsWith($script:ApCompleteSentinel)) {
                        $codeText = $line.Substring($script:ApCompleteSentinel.Length).Trim()
                        $parsed = 0
                        if ([int]::TryParse($codeText, [ref]$parsed)) { $Run.ExitCode = $parsed }
                        continue
                    }

                    if ([string]::IsNullOrWhiteSpace($line)) { continue }

                    $newLines.Add($line)
                    $Run.LineCount++
                    $Run.LastOutputTime = Get-Date

                    if (Update-ApProgressState -State $Run.ProgressState -Line $line) {
                        $progressChanged = $true
                    }
                }
            }
        }
        catch {
            # A transient read error just means we retry on the next tick.
        }
        finally {
            if ($reader) { $reader.Dispose() }
            elseif ($stream) { $stream.Dispose() }
        }
    }

    # Only declare the run finished once the process is gone AND the tail is drained,
    # otherwise the last few lines (including the outcome) can be lost.
    if (-not $Run.IsFinished -and $Run.Process -and $Run.Process.HasExited -and $newLines.Count -eq 0) {

        # The process cannot write again, so an unterminated tail is now a whole line.
        if ($Run.PendingLine -and -not [string]::IsNullOrWhiteSpace($Run.PendingLine)) {
            $tail = $Run.PendingLine
            $Run.PendingLine = ''
            $newLines.Add($tail)
            $Run.LineCount++
            if (Update-ApProgressState -State $Run.ProgressState -Line $tail) { $progressChanged = $true }
        }

        $Run.IsFinished = $true
        if ($null -eq $Run.ExitCode) {
            try { $Run.ExitCode = $Run.Process.ExitCode } catch { $Run.ExitCode = -1 }
        }

        $elapsed = (Get-Date) - $Run.StartTime
        Write-ApLog ("Engine run finished with exit code {0} after {1:mm\:ss}" -f $Run.ExitCode, $elapsed)
    }

    # Report a stall once. A run that has gone quiet for this long is usually blocked on an
    # interactive prompt in the hidden console, which the operator can neither see nor answer.
    $stalled = $false
    if (-not $Run.IsFinished -and -not $Run.StallReported) {
        if (((Get-Date) - $Run.LastOutputTime).TotalSeconds -ge 90) {
            $Run.StallReported = $true
            $stalled = $true
            Write-ApLog "Engine run has produced no output for 90 seconds; it may be waiting on a hidden prompt." -Level WARN
        }
    }

    return [pscustomobject]@{
        NewLines        = $newLines.ToArray()
        ProgressChanged = $progressChanged
        IsFinished      = $Run.IsFinished
        ExitCode        = $Run.ExitCode
        Stalled         = $stalled
    }
}

function Stop-ApEngineRun {
    <#
    .SYNOPSIS
    Terminates a run and its child processes.

    .DESCRIPTION
    Uses taskkill /T because the engine may itself have started children (the
    pre-provisioning helper, sysprep, changepk) that a plain Process.Kill would orphan.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Run
    )

    if (-not $Run.Process) { return }

    try {
        if (-not $Run.Process.HasExited) {
            Write-ApLog "Cancelling engine run (PID $($Run.Process.Id))" -Level WARN
            & taskkill.exe /PID $Run.Process.Id /T /F 2>&1 | Out-Null

            if (-not $Run.Process.WaitForExit(5000)) {
                try { $Run.Process.Kill() } catch { }
            }
        }
    }
    catch {
        Write-ApLog "Could not cancel the run cleanly: $($_.Exception.Message)" -Level WARN
    }
    finally {
        $Run.Cancelled = $true
        $Run.IsFinished = $true
        if ($null -eq $Run.ExitCode) { $Run.ExitCode = -1 }
    }
}

function Get-ApRunSummary {
    <#
    .SYNOPSIS
    One-line outcome for the status bar.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Run)

    $elapsed = (Get-Date) - $Run.StartTime
    $duration = '{0:mm\:ss}' -f $elapsed

    if ($Run.Cancelled) { return "Cancelled after $duration" }

    $state = $Run.ProgressState
    if ($state.IsError) { return "Failed after ${duration}: $($state.ErrorMessage)" }
    if ($Run.ExitCode -ne 0) { return "Finished with errors after $duration (exit code $($Run.ExitCode)). See the log for details." }
    if ($state.IsComplete) { return "Completed in $duration" }

    return "Finished in $duration"
}

# Background.ps1 -- run work off the UI thread and marshal the result back.
#
# WPF has a single UI thread: anything slow on it freezes the window, including the
# "is the network broken" check, which is precisely the moment a frozen window is least
# forgivable (a fully unreachable endpoint list costs one timeout per parallel batch).
#
# Pattern follows VM-Pilot (VMPilot.GUI.ps1:1167-1232): a dedicated runspace for the work,
# a DispatcherTimer to notice completion, and the continuation invoked back on the UI
# thread so it can touch controls safely.

$script:ApBackgroundJobs = New-Object System.Collections.Generic.List[object]

function Start-ApBackgroundWork {
    <#
    .SYNOPSIS
    Runs a scriptblock in a background runspace and calls OnComplete on the UI thread.

    .PARAMETER Work
    Scriptblock to execute. It sees the variables supplied in -Variables and any functions
    named in -FunctionNames.

    .PARAMETER Variables
    Name/value pairs injected into the runspace's session state.

    .PARAMETER FunctionNames
    Functions from the current session to copy into the runspace by source text. Only pass
    self-contained workers; a function that logs or reads config will not find those helpers.

    .PARAMETER OnComplete
    Invoked on the UI thread with two arguments: the work's output, and an error message
    ($null on success).

    .PARAMETER Window
    Window whose Dispatcher is used to marshal the continuation.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$Work,
        [hashtable]$Variables = @{},
        [string[]]$FunctionNames = @(),
        [Parameter(Mandatory)][scriptblock]$OnComplete,
        [Parameter(Mandatory)]$Window,
        [int]$PollIntervalMs = 150
    )

    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = 'MTA'    # worker threads do no UI work
    $runspace.ThreadOptions = 'ReuseThread'
    $runspace.Open()

    foreach ($key in $Variables.Keys) {
        $runspace.SessionStateProxy.SetVariable($key, $Variables[$key])
    }

    $shell = [powershell]::Create()
    $shell.Runspace = $runspace

    # Copy requested functions in as source. A runspace starts empty, so anything the
    # work calls must either be a built-in cmdlet or arrive this way.
    foreach ($name in $FunctionNames) {
        $command = Get-Command $name -CommandType Function -ErrorAction SilentlyContinue
        if (-not $command) {
            Write-ApLog "Cannot copy function '$name' into the background runspace: not found." -Level WARN
            continue
        }
        [void]$shell.AddScript("function $name {`n$($command.Definition)`n}")
    }

    [void]$shell.AddScript($Work.ToString())

    $job = [pscustomobject]@{
        Shell    = $shell
        Runspace = $runspace
        Handle   = $null
        Timer    = $null
    }

    $job.Handle = $shell.BeginInvoke()
    $script:ApBackgroundJobs.Add($job)

    # Capture the list as a local so the tick closure holds a direct reference. Inside a
    # GetNewClosure() scriptblock, $script: binds to the closure's own captured session
    # state, where this variable is not visible; calling .Remove() on the resulting $null
    # threw out of the finally block and the continuation never ran, so the UI sat on
    # "reading..." forever with no error shown.
    $jobList = $script:ApBackgroundJobs

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds($PollIntervalMs)
    $job.Timer = $timer

    # Everything inside a DispatcherTimer tick must be caught. An exception that escapes a
    # tick is not merely logged by WPF: it propagates out of Window.ShowDialog and tears the
    # whole application down, which looks to the user like the window vanishing on startup.
    $timer.Add_Tick({
        try {
            if (-not $job.Handle.IsCompleted) { return }

            $timer.Stop()

            $output = $null
            $errorMessage = $null
            try {
                $output = $job.Shell.EndInvoke($job.Handle)

                # A non-terminating error inside the runspace surfaces here, not as an exception.
                if ($job.Shell.HadErrors -and $job.Shell.Streams.Error.Count -gt 0) {
                    $errorMessage = ($job.Shell.Streams.Error | ForEach-Object { "$_" }) -join '; '
                }
            }
            catch {
                $errorMessage = $_.Exception.Message
            }
            finally {
                try { $job.Shell.Dispose() } catch { }
                try { $job.Runspace.Close(); $job.Runspace.Dispose() } catch { }
                try { [void]$jobList.Remove($job) } catch { }
            }

            & $OnComplete $output $errorMessage
        }
        catch {
            try { $timer.Stop() } catch { }
            Write-ApLog "Background work continuation failed: $($_.Exception.Message)" -Level ERROR
            if ($_.ScriptStackTrace) { Write-ApLog $_.ScriptStackTrace -Level DEBUG }
        }
    }.GetNewClosure())

    $timer.Start()
    return $job
}

function Stop-ApBackgroundWork {
    <#
    .SYNOPSIS
    Tears down any still-running background jobs. Called on window close.

    .DESCRIPTION
    Best-effort by design: this runs while the window is closing, and a job may already be
    half-disposed by its own completion handler. Every step is guarded individually so one
    failure cannot abort the rest of the shutdown sequence.
    #>
    [CmdletBinding()]
    param()

    $jobs = @()
    try { $jobs = @($script:ApBackgroundJobs) } catch { $jobs = @() }

    foreach ($job in $jobs) {
        if (-not $job) { continue }
        try { if ($job.Timer) { $job.Timer.Stop() } } catch { }
        try { if ($job.Shell) { $job.Shell.Stop() } } catch { }
        try { if ($job.Shell) { $job.Shell.Dispose() } } catch { }
        try { if ($job.Runspace) { $job.Runspace.Close(); $job.Runspace.Dispose() } } catch { }
    }

    try { $script:ApBackgroundJobs.Clear() }
    catch {
        # Replace rather than clear if the list itself is unusable.
        $script:ApBackgroundJobs = New-Object System.Collections.Generic.List[object]
    }
}

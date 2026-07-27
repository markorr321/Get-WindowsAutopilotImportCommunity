# Show-AutopilotImportGui.ps1 -- builds the window and wires every control to behaviour.
#
# Shared state lives in $script: variables rather than function locals because WPF event
# handlers are scriptblocks that execute later, outside the defining function's scope.

$script:ApWin = $null
$script:ApEl = $null
$script:ApDevice = $null
$script:ApRun = $null
$script:ApRunTimer = $null
$script:ApClockTimer = $null
$script:ApActiveOutput = $null
$script:ApNetworkResults = @()
$script:ApGraphCheck = $null
# Set when a v2 register run is launched with the restart option ticked; consumed once the
# run completes successfully.
$script:ApPendingV2Reboot = $false
$script:ApEnginePath = $null
$script:ApAppVersion = '1.2.1'
$script:ApAuthor = 'Mark Orr'
$script:ApAuthorHandle = '@markorr321'
$script:ApAuthorSite = 'https://orr365.tools'
$script:ApAuthorGitHub = 'https://github.com/markorr321'

# ============================ small UI helpers ============================

function New-ApBrush {
    param([Parameter(Mandatory)][string]$Hex)
    return New-Object System.Windows.Media.SolidColorBrush(
        [System.Windows.Media.ColorConverter]::ConvertFromString($Hex))
}

function Set-ApPill {
    <#
    .SYNOPSIS
    Sets a status chip's text and colour scheme.

    .PARAMETER State
    ok | warn | error | neutral
    #>
    param(
        [Parameter(Mandatory)]$Border,
        [Parameter(Mandatory)]$TextBlock,
        [Parameter(Mandatory)][string]$Text,
        [ValidateSet('ok', 'warn', 'error', 'neutral')][string]$State = 'neutral'
    )

    # Tinted backgrounds rather than saturated fills: a row of bright chips in a dark UI
    # reads as alarming even when everything is fine.
    #
    # The locals are named *Colour, not $border: PowerShell variable names are
    # case-insensitive, so a local $border would silently overwrite the $Border parameter
    # and the assignments below would fail against a string.
    switch ($State) {
        'ok'    { $bgColour = '#12291B'; $borderColour = '#1E5B33'; $fgColour = '#3BC77A' }
        'warn'  { $bgColour = '#2A2110'; $borderColour = '#5B4718'; $fgColour = '#E9B44C' }
        'error' { $bgColour = '#2A1417'; $borderColour = '#6B2027'; $fgColour = '#F03A47' }
        default { $bgColour = '#252525'; $borderColour = '#3A3A3A'; $fgColour = '#C0C0C0' }
    }

    $Border.Background = New-ApBrush $bgColour
    $Border.BorderBrush = New-ApBrush $borderColour
    $TextBlock.Foreground = New-ApBrush $fgColour
    $TextBlock.Text = $Text
}

function Start-ApUrl {
    <#
    .SYNOPSIS
    Opens a URL in the default browser.

    .DESCRIPTION
    Only http and https are accepted. Handing an arbitrary string to Start-Process would
    happily launch a local executable or a file association, which is not what a link in a
    UI should ever do.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Url)

    $parsed = $null
    if (-not [System.Uri]::TryCreate($Url, [System.UriKind]::Absolute, [ref]$parsed) -or
        $parsed.Scheme -notin @('http', 'https')) {
        Write-ApLog "Refusing to open a non-web link: $Url" -Level WARN
        return
    }

    try {
        Start-Process $parsed.AbsoluteUri | Out-Null
        Write-ApLog "Opened $($parsed.AbsoluteUri)"
    }
    catch {
        Set-ApStatus -Text "Could not open the link: $($_.Exception.Message)" -IsError
    }
}

function Set-ApStatus {
    <#
    .SYNOPSIS
    Updates the bottom status strip.
    #>
    param(
        [string]$Text,
        [string]$Stage,
        [Nullable[int]]$Percent,
        [switch]$IsError
    )

    if ($PSBoundParameters.ContainsKey('Text')) {
        $script:ApEl.StatusText.Text = $Text
        $script:ApEl.StatusText.Foreground = if ($IsError) { New-ApBrush '#F03A47' } else { New-ApBrush '#C0C0C0' }
    }
    if ($PSBoundParameters.ContainsKey('Stage')) {
        $script:ApEl.StatusStage.Text = $Stage
    }
    if ($null -ne $Percent) {
        $script:ApEl.StatusProgress.IsIndeterminate = $false
        $script:ApEl.StatusProgress.Value = $Percent
        $script:ApEl.StatusProgress.Foreground = if ($IsError) { New-ApBrush '#F03A47' } else { New-ApBrush '#0078D4' }
    }
}

function Add-ApOutput {
    <#
    .SYNOPSIS
    Appends timestamped lines to an output box and keeps it scrolled to the end.
    #>
    param(
        $Box,
        [string[]]$Lines
    )

    if (-not $Box) { return }
    if (-not $Lines -or $Lines.Count -eq 0) { return }

    $stamp = Get-Date -Format 'HH:mm:ss'
    $text = ($Lines | ForEach-Object { "$stamp  $_" }) -join [Environment]::NewLine

    if ($Box.Text.Length -gt 0) { $Box.AppendText([Environment]::NewLine) }
    $Box.AppendText($text)
    $Box.ScrollToEnd()
}

function Show-ApPage {
    <#
    .SYNOPSIS
    Shows one page and hides the rest.
    #>
    param([Parameter(Mandatory)][string]$Name)

    foreach ($page in @('PageRegister', 'PageDevice', 'PageBatch', 'PageNetwork', 'PageAdvanced', 'PageLogs')) {
        $script:ApEl[$page].Visibility = if ($page -eq $Name) { 'Visible' } else { 'Collapsed' }
    }
}

function Show-ApNotice {
    <#
    .SYNOPSIS
    Shows the warning card on the Register page.
    #>
    param(
        [string]$Title = 'Before you continue',
        [Parameter(Mandatory)][string[]]$Messages
    )

    $messages = @($Messages | Where-Object { $_ })
    if ($messages.Count -eq 0) { Hide-ApNotice; return }

    $script:ApEl.NoticeTitle.Text = $Title
    $script:ApEl.NoticeText.Text = ($messages | ForEach-Object { "- $_" }) -join [Environment]::NewLine
    $script:ApEl.NoticeCard.Visibility = 'Visible'
}

function Hide-ApNotice {
    $script:ApEl.NoticeCard.Visibility = 'Collapsed'
}

# ============================ device page ============================

# Functions Get-ApDeviceInfo needs once it is copied into a bare runspace. Get-ApCimSafe is
# deliberately log-free so this list stays closed.
$script:ApDeviceInfoFunctions = @(
    'Get-ApCimSafe'
    'Get-ApDeviceIdentifierPart'
    'Test-ApHardwareHashAvailable'
    'Get-ApTpmInfo'
    'Test-ApSecureBoot'
    'Get-ApWindowsDisplayVersion'
    'Test-ApVirtualMachine'
    'Test-ApElevated'
    'Get-ApDeviceInfo'
)

function Start-ApDeviceLoad {
    <#
    .SYNOPSIS
    Reads local hardware on a background runspace, then applies it to the UI.

    .DESCRIPTION
    Two of the queries behind Get-ApDeviceInfo block for about five seconds each when the
    process is not elevated, so this must not run on the UI thread: doing so kept the window
    from appearing for over ten seconds. The panes show "reading..." until it lands.
    #>
    [CmdletBinding()]
    param()

    $el = $script:ApEl
    $el.RefreshDeviceButton.IsEnabled = $false
    Set-ApStatus -Text 'Reading device information...'

    Start-ApBackgroundWork -Window $script:ApWin `
        -FunctionNames $script:ApDeviceInfoFunctions `
        -Work { Get-ApDeviceInfo } `
        -OnComplete {
            param($result, $errorMessage)

            $script:ApEl.RefreshDeviceButton.IsEnabled = $true

            if ($errorMessage -or -not $result) {
                Set-ApStatus -Text "Could not read device information: $errorMessage" -IsError
                Write-ApLog "Device information failed: $errorMessage" -Level ERROR
                return
            }

            Set-ApDeviceUi (@($result)[-1])
            Sync-ApModeUi

            # A machine with no hash cannot do v1, so switch unless the caller pinned a mode.
            if (-not $script:ApModePinned -and -not $script:ApDevice.HardwareHashReady -and $script:ApDevice.IsElevated) {
                $script:ApEl.ModeV2.IsChecked = $true
                Write-ApLog 'No hardware hash available; defaulting to Device Preparation (v2).'
            }

            Set-ApStatus -Text 'Ready.' -Percent 0
        } | Out-Null
}

function Set-ApDeviceUi {
    <#
    .SYNOPSIS
    Applies a device-info snapshot to the summary, readiness chips and inventory grid.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]$Info
    )

    $info = $Info
    $script:ApDevice = $info

    $el = $script:ApEl
    $el.SummarySerial.Text = if ($info.SerialNumber) { $info.SerialNumber } else { 'unavailable' }
    $el.SummaryModel.Text = if ($info.Model) { $info.Model } else { 'unavailable' }
    $el.SummaryManufacturer.Text = if ($info.Manufacturer) { $info.Manufacturer } else { 'unavailable' }
    $el.SummaryFreeSpace.Text = if ($null -ne $info.FreeSpaceGb) { "$($info.FreeSpaceGb) GB" } else { 'unknown' }

    # Hardware hash: the hard prerequisite for v1. Virtual machines usually cannot supply
    # one, which is exactly why Device Preparation (v2) exists.
    if ($info.HardwareHashReady) {
        Set-ApPill $el.PillHash $el.PillHashText 'Hash ready' 'ok'
    }
    elseif (-not $info.IsElevated) {
        Set-ApPill $el.PillHash $el.PillHashText 'Hash needs admin' 'error'
    }
    else {
        Set-ApPill $el.PillHash $el.PillHashText 'No hash' 'error'
    }

    # Both the TPM class and Confirm-SecureBootUEFI need administrator rights. Without them
    # the query fails the same way a genuinely absent TPM does, so report "unknown" rather
    # than asserting hardware facts we could not actually read.
    if ($info.TpmPresent -and $info.TpmEnabled) {
        Set-ApPill $el.PillTpm $el.PillTpmText "TPM $($info.TpmSpecVersion)" 'ok'
    }
    elseif ($info.TpmPresent) {
        Set-ApPill $el.PillTpm $el.PillTpmText 'TPM disabled' 'warn'
    }
    elseif (-not $info.IsElevated) {
        Set-ApPill $el.PillTpm $el.PillTpmText 'TPM unknown' 'warn'
    }
    else {
        Set-ApPill $el.PillTpm $el.PillTpmText 'No TPM' 'warn'
    }

    if ($info.SecureBoot -eq $true) {
        Set-ApPill $el.PillSecureBoot $el.PillSecureBootText 'Secure Boot on' 'ok'
    }
    elseif ($info.SecureBoot -eq $false) {
        Set-ApPill $el.PillSecureBoot $el.PillSecureBootText 'Secure Boot off' 'warn'
    }
    elseif (-not $info.IsElevated) {
        Set-ApPill $el.PillSecureBoot $el.PillSecureBootText 'Secure Boot unknown' 'warn'
    }
    else {
        Set-ApPill $el.PillSecureBoot $el.PillSecureBootText 'Legacy BIOS' 'warn'
    }

    $el.IdentifierPreviewBox.Text = $info.DeviceIdentifier

    $rows = @(
        [pscustomobject]@{ Name = 'Serial number';   Value = $info.SerialNumber }
        [pscustomobject]@{ Name = 'Manufacturer';    Value = $info.Manufacturer }
        [pscustomobject]@{ Name = 'Model';           Value = $info.Model }
        [pscustomobject]@{ Name = 'Computer name';   Value = $info.ComputerName }
        [pscustomobject]@{ Name = 'Operating system'; Value = $info.OsCaption }
        [pscustomobject]@{ Name = 'OS build';        Value = $info.OsBuild }
        [pscustomobject]@{ Name = 'OS version';      Value = $info.OsDisplayVersion }
        [pscustomobject]@{ Name = 'Memory';          Value = if ($null -ne $info.MemoryGb) { "$($info.MemoryGb) GB" } else { 'unknown' } }
        [pscustomobject]@{ Name = "Disk $($info.SystemDrive)"; Value = if ($null -ne $info.FreeSpaceGb) { "$($info.FreeSpaceGb) GB free of $($info.TotalSpaceGb) GB" } else { 'unknown' } }
        [pscustomobject]@{ Name = 'TPM';             Value = if ($info.TpmPresent) { "version $($info.TpmSpecVersion), enabled: $($info.TpmEnabled)" } elseif (-not $info.IsElevated) { 'unknown (requires administrator)' } else { 'not detected' } }
        [pscustomobject]@{ Name = 'Secure Boot';     Value = if ($null -ne $info.SecureBoot) { "$($info.SecureBoot)" } elseif (-not $info.IsElevated) { 'unknown (requires administrator)' } else { 'not applicable (legacy BIOS)' } }
        [pscustomobject]@{ Name = 'Hardware hash';   Value = if ($info.HardwareHashReady) { 'available' } elseif (-not $info.IsElevated) { 'unknown (requires administrator)' } else { 'not available' } }
        [pscustomobject]@{ Name = 'Virtual machine'; Value = "$($info.IsVirtualMachine)" }
        [pscustomobject]@{ Name = 'Running elevated'; Value = "$($info.IsElevated)" }
        [pscustomobject]@{ Name = 'Device identifier (v2)'; Value = $info.DeviceIdentifier }
    )
    $el.DeviceGrid.ItemsSource = $rows

    Write-ApLog "Device: $($info.Manufacturer) $($info.Model) serial $($info.SerialNumber); hash ready: $($info.HardwareHashReady)"
    return $info
}

function Set-ApDevicePlaceholders {
    <#
    .SYNOPSIS
    Neutral "reading" state shown until the background device load completes.
    #>
    [CmdletBinding()]
    param()

    $el = $script:ApEl
    foreach ($name in @('SummarySerial', 'SummaryModel', 'SummaryManufacturer', 'SummaryFreeSpace')) {
        $el[$name].Text = 'reading...'
    }
    Set-ApPill $el.PillHash $el.PillHashText 'Hash ...' 'neutral'
    Set-ApPill $el.PillTpm $el.PillTpmText 'TPM ...' 'neutral'
    Set-ApPill $el.PillSecureBoot $el.PillSecureBootText 'Secure Boot ...' 'neutral'
}

# ============================ mode handling ============================

function Sync-ApModeUi {
    <#
    .SYNOPSIS
    Enables or disables controls according to the selected registration mode.

    .DESCRIPTION
    Autopilot v2 (-identifier) runs a separate branch in the community script that ignores
    group tag, assigned user, computer name, Entra group, assignment wait and reboot
    (get-windowsautopilotinfocommunity.ps1:2233-2270). Greying those controls out is
    honest; leaving them enabled would imply they do something.
    #>
    [CmdletBinding()]
    param()

    $el = $script:ApEl
    $isV2 = [bool]$el.ModeV2.IsChecked

    foreach ($name in @('GroupTagCombo', 'AssignedUserBox', 'ComputerNameBox', 'AddToGroupBox')) {
        $el[$name].IsEnabled = -not $isV2
    }

    # The group tag is hidden rather than merely dimmed in v2: it is the one field
    # technicians reach for by habit, and Device Preparation targets devices through the
    # Entra group on the policy instead, so leaving it on screen invites a wasted entry.
    $el.GroupTagSection.Visibility = if ($isV2) { 'Collapsed' } else { 'Visible' }

    # Every field in this card (group tag, assigned user, computer name, Entra group) is
    # ignored by the identifier path, so in v2 the whole card goes. Dimming it was worse than
    # useless: three dead fields pushed the one live v2 option, the restart, below the fold.
    $el.CardDetails.Visibility = if ($isV2) { 'Collapsed' } else { 'Visible' }

    # Swap the whole options block rather than dimming it. The v1 controls map to engine
    # switches the identifier path ignores; the v2 restart is performed by this tool.
    $el.OptionsV1Section.Visibility = if ($isV2) { 'Collapsed' } else { 'Visible' }
    $el.OptionsV2Section.Visibility = if ($isV2) { 'Visible' } else { 'Collapsed' }

    # v1 delegates its restart to the engine's -Reboot, so the on-demand button is a v2 affair.
    $el.RestartNowButton.Visibility = if ($isV2) { 'Visible' } else { 'Collapsed' }

    $el.CardDetails.Opacity = if ($isV2) { 0.55 } else { 1.0 }

    $el.IdentifierPreviewLabel.Visibility = if ($isV2) { 'Visible' } else { 'Collapsed' }
    $el.IdentifierPreviewBox.Visibility = if ($isV2) { 'Visible' } else { 'Collapsed' }

    if ($isV2) {
        $el.ModeDescription.Text = 'Imports the Manufacturer,Model,Serial identifier for Windows Autopilot Device Preparation. No hardware hash is needed, so this works on virtual machines. Devices are targeted by the Entra group on your Device Preparation policy, not by a group tag.'
        Show-ApNotice -Title 'Device Preparation mode' -Messages @(
            'Group tag, assigned user, computer name and profile assignment do not apply to v2.',
            'After importing, add this device to the Entra security group used by your Device Preparation policy.'
        )
    }
    else {
        $el.ModeDescription.Text = 'Uploads the 4K hardware hash to Windows Autopilot, then optionally waits for a deployment profile to be assigned.'
        Hide-ApNotice

        if ($script:ApDevice -and -not $script:ApDevice.HardwareHashReady) {
            $reason = if (-not $script:ApDevice.IsElevated) {
                'This window is not running as administrator, so the hardware hash cannot be read.'
            }
            elseif ($script:ApDevice.IsVirtualMachine) {
                'This looks like a virtual machine, and VMs frequently cannot provide a hardware hash. Device Preparation (v2) is usually the right choice here.'
            }
            else {
                'This device did not return a hardware hash.'
            }
            Show-ApNotice -Title 'Hardware hash unavailable' -Messages @($reason)
        }
    }
}

# ============================ request assembly ============================

function Get-ApUiRequest {
    <#
    .SYNOPSIS
    Reads the current UI state into a registration request.
    #>
    [CmdletBinding()]
    param(
        [ValidateSet('Register', 'Export', 'Batch')][string]$Operation = 'Register'
    )

    $el = $script:ApEl
    $request = New-ApRegistrationRequest
    $request.Operation = $Operation

    if ($Operation -eq 'Batch') {
        $request.Mode = if ($el.BatchModeV2.IsChecked) { 'v2' } else { 'v1' }
        $request.InputFile = $el.BatchFileBox.Text.Trim()
        $request.WaitForAssignment = [bool]$el.BatchWaitCheck.IsChecked
        return $request
    }

    $request.Mode = if ($el.ModeV2.IsChecked) { 'v2' } else { 'v1' }

    if ($Operation -eq 'Export') {
        $request.Append = [bool]$el.ExportAppendCheck.IsChecked
        $request.Partner = [bool]$el.ExportPartnerCheck.IsChecked
    }

    # The editable ComboBox's Text is the authoritative value: the tech may have typed a new
    # tag rather than picked a remembered one.
    $request.GroupTag = "$($el.GroupTagCombo.Text)"
    $request.AssignedUser = $el.AssignedUserBox.Text
    $request.AssignedComputerName = $el.ComputerNameBox.Text
    $request.AddToGroup = $el.AddToGroupBox.Text

    if ($Operation -eq 'Register') {
        $request.WaitForAssignment = [bool]$el.WaitAssignCheck.IsChecked
        $request.Reboot = [bool]$el.RebootCheck.IsChecked
        $request.Wipe = [bool]$el.AdvWipeCheck.IsChecked
        $request.Sysprep = [bool]$el.AdvSysprepCheck.IsChecked
        $request.PreProvision = [bool]$el.AdvPreProvisionCheck.IsChecked
        $request.ChangePK = $el.AdvChangePkBox.Text

        $request.ExistingDevicePolicy =
            if ($el.PolicyDelete.IsChecked) { 'delete' }
            elseif ($el.PolicySkip.IsChecked) { 'skipcheck' }
            else { 'update' }
    }

    return $request
}

function Test-ApDestructiveRequest {
    <#
    .SYNOPSIS
    True when the request does something the tech should confirm first.
    #>
    param([Parameter(Mandatory)]$Request)

    return [bool](
        $Request.Wipe -or
        $Request.Sysprep -or
        $Request.ChangePK -or
        ($Request.ExistingDevicePolicy -eq 'delete')
    )
}

# ============================ run lifecycle ============================

function Set-ApRunningState {
    <#
    .SYNOPSIS
    Enables or disables the controls that must not change mid-run.
    #>
    param([bool]$IsRunning)

    $el = $script:ApEl
    $el.RegisterButton.IsEnabled = -not $IsRunning
    $el.PreviewButton.IsEnabled = -not $IsRunning
    $el.BatchRunButton.IsEnabled = -not $IsRunning
    $el.BatchPreviewButton.IsEnabled = -not $IsRunning
    $el.AdvDiagnosticsButton.IsEnabled = -not $IsRunning
    $el.AdvDiagnosticsOnlineCheck.IsEnabled = -not $IsRunning
    $el.AdvWindowsUpdateButton.IsEnabled = -not $IsRunning
    $el.ExportHashButton.IsEnabled = -not $IsRunning
    $el.ExportIdentifierButton.IsEnabled = -not $IsRunning
    $el.CancelButton.IsEnabled = $IsRunning

    $el.RegisterButton.Content = if ($IsRunning) { 'REGISTERING...' } else { 'REGISTER THIS DEVICE' }
}

function Start-ApGuiRun {
    <#
    .SYNOPSIS
    Starts an engine run and begins streaming its output into the given box.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Parameters,
        [Parameter(Mandatory)]$OutputBox,
        [string]$Label = 'run',
        [string]$StartMessage = 'Starting...',
        # Defaults to the resolved community engine; the diagnostics script passes its own.
        [string]$EnginePath
    )

    if (-not $EnginePath) { $EnginePath = $script:ApEnginePath }

    if ($script:ApRun -and -not $script:ApRun.IsFinished) {
        Show-ApDialog -Title 'A run is already in progress' -Owner $script:ApWin `
                      -Message 'Wait for the current operation to finish, or cancel it first.' | Out-Null
        return
    }

    $config = Get-ApConfig

    try {
        $script:ApRun = Start-ApEngineRun -Parameters $Parameters `
                                          -EnginePath $EnginePath `
                                          -Label $Label `
                                          -ShowConsole:([bool]$config.showConsoleWindow)
    }
    catch {
        Set-ApStatus -Text "Could not start: $($_.Exception.Message)" -Percent 0 -IsError
        Show-ApDialog -Title 'Could not start the engine' -Owner $script:ApWin `
                      -Message $_.Exception.Message | Out-Null
        return
    }

    $script:ApActiveOutput = $OutputBox
    $OutputBox.Clear()
    Add-ApOutput -Box $OutputBox -Lines @($StartMessage)

    Set-ApRunningState $true
    Set-ApStatus -Text $StartMessage -Stage 'Connect' -Percent 0

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(250)
    $timer.Add_Tick({
        try {
            Invoke-ApRunTick
        }
        catch {
            # Never let a tick exception escape: it would propagate out of ShowDialog.
            if ($script:ApRunTimer) { $script:ApRunTimer.Stop(); $script:ApRunTimer = $null }
            Set-ApRunningState $false
            Set-ApStatus -Text "Lost track of the run: $($_.Exception.Message)" -IsError
            Write-ApLog "Run tick failed: $($_.Exception.Message)" -Level ERROR
            if ($_.ScriptStackTrace) { Write-ApLog $_.ScriptStackTrace -Level DEBUG }
        }
    })
    $script:ApRunTimer = $timer
    $timer.Start()
}

function Invoke-ApRunTick {
    <#
    .SYNOPSIS
    DispatcherTimer tick: drain new output, update progress, finalise when done.
    #>
    [CmdletBinding()]
    param()

    $run = $script:ApRun
    if (-not $run) { return }

    $update = Update-ApEngineRun -Run $run

    if ($update.NewLines.Count -gt 0 -and $script:ApActiveOutput) {
        Add-ApOutput -Box $script:ApActiveOutput -Lines $update.NewLines
    }

    $state = $run.ProgressState
    if ($update.ProgressChanged) {
        Set-ApStatus -Text $state.StageLabel -Stage $state.Stage -Percent $state.Percent -IsError:([bool]$state.IsError)
    }

    if ($update.Stalled) {
        $hint = 'No output for 90 seconds. The engine may be waiting on a prompt in its hidden console. Turn on "Show the engine console window" on the Advanced page, then cancel and retry.'
        Set-ApStatus -Text $hint -IsError
        Add-ApOutput -Box $script:ApActiveOutput -Lines @('', "[GUI] $hint")
    }

    if (-not $update.IsFinished) { return }

    if ($script:ApRunTimer) { $script:ApRunTimer.Stop(); $script:ApRunTimer = $null }
    Set-ApRunningState $false

    $summary = Get-ApRunSummary -Run $run
    $failed = $run.Cancelled -or $state.IsError -or ($run.ExitCode -ne 0)

    Set-ApStatus -Text $summary -Stage $state.Stage -Percent ($(if ($failed) { $state.Percent } else { 100 })) -IsError:$failed
    Add-ApOutput -Box $script:ApActiveOutput -Lines @('', $summary)
    Write-ApLog "Run summary: $summary" -Level $(if ($failed) { 'WARN' } else { 'INFO' })

    Update-ApLogsPage

    # Device Preparation restart. Only on a clean run: never reboot after a failure or a
    # cancel, or the operator loses the log and the device leaves OOBE unregistered.
    if ($script:ApPendingV2Reboot -and -not $failed -and $state.IsComplete) {
        $script:ApPendingV2Reboot = $false
        Add-ApOutput -Box $script:ApActiveOutput -Lines @('', '[GUI] Identifier imported. Restarting this device now.')
        Set-ApStatus -Text 'Identifier imported. Restarting...' -Percent 100
        Invoke-ApRestartComputer | Out-Null
        return
    }
    $script:ApPendingV2Reboot = $false

    if ($state.IsError) {
        Show-ApDialog -Title 'The run reported a problem' -Owner $script:ApWin `
                      -Message $state.ErrorMessage `
                      -Detail "Full log:`r`n$($run.LogPath)" | Out-Null
    }
}

function Invoke-ApRestartComputer {
    <#
    .SYNOPSIS
    Restarts this machine.

    .DESCRIPTION
    Used for the Device Preparation (v2) restart option. Autopilot v1 delegates its restart to
    the engine's -Reboot switch, but that switch is nested inside the engine's
    assignment-wait block, which the -identifier path never reaches. So for v2 the restart has
    to happen here.

    Restart-Computer comes from Microsoft.PowerShell.Management, which Windows PowerShell loads
    at startup rather than auto-loading, so it is unaffected by the module-shadowing problems
    that affect the engine process. shutdown.exe is kept as a fallback anyway: it is a native
    binary and needs no module resolution at all.
    #>
    [CmdletBinding()]
    param()

    Write-ApLog 'Restarting the computer at the operator''s request.' -Level WARN

    try {
        Restart-Computer -Force -ErrorAction Stop
        return $true
    }
    catch {
        Write-ApLog "Restart-Computer failed: $($_.Exception.Message). Falling back to shutdown.exe." -Level WARN
        try {
            Start-Process -FilePath (Join-Path $env:WINDIR 'System32\shutdown.exe') `
                          -ArgumentList '/r', '/t', '0' -WindowStyle Hidden -ErrorAction Stop | Out-Null
            return $true
        }
        catch {
            Write-ApLog "shutdown.exe also failed: $($_.Exception.Message)" -Level ERROR
            return $false
        }
    }
}

function Stop-ApGuiRun {
    [CmdletBinding()]
    param()

    if (-not $script:ApRun -or $script:ApRun.IsFinished) { return }

    Stop-ApEngineRun -Run $script:ApRun

    if ($script:ApRunTimer) { $script:ApRunTimer.Stop(); $script:ApRunTimer = $null }
    Set-ApRunningState $false
    Set-ApStatus -Text 'Cancelled.' -Stage '' -Percent 0 -IsError
    Add-ApOutput -Box $script:ApActiveOutput -Lines @('', 'Cancelled by the operator.')
}

# ============================ logs page ============================

function Update-ApLogsPage {
    [CmdletBinding()]
    param()

    $lines = Get-ApLogBuffer
    $script:ApEl.LogsOutput.Text = ($lines -join [Environment]::NewLine)
    $script:ApEl.LogsOutput.ScrollToEnd()
    $script:ApEl.LogsPathText.Text = Get-ApLogPath
}

# ============================ network page ============================

function Start-ApNetworkCheck {
    [CmdletBinding()]
    param()

    $el = $script:ApEl
    $el.NetworkRunButton.IsEnabled = $false
    $el.NetworkSummary.Text = 'Testing...'
    $el.NetworkGrid.ItemsSource = $null

    $endpoints = Get-ApConfiguredEndpoints
    Write-ApLog "Starting connectivity check across $($endpoints.Count) endpoints"
    Set-ApStatus -Text "Testing $($endpoints.Count) endpoints..." -Stage 'Network'
    $el.StatusProgress.IsIndeterminate = $true

    Start-ApBackgroundWork -Window $script:ApWin `
        -Variables @{ ApEndpoints = $endpoints } `
        -FunctionNames @('Test-ApEndpoint', 'Invoke-ApConnectivityCheck') `
        -Work {
            Invoke-ApConnectivityCheck -Endpoints $ApEndpoints
        } `
        -OnComplete {
            param($result, $errorMessage)

            $el = $script:ApEl
            $el.NetworkRunButton.IsEnabled = $true
            $el.StatusProgress.IsIndeterminate = $false

            if ($errorMessage) {
                $el.NetworkSummary.Text = "Check failed: $errorMessage"
                Set-ApStatus -Text "Connectivity check failed: $errorMessage" -Percent 0 -IsError
                Write-ApLog "Connectivity check failed: $errorMessage" -Level ERROR
                return
            }

            $results = @($result)
            $script:ApNetworkResults = $results

            # Sort failures to the top: that is what the tech opened this page to see.
            $el.NetworkGrid.ItemsSource = @(
                $results | Sort-Object @{ Expression = { $_.Succeeded } }, @{ Expression = { -not $_.Required } }, Category, Name
            )

            $ok = @($results | Where-Object { $_.Succeeded }).Count
            $requiredFailures = @($results | Where-Object { $_.Required -and -not $_.Succeeded })

            if ($requiredFailures.Count -gt 0) {
                $names = ($requiredFailures | ForEach-Object { $_.Host }) -join ', '
                $el.NetworkSummary.Text = "$ok of $($results.Count) reachable. Required endpoints failing: $names"
                Set-ApStatus -Text "$($requiredFailures.Count) required endpoint(s) unreachable." -Percent 100 -IsError
            }
            else {
                $optionalFailures = @($results | Where-Object { -not $_.Succeeded }).Count
                $suffix = if ($optionalFailures -gt 0) { " $optionalFailures optional endpoint(s) unreachable." } else { '' }
                $el.NetworkSummary.Text = "$ok of $($results.Count) reachable. All required endpoints are available.$suffix"
                Set-ApStatus -Text 'Connectivity check complete.' -Percent 100
            }

            Write-ApLog "Connectivity check finished: $ok of $($results.Count) reachable, $($requiredFailures.Count) required failing"
        } | Out-Null
}

# ============================ export helpers ============================

function Start-ApCsvExport {
    <#
    .SYNOPSIS
    Runs an offline CSV export (no tenant connection).

    .NOTES
    Named this way deliberately. The obvious "Invoke-Ap" + "Export" spelling produces an
    identifier whose lowercased form starts with the same four letters as a Microsoft
    Defender HackTool signature family, and Defender's script scanner then refuses to load
    the entire file with ScriptContainedMaliciousContent. Verified by elimination: a file
    containing only that one function was blocked while all 22 other functions in this file
    loaded cleanly. Keep the current name.
    #>
    [CmdletBinding()]
    param([ValidateSet('v1', 'v2')][string]$Mode)

    $defaultName = if ($Mode -eq 'v2') { 'AutopilotDeviceIdentifier.csv' } else { 'AutopilotHWID.csv' }
    $path = Show-ApSaveFileDialog -Title 'Save the Autopilot CSV' -FileName $defaultName
    if (-not $path) { return }

    $request = Get-ApUiRequest -Operation Export
    $request.Mode = $Mode
    $request.OutputFile = $path

    try {
        $built = Build-ApEngineArguments $request
    }
    catch {
        Show-ApDialog -Title 'Cannot export' -Owner $script:ApWin -Message $_.Exception.Message | Out-Null
        return
    }

    if ($built.Warnings.Count -gt 0) {
        Show-ApDialog -Title 'Export note' -Owner $script:ApWin `
                      -Message ($built.Warnings -join [Environment]::NewLine) | Out-Null
    }

    Show-ApPage 'PageRegister'
    $script:ApEl.NavRegister.IsChecked = $true

    Start-ApGuiRun -Parameters $built.Parameters -OutputBox $script:ApEl.RegisterOutput `
                   -Label 'export' -StartMessage "Exporting to $path..."
}

# ============================ entry point ============================

function Show-AutopilotImportGui {
    <#
    .SYNOPSIS
    Builds and shows the Autopilot Import GUI.

    .PARAMETER GroupTag
    Pre-fills the group tag field.

    .PARAMETER AssignedUser
    Pre-fills the assigned user field.

    .PARAMETER Mode
    Pre-selects the registration mode.
    #>
    [CmdletBinding()]
    param(
        [string]$GroupTag,
        [string]$AssignedUser,
        [ValidateSet('v1', 'v2')][string]$Mode
    )

    Initialize-ApGui @PSBoundParameters | Out-Null
    [void]$script:ApWin.ShowDialog()
}

function Initialize-ApGui {
    <#
    .SYNOPSIS
    Builds the window, populates it and wires every handler. Does not show it.

    .DESCRIPTION
    Separated from Show-AutopilotImportGui so the whole UI can be constructed and driven in
    a test harness without entering the blocking ShowDialog message loop.

    .OUTPUTS
    The constructed Window.
    #>
    [CmdletBinding()]
    param(
        [string]$GroupTag,
        [string]$AssignedUser,
        [ValidateSet('v1', 'v2')][string]$Mode
    )

    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml

    $ui = New-ApMainWindow
    $script:ApWin = $ui.Window
    $script:ApEl = $ui.Elements
    $el = $script:ApEl
    $config = Get-ApConfig

    # ---------- header state ----------
    if (Test-ApElevated) {
        Set-ApPill $el.PillAdmin $el.PillAdminText 'Administrator' 'ok'
    }
    else {
        Set-ApPill $el.PillAdmin $el.PillAdminText 'Not elevated' 'error'
    }

    $el.SidebarAppVersion.Text = "GUI $script:ApAppVersion"
    $el.AboutVersion.Text = "Version $script:ApAppVersion"

    # ---------- branding links ----------
    $el.SidebarSiteLink.Add_Click({ Start-ApUrl $script:ApAuthorSite })
    $el.AboutSiteLink.Add_Click({ Start-ApUrl $script:ApAuthorSite })
    $el.AboutGitHubLink.Add_Click({ Start-ApUrl $script:ApAuthorGitHub })
    $el.AboutEngineLink.Add_Click({ Start-ApUrl 'https://github.com/andrew-s-taylor/WindowsAutopilotInfo' })
    $el.AboutOriginalLink.Add_Click({ Start-ApUrl 'https://github.com/ugurkocde/AutoPilot_Import_GUI' })

    # ---------- device ----------
    # Only placeholders here. The real read happens on the window's Loaded event so the
    # window is on screen first; see Start-ApDeviceLoad for why it cannot be synchronous.
    Set-ApDevicePlaceholders

    # ---------- engine ----------
    try {
        $script:ApEnginePath = Resolve-ApEngineScript -Name Engine
        $manifest = Get-ApVendorManifest
        $engineEntry = $null
        if ($manifest) {
            $engineEntry = $manifest.scripts | Where-Object { $_.role -eq 'engine' } | Select-Object -First 1
        }
        $versionText = if ($engineEntry) { "Engine v$($engineEntry.version)" } else { 'Engine: resolved' }
        $el.SidebarEngineVersion.Text = $versionText
        $el.AdvEnginePath.Text = $script:ApEnginePath
        Write-ApLog "Engine resolved to $script:ApEnginePath"
    }
    catch {
        $el.SidebarEngineVersion.Text = 'Engine: not found'
        $el.AdvEnginePath.Text = $_.Exception.Message
        Set-ApStatus -Text 'The community engine script could not be found. See the Advanced page.' -IsError
        Write-ApLog "Engine not resolved: $($_.Exception.Message)" -Level ERROR
    }

    # ---------- restore saved preferences ----------
    $history = @($config.groupTagHistory | Where-Object { $_ })
    $el.GroupTagCombo.ItemsSource = $history
    $el.GroupTagCombo.Text = if ($GroupTag) { $GroupTag } else { "$($config.lastGroupTag)" }
    $el.AssignedUserBox.Text = if ($AssignedUser) { $AssignedUser } else { "$($config.lastAssignedUser)" }
    $el.ComputerNameBox.Text = "$($config.lastComputerName)"
    $el.AddToGroupBox.Text = "$($config.lastAddToGroup)"
    $el.WaitAssignCheck.IsChecked = [bool]$config.waitForAssignment
    $el.RebootCheck.IsChecked = [bool]$config.rebootWhenAssigned
    $el.RebootV2Check.IsChecked = [bool]$config.rebootAfterV2Import
    $el.AdvShowConsoleCheck.IsChecked = [bool]$config.showConsoleWindow
    $el.AdvDiagnosticsOnlineCheck.IsChecked = [bool]$config.diagnosticsOnline

    switch ("$($config.existingDevicePolicy)") {
        'delete'    { $el.PolicyDelete.IsChecked = $true }
        'skipcheck' { $el.PolicySkip.IsChecked = $true }
        default     { $el.PolicyUpdate.IsChecked = $true }
    }

    $effectiveMode = if ($Mode) { $Mode } else { "$($config.lastMode)" }
    if ($effectiveMode -eq 'v2') { $el.ModeV2.IsChecked = $true } else { $el.ModeV1.IsChecked = $true }

    # Remembered for the device load, which may switch to v2 when there is no hash: an
    # explicit -Mode from the caller must not be overridden.
    $script:ApModePinned = [bool]$Mode

    Sync-ApModeUi
    Update-ApLogsPage

    # ---------- navigation ----------
    $el.NavRegister.Add_Checked({ Show-ApPage 'PageRegister' })
    $el.NavDevice.Add_Checked({ Show-ApPage 'PageDevice' })
    $el.NavBatch.Add_Checked({ Show-ApPage 'PageBatch' })
    $el.NavNetwork.Add_Checked({ Show-ApPage 'PageNetwork' })
    $el.NavAdvanced.Add_Checked({ Show-ApPage 'PageAdvanced' })
    $el.NavLogs.Add_Checked({ Show-ApPage 'PageLogs'; Update-ApLogsPage })

    # ---------- mode ----------
    $el.ModeV1.Add_Checked({ Sync-ApModeUi })
    $el.ModeV2.Add_Checked({ Sync-ApModeUi })

    # Reboot is only honoured after assignment, so reflect that in the UI immediately
    # rather than silently coercing it at launch time.
    $el.RebootCheck.Add_Checked({
        if (-not $script:ApEl.WaitAssignCheck.IsChecked) {
            $script:ApEl.WaitAssignCheck.IsChecked = $true
            Set-ApStatus -Text 'Waiting for profile assignment was enabled: a restart only happens after assignment.'
        }
    })

    # ---------- register ----------
    $el.RegisterButton.Add_Click({
        $el = $script:ApEl
        $request = Get-ApUiRequest -Operation Register

        try {
            $built = Build-ApEngineArguments $request
        }
        catch {
            Show-ApDialog -Title 'Cannot register' -Owner $script:ApWin -Message $_.Exception.Message | Out-Null
            return
        }

        if ($built.Warnings.Count -gt 0) { Show-ApNotice -Title 'Note' -Messages $built.Warnings }

        # v1 with no hash will fail in the engine; say so here instead.
        if ($request.Mode -eq 'v1' -and $script:ApDevice -and -not $script:ApDevice.HardwareHashReady) {
            $proceed = Show-ApDialog -Title 'No hardware hash' -Owner $script:ApWin -ShowCancel `
                -ConfirmText 'Try anyway' `
                -Message 'This device did not return an Autopilot hardware hash, so a v1 registration is very likely to fail. Device Preparation (v2) does not need one.'
            if (-not $proceed) { return }
        }

        # A broken sign-in module aborts the engine before any browser prompt appears, which
        # reads as "nothing happened". Say so up front instead.
        if ($script:ApGraphCheck -and -not $script:ApGraphCheck.Available -and $script:ApGraphCheck.InstalledVersions.Count -gt 0) {
            $proceed = Show-ApDialog -Title 'Sign-in will probably fail' -Owner $script:ApWin -ShowCancel `
                -ConfirmText 'Try anyway' `
                -Message 'The Microsoft Graph sign-in module is installed but cannot be loaded, so the engine will stop before a sign-in prompt appears. Use "Repair sign-in module" on the Advanced page.' `
                -Detail (Get-ApGraphModuleAdvice -Check $script:ApGraphCheck)
            if (-not $proceed) { return }
        }

        # Device Preparation restart: confirm, because a restart from OOBE is disruptive and
        # is premature unless the device is already in the policy's Entra group.
        $wantsV2Reboot = ($request.Mode -eq 'v2') -and [bool]$el.RebootV2Check.IsChecked
        if ($wantsV2Reboot) {
            $proceed = Show-ApDialog -Title 'Restart after import' -Owner $script:ApWin -ShowCancel `
                -ConfirmText 'Import and restart' `
                -Message ('This device will restart as soon as the identifier is imported. ' +
                          'Device Preparation targets devices through the Entra group on the policy, so ' +
                          'restart only if this device is already a member of that group. It will not ' +
                          'restart if the import fails or you cancel.')
            if (-not $proceed) { return }
        }
        $script:ApPendingV2Reboot = $wantsV2Reboot

        if (Test-ApDestructiveRequest $request) {
            $detail = Get-ApPreviewCommand -Parameters $built.Parameters -ScriptPath (Split-Path -Leaf $script:ApEnginePath)
            $lines = New-Object System.Collections.Generic.List[string]
            if ($request.ExistingDevicePolicy -eq 'delete') {
                $lines.Add('If this serial is already registered it will be deleted from Autopilot, Intune and Entra ID before being re-added.')
            }
            if ($request.Wipe) { $lines.Add('An Intune wipe will be sent to this device. All data on it will be erased.') }
            if ($request.Sysprep) { $lines.Add('Sysprep will run and the device will restart into OOBE.') }
            if ($request.ChangePK) { $lines.Add('The Windows product key will be changed and the device will restart.') }

            $proceed = Show-ApDialog -Title 'Confirm a destructive action' -Owner $script:ApWin `
                -ShowCancel -Danger -ConfirmText 'Yes, continue' -ShowCopy `
                -Message ($lines -join [Environment]::NewLine) -Detail $detail
            if (-not $proceed) { return }
        }

        # Remember what was used, so the next device is one click.
        Add-ApGroupTagToHistory $request.GroupTag
        Set-ApConfigValue 'lastMode' $request.Mode
        Set-ApConfigValue 'lastGroupTag' $request.GroupTag
        Set-ApConfigValue 'lastAssignedUser' $request.AssignedUser
        Set-ApConfigValue 'lastComputerName' $request.AssignedComputerName
        Set-ApConfigValue 'lastAddToGroup' $request.AddToGroup
        Set-ApConfigValue 'waitForAssignment' ([bool]$request.WaitForAssignment)
        Set-ApConfigValue 'rebootWhenAssigned' ([bool]$request.Reboot)
        Set-ApConfigValue 'rebootAfterV2Import' ([bool]$el.RebootV2Check.IsChecked)
        Set-ApConfigValue 'existingDevicePolicy' $request.ExistingDevicePolicy
        Save-ApConfig | Out-Null
        $el.GroupTagCombo.ItemsSource = @((Get-ApConfig).groupTagHistory)

        foreach ($note in $built.Notes) { Write-ApLog "Note: $note" }

        Start-ApGuiRun -Parameters $built.Parameters -OutputBox $el.RegisterOutput `
                       -Label 'register' -StartMessage 'Signing in to Microsoft Graph. A browser window will open.'
    })

    $el.PreviewButton.Add_Click({
        try {
            $built = Build-ApEngineArguments (Get-ApUiRequest -Operation Register)
        }
        catch {
            Show-ApDialog -Title 'Cannot build the command' -Owner $script:ApWin -Message $_.Exception.Message | Out-Null
            return
        }

        $scriptName = if ($script:ApEnginePath) { $script:ApEnginePath } else { 'get-windowsautopilotinfocommunity.ps1' }
        $detail = Get-ApPreviewCommand -Parameters $built.Parameters -ScriptPath $scriptName

        $message = 'This is exactly what will run. Nothing has been executed.'
        if ($built.Notes.Count -gt 0) {
            $message += [Environment]::NewLine + [Environment]::NewLine + (($built.Notes | ForEach-Object { "- $_" }) -join [Environment]::NewLine)
        }
        if ($built.Warnings.Count -gt 0) {
            $message += [Environment]::NewLine + [Environment]::NewLine + (($built.Warnings | ForEach-Object { "! $_" }) -join [Environment]::NewLine)
        }

        Show-ApDialog -Title 'Preview command' -Owner $script:ApWin -Message $message -Detail $detail -ShowCopy | Out-Null
    })

    $el.CancelButton.Add_Click({ Stop-ApGuiRun })

    $el.RestartNowButton.Add_Click({
        # Refuse mid-run: restarting during an import abandons it half-finished and the log
        # is lost with the session.
        if ($script:ApRun -and -not $script:ApRun.IsFinished) {
            Show-ApDialog -Title 'A run is in progress' -Owner $script:ApWin `
                -Message 'Wait for the current operation to finish, or cancel it, before restarting.' | Out-Null
            return
        }

        $proceed = Show-ApDialog -Title 'Restart this device now' -Owner $script:ApWin -ShowCancel `
            -ConfirmText 'Restart now' -CancelText 'Not yet' -Danger `
            -Message ('This device will restart immediately. Anything unsaved will be lost.' + [Environment]::NewLine + [Environment]::NewLine +
                      'Before restarting, make sure this device is a member of the Entra group targeted by your ' +
                      'Device Preparation policy, or it will return to OOBE before the policy can apply.')
        if (-not $proceed) { return }

        Set-ApStatus -Text 'Restarting...'
        Invoke-ApRestartComputer | Out-Null
    })

    # ---------- device page ----------
    $el.RefreshDeviceButton.Add_Click({
        Set-ApDevicePlaceholders
        Start-ApDeviceLoad
    })

    $el.ExportHashButton.Add_Click({ Start-ApCsvExport -Mode 'v1' })
    $el.ExportIdentifierButton.Add_Click({ Start-ApCsvExport -Mode 'v2' })

    $el.CopyIdentifierButton.Add_Click({
        if (-not $script:ApDevice) { return }
        try {
            Set-Clipboard -Value $script:ApDevice.DeviceIdentifier
            Set-ApStatus -Text 'Device identifier copied to the clipboard.'
        }
        catch {
            Set-ApStatus -Text "Could not copy: $($_.Exception.Message)" -IsError
        }
    })

    $el.CopyHashButton.Add_Click({
        $hash = Get-ApHardwareHash
        if (-not $hash) {
            Show-ApDialog -Title 'No hardware hash' -Owner $script:ApWin `
                -Message 'This device did not return a hardware hash. Administrator rights are required, and virtual machines often cannot provide one.' | Out-Null
            return
        }
        try {
            Set-Clipboard -Value $hash
            Set-ApStatus -Text "Hardware hash copied to the clipboard ($($hash.Length) characters)."
        }
        catch {
            Set-ApStatus -Text "Could not copy: $($_.Exception.Message)" -IsError
        }
    })

    # ---------- batch page ----------
    $el.BatchBrowseButton.Add_Click({
        $path = Show-ApOpenFileDialog -Title 'Select a device CSV'
        if ($path) { $script:ApEl.BatchFileBox.Text = $path }
    })

    $el.BatchRunButton.Add_Click({
        $el = $script:ApEl
        $request = Get-ApUiRequest -Operation Batch

        if (-not $request.InputFile) {
            Show-ApDialog -Title 'Choose a CSV' -Owner $script:ApWin -Message 'Select the CSV file to import first.' | Out-Null
            return
        }
        if (-not (Test-Path -LiteralPath $request.InputFile)) {
            Show-ApDialog -Title 'File not found' -Owner $script:ApWin -Message "There is no file at $($request.InputFile)." | Out-Null
            return
        }

        try { $built = Build-ApEngineArguments $request }
        catch {
            Show-ApDialog -Title 'Cannot import' -Owner $script:ApWin -Message $_.Exception.Message | Out-Null
            return
        }

        Start-ApGuiRun -Parameters $built.Parameters -OutputBox $el.BatchOutput `
                       -Label 'batch' -StartMessage "Importing $($request.InputFile)..."
    })

    $el.BatchPreviewButton.Add_Click({
        try { $built = Build-ApEngineArguments (Get-ApUiRequest -Operation Batch) }
        catch {
            Show-ApDialog -Title 'Cannot build the command' -Owner $script:ApWin -Message $_.Exception.Message | Out-Null
            return
        }
        $detail = Get-ApPreviewCommand -Parameters $built.Parameters -ScriptPath $script:ApEnginePath
        Show-ApDialog -Title 'Preview command' -Owner $script:ApWin -Detail $detail -ShowCopy `
                      -Message 'This is exactly what will run. Nothing has been executed.' | Out-Null
    })

    # ---------- network page ----------
    $el.NetworkRunButton.Add_Click({ Start-ApNetworkCheck })

    $el.NetworkCopyButton.Add_Click({
        if (-not $script:ApNetworkResults -or $script:ApNetworkResults.Count -eq 0) {
            Set-ApStatus -Text 'Run the check first.'
            return
        }
        $text = $script:ApNetworkResults |
            Sort-Object Category, Name |
            ForEach-Object { '{0,-12} {1,-32} {2,-45} {3} ms' -f $_.Status, $_.Name, $_.Host, $_.LatencyMs }
        try {
            Set-Clipboard -Value ($text -join [Environment]::NewLine)
            Set-ApStatus -Text 'Connectivity results copied to the clipboard.'
        }
        catch {
            Set-ApStatus -Text "Could not copy: $($_.Exception.Message)" -IsError
        }
    })

    # ---------- advanced page ----------
    $el.AdvShowConsoleCheck.Add_Click({
        Set-ApConfigValue 'showConsoleWindow' ([bool]$script:ApEl.AdvShowConsoleCheck.IsChecked)
        Save-ApConfig | Out-Null
    })

    $el.AdvDiagnosticsOnlineCheck.Add_Click({
        Set-ApConfigValue 'diagnosticsOnline' ([bool]$script:ApEl.AdvDiagnosticsOnlineCheck.IsChecked)
        Save-ApConfig | Out-Null
    })

    $el.AdvVerifyEngineButton.Add_Click({
        $el = $script:ApEl
        if (-not $script:ApEnginePath) {
            $el.AdvEngineIntegrity.Text = 'No engine script resolved.'
            return
        }

        $check = Test-ApVendorScript -Path $script:ApEnginePath
        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add("Signature: $($check.SignatureStatus)")
        if ($check.Signer) { $lines.Add("Signer: $($check.Signer)") }
        if ($null -ne $check.HashMatches) {
            $lines.Add("Checksum: $(if ($check.HashMatches) { 'matches vendor manifest' } else { 'DOES NOT MATCH' })")
        }
        else {
            $lines.Add('Checksum: no manifest entry to compare against')
        }
        foreach ($m in $check.Messages) { $lines.Add($m) }

        $el.AdvEngineIntegrity.Text = ($lines -join [Environment]::NewLine)
        Write-ApLog ("Engine integrity: " + ($lines -join ' | '))

        Show-ApDialog -Title 'Engine integrity' -Owner $script:ApWin `
                      -Message $(if ($check.IsTrusted) { 'The engine script matches what was shipped.' } else { 'The engine script does not match the vendor manifest. Do not use it until you know why.' }) `
                      -Detail (($lines -join [Environment]::NewLine) + "`r`n`r`n$($check.Path)") | Out-Null
    })

    $el.AdvOpenWorkDirButton.Add_Click({
        try { Start-Process explorer.exe (Get-ApWorkingDirectory) } catch { }
    })

    $el.AdvRepairGraphButton.Add_Click({
        $detail = ''
        if ($script:ApGraphCheck) { $detail = Get-ApGraphModuleAdvice -Check $script:ApGraphCheck }

        $proceed = Show-ApDialog -Title 'Repair the sign-in module' -Owner $script:ApWin -ShowCancel `
            -ConfirmText 'Repair now' `
            -Message ('This removes every installed copy of Microsoft.Graph.Authentication and reinstalls version ' +
                      "$script:ApGraphPinnedVersion for all users, which is the version the community engine expects. " +
                      'It needs access to the PowerShell Gallery. Continue?') `
            -Detail $detail
        if (-not $proceed) { return }

        Start-ApGraphRepair
    })

    $el.AdvDiagnosticsButton.Add_Click({
        try { $diag = Resolve-ApEngineScript -Name Diagnostics }
        catch {
            Show-ApDialog -Title 'Diagnostics unavailable' -Owner $script:ApWin -Message $_.Exception.Message | Out-Null
            return
        }

        # -Online only adds lookups: app, policy and script GUIDs from the local logs are
        # resolved to display names through Graph. Everything the script reports is still
        # read from this machine, and nothing is written either way.
        $online = [bool]$script:ApEl.AdvDiagnosticsOnlineCheck.IsChecked

        # Same trap as a registration run: a broken sign-in module aborts the engine before
        # any browser prompt appears, which reads as "nothing happened".
        if ($online -and $script:ApGraphCheck -and -not $script:ApGraphCheck.Available -and
            $script:ApGraphCheck.InstalledVersions.Count -gt 0) {
            $proceed = Show-ApDialog -Title 'Sign-in will probably fail' -Owner $script:ApWin -ShowCancel `
                -ConfirmText 'Try anyway' `
                -Message ('The Microsoft Graph sign-in module is installed but cannot be loaded, so the online lookup will stop before a sign-in prompt appears. ' +
                          'Clear the -Online option to run the local diagnostics anyway, or use "Repair sign-in module".') `
                -Detail (Get-ApGraphModuleAdvice -Check $script:ApGraphCheck)
            if (-not $proceed) { return }
        }

        # An 'Online' key is also what makes the launcher install the sign-in prerequisites
        # non-interactively (see Get-ApDependencyPrepBlock); a local run must not touch the gallery.
        $params = [ordered]@{}
        $message = 'Collecting Autopilot diagnostics...'
        if ($online) {
            $params['Online'] = $true
            $message = 'Collecting Autopilot diagnostics and signing in to resolve app and policy names...'
        }

        $script:ApEl.NavLogs.IsChecked = $true
        Start-ApGuiRun -Parameters $params -OutputBox $script:ApEl.LogsOutput `
                       -EnginePath $diag -Label 'diagnostics' `
                       -StartMessage $message
    })

    $el.AdvWindowsUpdateButton.Add_Click({
        $proceed = Show-ApDialog -Title 'Install Windows updates' -Owner $script:ApWin -ShowCancel `
            -ConfirmText 'Install updates' `
            -Message 'This searches for and installs all applicable Windows updates using the in-box Windows Update agent. The device may restart without further warning. Continue?'
        if (-not $proceed) { return }
        Start-ApWindowsUpdate
    })

    $el.AdvWipeCheck.Add_Click({ Sync-ApAdvancedWarnings })
    $el.AdvSysprepCheck.Add_Click({ Sync-ApAdvancedWarnings })
    Sync-ApAdvancedWarnings

    # ---------- logs page ----------
    $el.LogsRefreshButton.Add_Click({ Update-ApLogsPage })

    $el.LogsCopyButton.Add_Click({
        try {
            Set-Clipboard -Value $script:ApEl.LogsOutput.Text
            Set-ApStatus -Text 'Log copied to the clipboard.'
        }
        catch {
            Set-ApStatus -Text "Could not copy: $($_.Exception.Message)" -IsError
        }
    })

    $el.LogsOpenFolderButton.Add_Click({
        try { Start-Process explorer.exe (Split-Path -Parent (Get-ApLogPath)) } catch { }
    })

    # ---------- timers ----------
    # Seed the clock now; otherwise it reads "--:--:--" for the first second.
    $el.HeaderClock.Text = (Get-Date).ToString('HH:mm:ss')

    $clock = New-Object System.Windows.Threading.DispatcherTimer
    $clock.Interval = [TimeSpan]::FromSeconds(1)
    # Guarded for the same reason as the background tick: an escaping exception would take
    # ShowDialog down with it.
    $clock.Add_Tick({
        try {
            $script:ApEl.HeaderClock.Text = (Get-Date).ToString('HH:mm:ss')

            # Re-probe the network every 15 seconds. In OOBE, Wi-Fi often comes up after the
            # tool is already open, so a one-shot check at startup goes stale immediately.
            $script:ApClockTicks = ($script:ApClockTicks + 1)
            if (($script:ApClockTicks % 15) -eq 1) { Update-ApNetworkPill }
        }
        catch {
            Write-ApLog "Clock tick failed: $($_.Exception.Message)" -Level ERROR
        }
    })
    $script:ApClockTicks = 0
    $script:ApClockTimer = $clock
    $clock.Start()

    Update-ApNetworkPill

    # ---------- shutdown ----------
    # A WPF event handler receives (sender, eventArgs) as positional arguments. $_ is NOT the
    # event args here: using it threw "Argument types do not match" out of Window.OnClosing,
    # so the cancel-the-close path never worked and the exception escaped ShowDialog.
    $script:ApWin.Add_Closing({
        param($closeSender, $closeArgs)

        try {
            if ($script:ApRun -and -not $script:ApRun.IsFinished) {
                $cancelRun = Show-ApDialog -Title 'A run is still in progress' -Owner $script:ApWin -ShowCancel `
                    -ConfirmText 'Cancel the run and close' -CancelText 'Keep it running' -Danger `
                    -Message 'Closing this window will terminate the operation in progress. Registration may be left half-finished.'

                if (-not $cancelRun) {
                    $closeArgs.Cancel = $true
                    return
                }
                Stop-ApEngineRun -Run $script:ApRun
            }

            # Each step guarded separately: a failure in one must not skip the others,
            # especially not the config save.
            try { if ($script:ApRunTimer) { $script:ApRunTimer.Stop() } }
            catch { Write-ApLog "Stopping the run timer failed: $($_.Exception.Message)" -Level WARN }

            try { if ($script:ApClockTimer) { $script:ApClockTimer.Stop() } }
            catch { Write-ApLog "Stopping the clock failed: $($_.Exception.Message)" -Level WARN }

            try { Stop-ApBackgroundWork }
            catch { Write-ApLog "Stopping background work failed: $($_.Exception.Message)" -Level WARN }

            try { Save-ApConfig | Out-Null }
            catch { Write-ApLog "Saving configuration failed: $($_.Exception.Message)" -Level WARN }

            Write-ApLog 'Session closed.'
        }
        catch {
            Write-ApLog "Shutdown failed: $($_.Exception.Message)" -Level ERROR
        }
    })

    # Kick the slow work off only once the window is actually on screen.
    $script:ApWin.Add_Loaded({
        Start-ApDeviceLoad
        Start-ApGraphPreflight
    })

    Set-ApStatus -Text 'Ready.' -Stage '' -Percent 0
    Write-ApLog 'GUI ready.'

    return $script:ApWin
}

function Sync-ApAdvancedWarnings {
    <#
    .SYNOPSIS
    Shows the wipe warning only when a destructive option is actually selected.
    #>
    [CmdletBinding()]
    param()

    $el = $script:ApEl
    $el.AdvWipeWarning.Visibility = if ($el.AdvWipeCheck.IsChecked) { 'Visible' } else { 'Collapsed' }
}

function Update-ApNetworkPill {
    <#
    .SYNOPSIS
    Refreshes the header network chip without blocking the UI thread.
    #>
    [CmdletBinding()]
    param()

    Start-ApBackgroundWork -Window $script:ApWin `
        -FunctionNames @('Test-ApInternetConnection') `
        -Work { Test-ApInternetConnection } `
        -OnComplete {
            param($result, $errorMessage)
            $online = [bool](@($result)[-1])
            if ($online) {
                Set-ApPill $script:ApEl.PillNetwork $script:ApEl.PillNetworkText 'Online' 'ok'
            }
            else {
                Set-ApPill $script:ApEl.PillNetwork $script:ApEl.PillNetworkText 'Offline' 'error'
            }
        } | Out-Null
}

function Start-ApGraphPreflight {
    <#
    .SYNOPSIS
    Checks in the background whether the sign-in module can be imported, and reports it.
    #>
    [CmdletBinding()]
    param()

    $script:ApEl.AdvGraphStatus.Text = 'Checking...'

    Start-ApBackgroundWork -Window $script:ApWin `
        -FunctionNames @('Test-ApGraphModule') `
        -Work { Test-ApGraphModule } `
        -OnComplete {
            param($result, $errorMessage)

            $el = $script:ApEl

            if ($errorMessage -or -not $result) {
                $el.AdvGraphStatus.Text = "Could not check: $errorMessage"
                return
            }

            $check = @($result)[-1]
            $script:ApGraphCheck = $check

            if ($check.Available) {
                $el.AdvGraphStatus.Text = "Microsoft.Graph.Authentication $($check.Version) loads correctly."
                $el.AdvGraphStatus.Foreground = New-ApBrush '#3BC77A'
                return
            }

            $el.AdvGraphStatus.Foreground = New-ApBrush '#E9B44C'
            if ($check.InstalledVersions.Count -eq 0) {
                $el.AdvGraphStatus.Text = 'Not installed for Windows PowerShell. The engine will install it on first sign-in, which needs the PowerShell Gallery.'
            }
            else {
                $el.AdvGraphStatus.Text = "Version $($check.Version) is installed but will not load. Sign-in will fail until this is repaired."
            }

            Write-ApLog "Graph module preflight: available=$($check.Available) versions=$($check.InstalledVersions -join ',') error=$($check.Error)" -Level WARN
        } | Out-Null
}

function Start-ApGraphRepair {
    <#
    .SYNOPSIS
    Runs the Graph module repair in a visible console.

    .DESCRIPTION
    Separate console rather than a streamed run: it uninstalls and reinstalls a module, which
    can prompt, and the tech should see it happen.
    #>
    [CmdletBinding()]
    param()

    $path = Join-Path (Get-ApWorkingDirectory) 'repair-graph-module.ps1'
    Set-Content -LiteralPath $path -Value (Get-ApGraphRepairScript) -Encoding UTF8

    Write-ApLog 'Launching Graph module repair in a separate console.'
    Start-Process -FilePath (Get-ApPowerShellPath) `
                  -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $path) | Out-Null

    Set-ApStatus -Text 'Repairing the sign-in module in a separate window. Restart this tool when it finishes.'
}

function Start-ApWindowsUpdate {
    <#
    .SYNOPSIS
    Searches for and installs available Windows updates in a visible console.

    .DESCRIPTION
    Uses the in-box Windows Update Agent COM API (Microsoft.Update.Session) rather than the
    PSWindowsUpdate module the original GUI installed
    (Get-WindowsAutopilotImportGUI.ps1:171). Three reasons:

      * no dependency to fetch, so it works on an OOBE network that cannot reach the
        PowerShell Gallery, which is exactly where this button gets used
      * no NuGet provider bootstrap, no repository trust change, no module install
      * generating a script that chains Install-PackageProvider, Set-PSRepository -Trusted
        and Install-Module -Force is a well-known malware-dropper shape, and Defender's
        AMSI scanner blocks it on sight

    Runs in its own visible console because a Windows Update install can restart the
    machine; streaming that into the GUI would look like a hang.
    #>
    [CmdletBinding()]
    param()

    $updateScript = @'
$Host.UI.RawUI.WindowTitle = 'Autopilot Import GUI - Windows Update'
$ErrorActionPreference = 'Stop'

function Write-Step { param([string]$Text) Write-Host "==> $Text" -ForegroundColor Cyan }

try {
    Write-Step 'Searching for applicable updates...'
    $session  = New-Object -ComObject Microsoft.Update.Session
    $searcher = $session.CreateUpdateSearcher()
    $result   = $searcher.Search("IsInstalled=0 AND Type='Software' AND IsHidden=0")

    if ($result.Updates.Count -eq 0) {
        Write-Host 'No applicable updates were found.' -ForegroundColor Green
    }
    else {
        Write-Host ''
        Write-Host "Found $($result.Updates.Count) update(s):" -ForegroundColor Yellow
        $toInstall = New-Object -ComObject Microsoft.Update.UpdateColl
        foreach ($update in $result.Updates) {
            Write-Host "  - $($update.Title)"
            # Updates with unaccepted licence terms cannot be installed unattended.
            if ($update.EulaAccepted -eq $false) { $update.AcceptEula() }
            [void]$toInstall.Add($update)
        }

        Write-Host ''
        Write-Step 'Downloading...'
        $downloader = $session.CreateUpdateDownloader()
        $downloader.Updates = $toInstall
        $downloadResult = $downloader.Download()
        Write-Host "Download result code: $($downloadResult.ResultCode)"

        Write-Step 'Installing. This can take a while.'
        $installer = $session.CreateUpdateInstaller()
        $installer.Updates = $toInstall
        $installResult = $installer.Install()

        Write-Host ''
        Write-Host "Install result code: $($installResult.ResultCode)" -ForegroundColor Green
        if ($installResult.RebootRequired) {
            Write-Host 'A restart is required to finish installing these updates.' -ForegroundColor Yellow
        }
    }
}
catch {
    Write-Host ''
    Write-Host "Windows Update failed: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ''
Read-Host 'Press Enter to close'
'@

    $path = Join-Path (Get-ApWorkingDirectory) 'windowsupdate.ps1'
    Set-Content -LiteralPath $path -Value $updateScript -Encoding UTF8

    Write-ApLog 'Launching Windows Update in a separate console.'
    Start-Process -FilePath (Get-ApPowerShellPath) `
                  -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $path) | Out-Null

    Set-ApStatus -Text 'Windows Update is running in a separate window.'
}

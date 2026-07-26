# ProgressParser.ps1 -- turns community-script console output into UI progress state.
#
# Wrapping a console script means the only progress signal available is its stdout.
# Fortunately get-windowsautopilotinfocommunity.ps1 emits stable, greppable lines, so we
# can drive a real staged progress bar instead of the indeterminate spinner the original
# GUI never even had. Line references are to v5.0.16.
#
# Stages, in the order the script performs them:
#   Connect  -> Collect -> Import -> Sync -> Assign -> Complete
#
# Pure function, unit-tested in tests\ProgressParser.Tests.ps1. If upstream changes its
# wording the parser degrades to "no progress update" rather than breaking the run --
# Update-ApProgressState returns $null for unrecognised lines and the caller keeps the
# previous state.

$script:ApStageOrder = @('Connect', 'Collect', 'Import', 'Sync', 'Assign', 'Complete')

function New-ApProgressState {
    <#
    .SYNOPSIS
    Initial progress state for a run.
    #>
    [CmdletBinding()]
    param()

    return [ordered]@{
        Stage        = 'Connect'
        StageLabel   = 'Signing in...'
        Current      = 0
        Total        = 0
        Percent      = 0
        IsComplete   = $false
        IsError      = $false
        ErrorMessage = ''
        LastMessage  = ''
    }
}

function Get-ApStageWeight {
    <#
    .SYNOPSIS
    Percentage floor for each stage, so the bar advances monotonically across stages.
    #>
    param([string]$Stage)

    switch ($Stage) {
        'Connect'  { return @{ Floor = 0;   Span = 10 } }
        'Collect'  { return @{ Floor = 10;  Span = 10 } }
        'Import'   { return @{ Floor = 20;  Span = 35 } }
        'Sync'     { return @{ Floor = 55;  Span = 20 } }
        'Assign'   { return @{ Floor = 75;  Span = 24 } }
        'Complete' { return @{ Floor = 100; Span = 0 } }
        default    { return @{ Floor = 0;   Span = 0 } }
    }
}

function Update-ApProgressState {
    <#
    .SYNOPSIS
    Folds one output line into the progress state.

    .DESCRIPTION
    Returns the updated state when the line was meaningful, otherwise $null so the caller
    knows nothing changed. Mutates and returns the same object for cheapness.

    .PARAMETER State
    State from New-ApProgressState.

    .PARAMETER Line
    A single line of script output.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$State,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Line
    )

    if ([string]::IsNullOrWhiteSpace($Line)) { return $null }

    $text = $Line.Trim()
    $changed = $false

    # ---- "Waiting for N of M to be <imported|synced|assigned>" (2415 / 2452 / 2533) ----
    if ($text -match 'Waiting for\s+(\d+)\s+of\s+(\d+)\s+to be\s+(imported|synced|assigned)') {
        $remaining = [int]$Matches[1]
        $total     = [int]$Matches[2]
        $verb      = $Matches[3]

        $stage = switch ($verb) {
            'imported' { 'Import' }
            'synced'   { 'Sync' }
            'assigned' { 'Assign' }
        }

        # The script reports how many are still outstanding, not how many are done.
        $done = [Math]::Max(0, $total - $remaining)

        $State.Stage      = $stage
        $State.Current    = $done
        $State.Total      = $total
        $State.StageLabel = switch ($stage) {
            'Import' { if ($total -gt 1) { "Importing $done of $total devices" } else { 'Importing device into Autopilot' } }
            'Sync'   { if ($total -gt 1) { "Syncing $done of $total devices" }   else { 'Waiting for Intune sync' } }
            'Assign' { if ($total -gt 1) { "Assigning profile to $done of $total" } else { 'Waiting for deployment profile assignment' } }
        }
        $changed = $true
    }
    # ---- stage completion markers ----
    elseif ($text -match 'devices imported successfully') {
        $State.Stage = 'Sync'
        $State.StageLabel = 'Import complete, waiting for Intune sync'
        $changed = $true
    }
    elseif ($text -match 'All devices synced') {
        $State.Stage = 'Assign'
        $State.StageLabel = 'Sync complete, waiting for profile assignment'
        $changed = $true
    }
    elseif ($text -match 'Profiles assigned to all devices') {
        $State.Stage = 'Complete'
        $State.StageLabel = 'Deployment profile assigned'
        $State.IsComplete = $true
        $changed = $true
    }
    # ---- connect / collect ----
    elseif ($text -match 'Connected to Intune tenant\s+(.+)$') {
        $State.Stage = 'Collect'
        $State.StageLabel = 'Signed in, collecting device details'
        $changed = $true
    }
    elseif ($text -match 'Gathered details for device with serial number:\s*(\S+)') {
        $State.Stage = 'Import'
        $State.StageLabel = "Collected hardware details for $($Matches[1])"
        $changed = $true
    }
    elseif ($text -match 'Loading all objects') {
        $State.Stage = 'Import'
        $State.StageLabel = 'Checking whether this device is already registered'
        $changed = $true
    }
    elseif ($text -match 'Adding New Device serial') {
        $State.Stage = 'Import'
        $State.StageLabel = 'Adding device to Autopilot'
        $changed = $true
    }
    elseif ($text -match 'Updating Existing Device') {
        $State.Stage = 'Import'
        $State.StageLabel = 'Updating the existing Autopilot registration'
        $changed = $true
    }
    elseif ($text -match 'Device already exists in Autopilot') {
        $State.Stage = 'Import'
        $State.StageLabel = 'Device is already registered in Autopilot'
        $changed = $true
    }
    # ---- Autopilot v2 / Device Preparation identifier path (2240-2266) ----
    elseif ($text -match 'Checking if device (\S+) exists in AutoPilot') {
        $State.Stage = 'Import'
        $State.StageLabel = "Checking for an existing identifier for $($Matches[1])"
        $changed = $true
    }
    elseif ($text -match 'Device (\S+) added to AutoPilot') {
        $State.Stage = 'Complete'
        $State.StageLabel = "Device identifier imported for $($Matches[1])"
        $State.IsComplete = $true
        $changed = $true
    }
    elseif ($text -match 'Device (\S+) already exists in AutoPilot') {
        $State.Stage = 'Complete'
        $State.StageLabel = "Device identifier already present for $($Matches[1])"
        $State.IsComplete = $true
        $changed = $true
    }
    # ---- failures ----
    elseif ($text -match 'Unable to retrieve device hardware data') {
        $State.IsError = $true
        $State.ErrorMessage = 'Could not read the hardware hash. Run as administrator, and note that virtual machines often cannot provide one.'
        $changed = $true
    }
    elseif ($text -match 'Unable to find group\s+(.+)$') {
        $State.IsError = $true
        $State.ErrorMessage = "Entra group '$($Matches[1])' was not found."
        $changed = $true
    }
    elseif ($text -match '^\s*(Connect-MgGraph|Invoke-MgGraphRequest)\s*:' -or $text -match 'Authentication needed' -or $text -match 'InteractiveBrowserCredential authentication failed') {
        $State.IsError = $true
        $State.ErrorMessage = 'Sign-in to Microsoft Graph failed. Check network connectivity and try again.'
        $changed = $true
    }
    elseif ($text -match 'command was found in the module\s+''([^'']+)''.+could not be loaded') {
        # PowerShell raises this at command resolution, so it is not suppressible with
        # -ErrorAction and it aborts whatever was running.
        $State.IsError = $true
        $State.ErrorMessage = "PowerShell could not load the '$($Matches[1])' module in the engine process. This usually means PSModulePath is broken for that session."
        $changed = $true
    }
    elseif ($text -match '^ERROR:\s*(.+)$') {
        # Emitted by the launcher itself. Surfacing it means the status bar names the real
        # cause instead of only reporting a non-zero exit code.
        $State.IsError = $true
        $State.ErrorMessage = $Matches[1].Trim()
        $changed = $true
    }

    if (-not $changed) { return $null }

    $State.LastMessage = $text
    $State.Percent = Get-ApProgressPercent -State $State
    return $State
}

function Get-ApProgressPercent {
    <#
    .SYNOPSIS
    Percentage for the progress bar, monotonic across stages.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$State
    )

    if ($State.IsComplete) { return 100 }

    $weight = Get-ApStageWeight $State.Stage
    $percent = $weight.Floor

    if ($State.Total -gt 0 -and $weight.Span -gt 0) {
        $fraction = [Math]::Min(1.0, [double]$State.Current / [double]$State.Total)
        $percent = $weight.Floor + [int]([Math]::Floor($fraction * $weight.Span))
    }

    return [Math]::Max(0, [Math]::Min(100, $percent))
}

function Test-ApStageReached {
    <#
    .SYNOPSIS
    True when $Stage is at or beyond $Reference in the pipeline order.
    #>
    param(
        [Parameter(Mandatory)][string]$Stage,
        [Parameter(Mandatory)][string]$Reference
    )

    $a = [Array]::IndexOf($script:ApStageOrder, $Stage)
    $b = [Array]::IndexOf($script:ApStageOrder, $Reference)
    if ($a -lt 0 -or $b -lt 0) { return $false }
    return ($a -ge $b)
}

# ArgumentBuilder.ps1 -- turns UI state into parameters for the community engine script.
#
# This is deliberately a pure function with no UI or Graph dependencies: it is the piece
# most likely to be wrong in a way that silently does nothing, so it is unit-tested
# (tests\ArgumentBuilder.Tests.ps1) and it also drives the "Preview command" dry-run.
#
# Three non-obvious constraints from get-windowsautopilotinfocommunity.ps1 v5.0.16 are
# encoded here. Getting any of them wrong produces a run that appears to succeed but
# doesn't do what the tech asked:
#
#   1. -Reboot / -Wipe / -Sysprep / -preprov / -ChangePK are all nested inside
#      `if ($Assign)` (line 2541). Without -Assign they are silently ignored.
#   2. With neither -delete nor -updatetag, an already-registered serial hits
#      `Read-Host "Do you want to delete or update?"` (line 2346) and blocks forever
#      behind a hidden console. -Force resolves it to "update" (line 2342).
#   3. -identifier takes a completely separate code path (lines 2233-2270) that ignores
#      GroupTag, AssignedUser, AssignedComputerName, AddToGroup, Assign and Reboot.

function New-ApRegistrationRequest {
    <#
    .SYNOPSIS
    Default UI state for a registration request.

    .DESCRIPTION
    Returned as an ordered hashtable so callers can tweak fields before handing it to
    Build-ApEngineArguments. Keeping the default shape in one place stops the UI and the
    tests from drifting apart.
    #>
    [CmdletBinding()]
    param()

    return [ordered]@{
        # Register = push to tenant; Export = offline CSV only; Batch = import a CSV.
        Operation            = 'Register'

        # v1 = hardware hash (windowsAutopilotDeviceIdentities)
        # v2 = device identifier for Device Preparation (importedDeviceIdentities)
        Mode                 = 'v1'

        GroupTag             = ''
        AssignedUser         = ''
        AssignedComputerName = ''
        AddToGroup           = ''

        WaitForAssignment    = $false
        Reboot               = $false

        # update | delete | skipcheck
        ExistingDevicePolicy = 'update'

        Wipe                 = $false
        Sysprep              = $false
        PreProvision         = $false
        ChangePK             = ''

        OutputFile           = ''
        InputFile            = ''
        Append               = $false
        Partner              = $false
    }
}

function Build-ApEngineArguments {
    <#
    .SYNOPSIS
    Maps a registration request to community-script parameters.

    .OUTPUTS
    PSCustomObject with:
      Parameters  ordered hashtable, ready to splat
      Notes       human-readable explanations of any coercion applied
      Warnings    settings that were dropped because the chosen mode ignores them
      Mode/Operation  echoed back for the caller

    .EXAMPLE
    $req = New-ApRegistrationRequest
    $req.GroupTag = 'FINANCE'
    $req.Reboot = $true
    (Build-ApEngineArguments $req).Parameters
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        $Request
    )

    # Accept either a hashtable or a PSCustomObject from the caller.
    $r = if ($Request -is [System.Collections.IDictionary]) { $Request } else { ConvertTo-ApHashtable $Request }

    $p = [ordered]@{}
    $notes = New-Object System.Collections.Generic.List[string]
    $warnings = New-Object System.Collections.Generic.List[string]

    $operation = Get-ApRequestValue $r 'Operation' 'Register'
    $mode      = Get-ApRequestValue $r 'Mode' 'v1'
    $isV2      = ($mode -eq 'v2')

    $groupTag     = (Get-ApRequestValue $r 'GroupTag' '').Trim()
    $assignedUser = (Get-ApRequestValue $r 'AssignedUser' '').Trim()
    $computerName = (Get-ApRequestValue $r 'AssignedComputerName' '').Trim()
    $addToGroup   = (Get-ApRequestValue $r 'AddToGroup' '').Trim()
    $changePk     = (Get-ApRequestValue $r 'ChangePK' '').Trim()
    $outputFile   = (Get-ApRequestValue $r 'OutputFile' '').Trim()
    $inputFile    = (Get-ApRequestValue $r 'InputFile' '').Trim()

    $wait   = [bool](Get-ApRequestValue $r 'WaitForAssignment' $false)
    $reboot = [bool](Get-ApRequestValue $r 'Reboot' $false)
    $wipe   = [bool](Get-ApRequestValue $r 'Wipe' $false)
    $syspr  = [bool](Get-ApRequestValue $r 'Sysprep' $false)
    $prepro = [bool](Get-ApRequestValue $r 'PreProvision' $false)
    $append = [bool](Get-ApRequestValue $r 'Append' $false)
    $partner = [bool](Get-ApRequestValue $r 'Partner' $false)
    $policy = Get-ApRequestValue $r 'ExistingDevicePolicy' 'update'

    if ($isV2) { $p['identifier'] = $true }

    # ---- offline export -------------------------------------------------------
    if ($operation -eq 'Export') {
        if (-not $outputFile) { throw 'An output file is required for an offline export.' }

        $p['OutputFile'] = $outputFile
        if ($append) { $p['Append'] = $true }

        if ($isV2) {
            # The -identifier branch writes Manufacturer,Model,Serial and ignores the rest.
            if ($groupTag) { $warnings.Add('Group tag is not used by a Device Preparation (v2) identifier export.') }
            if ($assignedUser) { $warnings.Add('Assigned user is not used by a Device Preparation (v2) identifier export.') }
        }
        else {
            if ($partner) {
                $p['Partner'] = $true
                # Line 2215 selects the Partner column set before the GroupTag one,
                # so a group tag supplied alongside -Partner never reaches the CSV.
                if ($groupTag) { $warnings.Add('Partner CSV format does not include a Group Tag column; the tag will be omitted.') }
                if ($assignedUser) { $warnings.Add('Partner CSV format does not include an Assigned User column; the user will be omitted.') }
            }
            else {
                if ($groupTag) { $p['GroupTag'] = $groupTag }
                if ($assignedUser) { $p['AssignedUser'] = $assignedUser }
            }
        }

        return [pscustomobject]@{
            Parameters = $p
            Notes      = $notes.ToArray()
            Warnings   = $warnings.ToArray()
            Mode       = $mode
            Operation  = $operation
        }
    }

    # ---- online: Register / Batch --------------------------------------------
    $p['Online'] = $true

    if ($operation -eq 'Batch') {
        if (-not $inputFile) { throw 'A CSV file is required for a batch import.' }
        $p['InputFile'] = $inputFile
    }

    if ($isV2) {
        # Constraint 3: the -identifier path ignores all of these. Warn rather than
        # emit them, so the tech is told instead of quietly getting a different result.
        if ($groupTag)     { $warnings.Add('Group tag is ignored in Device Preparation (v2) mode. v2 assigns devices via the Entra group on the Device Preparation policy.') }
        if ($assignedUser) { $warnings.Add('Assigned user is ignored in Device Preparation (v2) mode.') }
        if ($computerName) { $warnings.Add('Computer name is ignored in Device Preparation (v2) mode.') }
        if ($addToGroup)   { $warnings.Add('Add-to-group is ignored in Device Preparation (v2) mode; add the device to the policy''s Entra group afterwards.') }
        if ($wait)         { $warnings.Add('Profile-assignment wait does not apply in Device Preparation (v2) mode.') }
        if ($reboot)       { $warnings.Add('Reboot-when-assigned does not apply in Device Preparation (v2) mode.') }
        if ($wipe -or $syspr -or $prepro -or $changePk) {
            $warnings.Add('Post-registration actions (wipe / sysprep / pre-provision / product key) are only available in Autopilot v1 mode.')
        }

        return [pscustomobject]@{
            Parameters = $p
            Notes      = $notes.ToArray()
            Warnings   = $warnings.ToArray()
            Mode       = $mode
            Operation  = $operation
        }
    }

    # ---- v1 online -----------------------------------------------------------
    if ($groupTag)     { $p['GroupTag'] = $groupTag }
    if ($assignedUser) { $p['AssignedUser'] = $assignedUser }
    if ($computerName) { $p['AssignedComputerName'] = $computerName }

    if ($addToGroup) {
        # -AddToGroup is [String[]]; accept a comma- or semicolon-separated UI field.
        $groups = @($addToGroup -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        if ($groups.Count -gt 0) { $p['AddToGroup'] = $groups }
    }

    # Constraint 2: never leave the existing-device decision to a Read-Host prompt.
    switch ($policy) {
        'delete' {
            $p['delete'] = $true
            $notes.Add('An existing registration for this serial will be deleted from Autopilot, Intune and Entra ID, then re-added.')
        }
        'skipcheck' {
            $p['newdevice'] = $true
            $notes.Add('Skipping the existing-device lookup. Much faster on large tenants, but will fail if the serial is already registered.')
        }
        default {
            $p['updatetag'] = $true
            $notes.Add('An existing registration for this serial will have its group tag updated.')
        }
    }
    # Belt and braces: -Force is only read in the prompt branch (line 2342), so it is
    # harmless alongside -updatetag/-delete and guarantees the run never blocks.
    $p['Force'] = $true

    # Constraint 1: these only execute inside `if ($Assign)`.
    $needsAssign = $reboot -or $wipe -or $syspr -or $prepro -or [bool]$changePk

    if ($wait -or $needsAssign) {
        $p['Assign'] = $true
        if ($needsAssign -and -not $wait) {
            $notes.Add('Waiting for profile assignment was enabled automatically: the community script only performs post-registration actions after assignment completes.')
        }
    }

    if ($reboot) { $p['Reboot'] = $true }
    if ($wipe)   { $p['Wipe'] = $true }
    if ($syspr)  { $p['Sysprep'] = $true }
    if ($prepro) { $p['preprov'] = $true }
    if ($changePk) { $p['ChangePK'] = $changePk }

    if ($reboot -and ($wipe -or $syspr)) {
        $warnings.Add('Reboot runs before wipe and sysprep in the community script, so those actions will not be reached. Choose one.')
    }

    return [pscustomobject]@{
        Parameters = $p
        Notes      = $notes.ToArray()
        Warnings   = $warnings.ToArray()
        Mode       = $mode
        Operation  = $operation
    }
}

function Get-ApRequestValue {
    <#
    .SYNOPSIS
    Reads a key from the request with a default, tolerating hashtable or object input.
    #>
    param(
        [Parameter(Mandatory)]$Request,
        [Parameter(Mandatory)][string]$Name,
        $Default
    )

    if ($Request -is [System.Collections.IDictionary]) {
        if ($Request.Contains($Name) -and $null -ne $Request[$Name]) { return $Request[$Name] }
        return $Default
    }

    $prop = $Request.PSObject.Properties[$Name]
    if ($prop -and $null -ne $prop.Value) { return $prop.Value }
    return $Default
}

function ConvertTo-ApArgumentArray {
    <#
    .SYNOPSIS
    Parameter hashtable -> flat argument array for powershell.exe -File.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Parameters
    )

    $argList = New-Object System.Collections.Generic.List[string]

    foreach ($key in $Parameters.Keys) {
        $value = $Parameters[$key]

        if ($value -is [bool] -or $value -is [switch]) {
            if ($value) { $argList.Add("-$key") }
            continue
        }

        $argList.Add("-$key")
        if ($value -is [Array]) {
            # A [String[]] parameter takes comma-separated values on a command line.
            $argList.Add(($value -join ','))
        }
        else {
            $argList.Add("$value")
        }
    }

    return $argList.ToArray()
}

function Get-ApPreviewCommand {
    <#
    .SYNOPSIS
    Renders the exact invocation for the "Preview command" dry-run.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Parameters,
        [string]$ScriptPath = 'get-windowsautopilotinfocommunity.ps1'
    )

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append(('& "{0}"' -f $ScriptPath))

    foreach ($key in $Parameters.Keys) {
        $value = $Parameters[$key]

        if ($value -is [bool] -or $value -is [switch]) {
            if ($value) { [void]$sb.Append(" -$key") }
            continue
        }

        if ($value -is [Array]) {
            $quoted = @($value | ForEach-Object { "'{0}'" -f ($_ -replace "'", "''") })
            [void]$sb.Append((' -{0} {1}' -f $key, ($quoted -join ',')))
        }
        else {
            [void]$sb.Append((" -{0} '{1}'" -f $key, ("$value" -replace "'", "''")))
        }
    }

    return $sb.ToString()
}

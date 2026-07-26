# Pester tests for Build-ApEngineArguments.
#
# These lock in the three community-script constraints documented in ArgumentBuilder.ps1.
# All three produce runs that look fine but do the wrong thing, so they are exactly the
# kind of bug a unit test needs to catch -- none of them require a tenant to verify.

BeforeAll {
    $srcRoot = Join-Path (Split-Path -Parent $PSScriptRoot) 'src\Private'
    . (Join-Path $srcRoot 'Logging.ps1')
    . (Join-Path $srcRoot 'Config.ps1')
    . (Join-Path $srcRoot 'ArgumentBuilder.ps1')
}

Describe 'New-ApRegistrationRequest' {
    It 'defaults to an online v1 registration that updates an existing tag' {
        $r = New-ApRegistrationRequest
        $r.Operation | Should -Be 'Register'
        $r.Mode | Should -Be 'v1'
        $r.ExistingDevicePolicy | Should -Be 'update'
    }
}

Describe 'Build-ApEngineArguments: Autopilot v1 registration' {
    It 'emits -Online and an explicit existing-device policy' {
        $r = New-ApRegistrationRequest
        $p = (Build-ApEngineArguments $r).Parameters

        $p['Online'] | Should -BeTrue
        $p['updatetag'] | Should -BeTrue
        # -Force guarantees the run never blocks on the Read-Host at line 2346.
        $p['Force'] | Should -BeTrue
        $p.Contains('identifier') | Should -BeFalse
    }

    It 'passes the group tag, assigned user and computer name through' {
        $r = New-ApRegistrationRequest
        $r.GroupTag = '  FINANCE  '
        $r.AssignedUser = 'user@contoso.com'
        $r.AssignedComputerName = 'FIN-0042'

        $p = (Build-ApEngineArguments $r).Parameters

        # Whitespace is trimmed so a stray space cannot create a second distinct tag.
        $p['GroupTag'] | Should -Be 'FINANCE'
        $p['AssignedUser'] | Should -Be 'user@contoso.com'
        $p['AssignedComputerName'] | Should -Be 'FIN-0042'
    }

    It 'omits empty optional fields entirely' {
        $p = (Build-ApEngineArguments (New-ApRegistrationRequest)).Parameters

        $p.Contains('GroupTag') | Should -BeFalse
        $p.Contains('AssignedUser') | Should -BeFalse
        $p.Contains('AssignedComputerName') | Should -BeFalse
        $p.Contains('AddToGroup') | Should -BeFalse
    }

    It 'splits the add-to-group field on commas and semicolons' {
        $r = New-ApRegistrationRequest
        $r.AddToGroup = 'Autopilot Devices; Finance Laptops ,  Kiosks'

        $p = (Build-ApEngineArguments $r).Parameters

        $p['AddToGroup'] | Should -Be @('Autopilot Devices', 'Finance Laptops', 'Kiosks')
    }
}

Describe 'Build-ApEngineArguments: -Assign coercion (community script line 2541)' {
    It 'adds -Assign when the caller asks to wait for assignment' {
        $r = New-ApRegistrationRequest
        $r.WaitForAssignment = $true

        (Build-ApEngineArguments $r).Parameters['Assign'] | Should -BeTrue
    }

    It 'forces -Assign on when reboot is requested without waiting' {
        $r = New-ApRegistrationRequest
        $r.WaitForAssignment = $false
        $r.Reboot = $true

        $result = Build-ApEngineArguments $r

        # Without -Assign the community script never reaches the Restart-Computer call,
        # so the device would sit at OOBE forever with no error shown.
        $result.Parameters['Assign'] | Should -BeTrue
        $result.Parameters['Reboot'] | Should -BeTrue
        $result.Notes -join ' ' | Should -Match 'enabled automatically'
    }

    It 'forces -Assign on for wipe, sysprep, pre-provision and product key' -ForEach @(
        @{ Field = 'Wipe';         Value = $true;         Emitted = 'Wipe' }
        @{ Field = 'Sysprep';      Value = $true;         Emitted = 'Sysprep' }
        @{ Field = 'PreProvision'; Value = $true;         Emitted = 'preprov' }
        @{ Field = 'ChangePK';     Value = 'XXXXX-XXXXX'; Emitted = 'ChangePK' }
    ) {
        $r = New-ApRegistrationRequest
        $r[$Field] = $Value

        $p = (Build-ApEngineArguments $r).Parameters

        $p['Assign'] | Should -BeTrue
        $p.Contains($Emitted) | Should -BeTrue
    }

    It 'does not add -Assign for a plain registration' {
        (Build-ApEngineArguments (New-ApRegistrationRequest)).Parameters.Contains('Assign') | Should -BeFalse
    }

    It 'warns that reboot pre-empts wipe and sysprep' {
        $r = New-ApRegistrationRequest
        $r.Reboot = $true
        $r.Wipe = $true

        (Build-ApEngineArguments $r).Warnings -join ' ' | Should -Match 'will not be reached'
    }
}

Describe 'Build-ApEngineArguments: existing-device policy' {
    It 'maps "update" to -updatetag' {
        $r = New-ApRegistrationRequest
        $r.ExistingDevicePolicy = 'update'
        $p = (Build-ApEngineArguments $r).Parameters

        $p['updatetag'] | Should -BeTrue
        $p.Contains('delete') | Should -BeFalse
        $p.Contains('newdevice') | Should -BeFalse
    }

    It 'maps "delete" to -delete' {
        $r = New-ApRegistrationRequest
        $r.ExistingDevicePolicy = 'delete'
        $p = (Build-ApEngineArguments $r).Parameters

        $p['delete'] | Should -BeTrue
        $p.Contains('updatetag') | Should -BeFalse
    }

    It 'maps "skipcheck" to -newdevice and explains the trade-off' {
        $r = New-ApRegistrationRequest
        $r.ExistingDevicePolicy = 'skipcheck'
        $result = Build-ApEngineArguments $r

        $result.Parameters['newdevice'] | Should -BeTrue
        $result.Parameters.Contains('updatetag') | Should -BeFalse
        $result.Notes -join ' ' | Should -Match 'large tenants'
    }

    It 'always includes -Force regardless of policy' -ForEach @(
        @{ Policy = 'update' }, @{ Policy = 'delete' }, @{ Policy = 'skipcheck' }
    ) {
        $r = New-ApRegistrationRequest
        $r.ExistingDevicePolicy = $Policy

        (Build-ApEngineArguments $r).Parameters['Force'] | Should -BeTrue
    }
}

Describe 'Build-ApEngineArguments: Device Preparation (v2)' {
    It 'emits -identifier and -Online only' {
        $r = New-ApRegistrationRequest
        $r.Mode = 'v2'

        $p = (Build-ApEngineArguments $r).Parameters

        $p['identifier'] | Should -BeTrue
        $p['Online'] | Should -BeTrue
        # The -identifier branch (lines 2233-2270) never consults these.
        $p.Contains('GroupTag') | Should -BeFalse
        $p.Contains('Assign') | Should -BeFalse
        $p.Contains('Reboot') | Should -BeFalse
        $p.Contains('updatetag') | Should -BeFalse
        $p.Contains('Force') | Should -BeFalse
    }

    It 'warns instead of silently dropping v1-only settings' {
        $r = New-ApRegistrationRequest
        $r.Mode = 'v2'
        $r.GroupTag = 'FINANCE'
        $r.AssignedUser = 'user@contoso.com'
        $r.AddToGroup = 'Autopilot Devices'
        $r.Reboot = $true

        $result = Build-ApEngineArguments $r
        $warnings = $result.Warnings -join ' '

        $warnings | Should -Match 'Group tag is ignored'
        $warnings | Should -Match 'Assigned user is ignored'
        $warnings | Should -Match 'Entra group'
        $result.Parameters.Contains('GroupTag') | Should -BeFalse
    }
}

Describe 'Build-ApEngineArguments: offline export' {
    It 'writes a CSV without connecting to the tenant' {
        $r = New-ApRegistrationRequest
        $r.Operation = 'Export'
        $r.OutputFile = 'C:\Temp\AutopilotHWID.csv'

        $p = (Build-ApEngineArguments $r).Parameters

        $p['OutputFile'] | Should -Be 'C:\Temp\AutopilotHWID.csv'
        $p.Contains('Online') | Should -BeFalse
    }

    It 'includes the group tag and assigned user columns when supplied' {
        $r = New-ApRegistrationRequest
        $r.Operation = 'Export'
        $r.OutputFile = 'C:\Temp\hwid.csv'
        $r.GroupTag = 'FINANCE'
        $r.AssignedUser = 'user@contoso.com'

        $p = (Build-ApEngineArguments $r).Parameters

        $p['GroupTag'] | Should -Be 'FINANCE'
        $p['AssignedUser'] | Should -Be 'user@contoso.com'
    }

    It 'exports a v2 identifier CSV with -identifier' {
        $r = New-ApRegistrationRequest
        $r.Operation = 'Export'
        $r.Mode = 'v2'
        $r.OutputFile = 'C:\Temp\identifiers.csv'

        $p = (Build-ApEngineArguments $r).Parameters

        $p['identifier'] | Should -BeTrue
        $p['OutputFile'] | Should -Be 'C:\Temp\identifiers.csv'
        $p.Contains('Online') | Should -BeFalse
    }

    It 'warns that the Partner format drops the group tag column' {
        $r = New-ApRegistrationRequest
        $r.Operation = 'Export'
        $r.OutputFile = 'C:\Temp\partner.csv'
        $r.Partner = $true
        $r.GroupTag = 'FINANCE'

        $result = Build-ApEngineArguments $r

        $result.Parameters['Partner'] | Should -BeTrue
        # Line 2215 picks the Partner column set before the GroupTag one.
        $result.Parameters.Contains('GroupTag') | Should -BeFalse
        $result.Warnings -join ' ' | Should -Match 'Partner CSV format'
    }

    It 'passes -Append through' {
        $r = New-ApRegistrationRequest
        $r.Operation = 'Export'
        $r.OutputFile = 'C:\Temp\hwid.csv'
        $r.Append = $true

        (Build-ApEngineArguments $r).Parameters['Append'] | Should -BeTrue
    }

    It 'requires an output file' {
        $r = New-ApRegistrationRequest
        $r.Operation = 'Export'

        { Build-ApEngineArguments $r } | Should -Throw '*output file is required*'
    }
}

Describe 'Build-ApEngineArguments: batch import' {
    It 'passes the CSV as -InputFile alongside -Online' {
        $r = New-ApRegistrationRequest
        $r.Operation = 'Batch'
        $r.InputFile = 'C:\Temp\devices.csv'

        $p = (Build-ApEngineArguments $r).Parameters

        $p['InputFile'] | Should -Be 'C:\Temp\devices.csv'
        $p['Online'] | Should -BeTrue
    }

    It 'requires a CSV file' {
        $r = New-ApRegistrationRequest
        $r.Operation = 'Batch'

        { Build-ApEngineArguments $r } | Should -Throw '*CSV file is required*'
    }
}

Describe 'ConvertTo-ApArgumentArray' {
    It 'renders switches without a value and strings with one' {
        $p = [ordered]@{ Online = $true; GroupTag = 'FINANCE'; Force = $true }

        ConvertTo-ApArgumentArray $p | Should -Be @('-Online', '-GroupTag', 'FINANCE', '-Force')
    }

    It 'omits switches that are false' {
        ConvertTo-ApArgumentArray ([ordered]@{ Online = $true; Reboot = $false }) | Should -Be @('-Online')
    }

    It 'joins array values with commas for [String[]] parameters' {
        $p = [ordered]@{ AddToGroup = @('Group A', 'Group B') }

        ConvertTo-ApArgumentArray $p | Should -Be @('-AddToGroup', 'Group A,Group B')
    }
}

Describe 'Get-ApPreviewCommand' {
    It 'renders a copy-pasteable invocation' {
        $p = [ordered]@{ Online = $true; GroupTag = 'FINANCE'; Assign = $true; Reboot = $true }

        Get-ApPreviewCommand -Parameters $p -ScriptPath 'C:\vendor\engine.ps1' |
            Should -Be "& ""C:\vendor\engine.ps1"" -Online -GroupTag 'FINANCE' -Assign -Reboot"
    }

    It 'escapes single quotes in values' {
        $p = [ordered]@{ AssignedComputerName = "O'Brien-PC" }

        Get-ApPreviewCommand -Parameters $p -ScriptPath 'e.ps1' |
            Should -Be "& ""e.ps1"" -AssignedComputerName 'O''Brien-PC'"
    }
}

# Pester tests for the progress parser.
#
# The sample lines below are the literal strings emitted by
# get-windowsautopilotinfocommunity.ps1 v5.0.16 (line numbers in comments). If upstream
# rewords them these tests fail, which is the intent -- it tells us the progress bar has
# gone blind before a tech notices in the field.

BeforeAll {
    $srcRoot = Join-Path (Split-Path -Parent $PSScriptRoot) 'src\Private'
    . (Join-Path $srcRoot 'ProgressParser.ps1')
}

Describe 'New-ApProgressState' {
    It 'starts at the Connect stage with nothing complete' {
        $s = New-ApProgressState

        $s.Stage | Should -Be 'Connect'
        $s.Percent | Should -Be 0
        $s.IsComplete | Should -BeFalse
        $s.IsError | Should -BeFalse
    }
}

Describe 'Update-ApProgressState: unrecognised input' {
    It 'returns $null for a line it does not understand' {
        $s = New-ApProgressState

        Update-ApProgressState -State $s -Line 'Some unrelated verbose chatter' | Should -BeNullOrEmpty
    }

    It 'returns $null for blank lines' {
        $s = New-ApProgressState

        Update-ApProgressState -State $s -Line '' | Should -BeNullOrEmpty
        Update-ApProgressState -State $s -Line '   ' | Should -BeNullOrEmpty
    }

    It 'leaves the state untouched when nothing matches' {
        $s = New-ApProgressState
        Update-ApProgressState -State $s -Line 'noise' | Out-Null

        $s.Stage | Should -Be 'Connect'
        $s.Percent | Should -Be 0
    }
}

Describe 'Update-ApProgressState: sign-in and collection' {
    It 'advances to Collect once the tenant connection is reported' {
        # Line 2077: "Connected to Intune tenant <guid>"
        $s = New-ApProgressState
        Update-ApProgressState -State $s -Line 'Connected to Intune tenant 00000000-1111-2222-3333-444444444444' | Out-Null

        $s.Stage | Should -Be 'Collect'
        $s.Percent | Should -BeGreaterOrEqual 10
    }

    It 'advances to Import once hardware details are gathered' {
        # Line 2196: "Gathered details for device with serial number: <serial>"
        $s = New-ApProgressState
        Update-ApProgressState -State $s -Line 'Gathered details for device with serial number: 5CG1234ABC' | Out-Null

        $s.Stage | Should -Be 'Import'
        $s.StageLabel | Should -Match '5CG1234ABC'
    }

    It 'reports the slow full-tenant lookup in plain language' {
        # Line 2290: "Loading all objects. This can take a while on large tenants"
        $s = New-ApProgressState
        Update-ApProgressState -State $s -Line 'Loading all objects. This can take a while on large tenants' | Out-Null

        $s.Stage | Should -Be 'Import'
        $s.StageLabel | Should -Match 'already registered'
    }
}

Describe 'Update-ApProgressState: "Waiting for N of M" lines' {
    It 'converts outstanding counts into completed counts' {
        # Line 2415 reports how many REMAIN, not how many are done.
        $s = New-ApProgressState
        Update-ApProgressState -State $s -Line 'Waiting for 2 of 5 to be imported' | Out-Null

        $s.Stage | Should -Be 'Import'
        $s.Total | Should -Be 5
        $s.Current | Should -Be 3
    }

    It 'maps each verb to its stage' -ForEach @(
        @{ Line = 'Waiting for 1 of 1 to be imported'; Stage = 'Import' }
        @{ Line = 'Waiting for 1 of 1 to be synced';   Stage = 'Sync' }
        @{ Line = 'Waiting for 1 of 1 to be assigned'; Stage = 'Assign' }
    ) {
        $s = New-ApProgressState
        Update-ApProgressState -State $s -Line $Line | Out-Null

        $s.Stage | Should -Be $Stage
    }

    It 'uses singular wording for a single device' {
        $s = New-ApProgressState
        Update-ApProgressState -State $s -Line 'Waiting for 1 of 1 to be assigned' | Out-Null

        $s.StageLabel | Should -Be 'Waiting for deployment profile assignment'
    }

    It 'uses counted wording for a batch' {
        $s = New-ApProgressState
        Update-ApProgressState -State $s -Line 'Waiting for 3 of 10 to be imported' | Out-Null

        $s.StageLabel | Should -Be 'Importing 7 of 10 devices'
    }

    It 'never reports a negative completed count' {
        $s = New-ApProgressState
        Update-ApProgressState -State $s -Line 'Waiting for 5 of 2 to be imported' | Out-Null

        $s.Current | Should -Be 0
    }
}

Describe 'Update-ApProgressState: stage completion' {
    It 'moves Import -> Sync -> Assign -> Complete in order' {
        $s = New-ApProgressState

        Update-ApProgressState -State $s -Line '1 devices imported successfully. Elapsed time to complete import: 42 seconds' | Out-Null
        $s.Stage | Should -Be 'Sync'

        Update-ApProgressState -State $s -Line 'All devices synced. Elapsed time to complete sync: 30 seconds' | Out-Null
        $s.Stage | Should -Be 'Assign'

        Update-ApProgressState -State $s -Line 'Profiles assigned to all devices. Elapsed time to complete assignment: 90 seconds' | Out-Null
        $s.Stage | Should -Be 'Complete'
        $s.IsComplete | Should -BeTrue
        $s.Percent | Should -Be 100
    }

    It 'increases the percentage monotonically through a full run' {
        $s = New-ApProgressState
        $lines = @(
            'Connected to Intune tenant 00000000-1111-2222-3333-444444444444'
            'Gathered details for device with serial number: 5CG1234ABC'
            'Waiting for 1 of 1 to be imported'
            '1 devices imported successfully. Elapsed time to complete import: 42 seconds'
            'Waiting for 1 of 1 to be synced'
            'All devices synced. Elapsed time to complete sync: 30 seconds'
            'Waiting for 1 of 1 to be assigned'
            'Profiles assigned to all devices. Elapsed time to complete assignment: 90 seconds'
        )

        $last = -1
        foreach ($line in $lines) {
            Update-ApProgressState -State $s -Line $line | Out-Null
            $s.Percent | Should -BeGreaterOrEqual $last
            $last = $s.Percent
        }

        $s.Percent | Should -Be 100
    }
}

Describe 'Update-ApProgressState: Device Preparation (v2) identifier path' {
    It 'reports the existence check' {
        # Line 2240
        $s = New-ApProgressState
        Update-ApProgressState -State $s -Line 'Checking if device 5CG1234ABC exists in AutoPilot' | Out-Null

        $s.Stage | Should -Be 'Import'
        $s.StageLabel | Should -Match '5CG1234ABC'
    }

    It 'completes when the identifier is added' {
        # Line 2245
        $s = New-ApProgressState
        Update-ApProgressState -State $s -Line 'Device 5CG1234ABC added to AutoPilot' | Out-Null

        $s.IsComplete | Should -BeTrue
        $s.Percent | Should -Be 100
    }

    It 'completes when the identifier is already present' {
        # Line 2248 -- already registered is a success, not a failure.
        $s = New-ApProgressState
        Update-ApProgressState -State $s -Line 'Device 5CG1234ABC already exists in AutoPilot' | Out-Null

        $s.IsComplete | Should -BeTrue
        $s.IsError | Should -BeFalse
    }
}

Describe 'Update-ApProgressState: failures' {
    It 'flags a missing hardware hash with actionable advice' {
        # Line 2189
        $s = New-ApProgressState
        Update-ApProgressState -State $s -Line 'Unable to retrieve device hardware data (hash) from computer localhost' | Out-Null

        $s.IsError | Should -BeTrue
        $s.ErrorMessage | Should -Match 'administrator'
    }

    It 'flags a missing Entra group by name' {
        # Line 2508
        $s = New-ApProgressState
        Update-ApProgressState -State $s -Line 'Unable to find group Finance Laptops' | Out-Null

        $s.IsError | Should -BeTrue
        $s.ErrorMessage | Should -Match 'Finance Laptops'
    }

    It 'flags a Graph sign-in failure' {
        $s = New-ApProgressState
        Update-ApProgressState -State $s -Line 'Connect-MgGraph : Authentication needed, call Connect-MgGraph.' | Out-Null

        $s.IsError | Should -BeTrue
        $s.ErrorMessage | Should -Match 'Sign-in'
    }

    It 'names the module when PowerShell cannot load one' {
        # Regression: this exact failure aborted a real run before sign-in, and the GUI only
        # reported "exit code 1" while the reason sat unread in the output pane.
        $s = New-ApProgressState
        Update-ApProgressState -State $s -Line "The 'Set-ExecutionPolicy' command was found in the module 'Microsoft.PowerShell.Security', but the module could not be loaded." | Out-Null

        $s.IsError | Should -BeTrue
        $s.ErrorMessage | Should -Match 'Microsoft\.PowerShell\.Security'
        $s.ErrorMessage | Should -Match 'PSModulePath'
    }

    It 'surfaces a launcher ERROR line verbatim' {
        $s = New-ApProgressState
        Update-ApProgressState -State $s -Line 'ERROR: the engine script is missing: C:\nope\engine.ps1' | Out-Null

        $s.IsError | Should -BeTrue
        $s.ErrorMessage | Should -Be 'the engine script is missing: C:\nope\engine.ps1'
    }
}

Describe 'Test-ApStageReached' {
    It 'treats a stage as having reached itself' {
        Test-ApStageReached -Stage 'Sync' -Reference 'Sync' | Should -BeTrue
    }

    It 'compares stages in pipeline order' {
        Test-ApStageReached -Stage 'Assign' -Reference 'Import' | Should -BeTrue
        Test-ApStageReached -Stage 'Import' -Reference 'Assign' | Should -BeFalse
    }

    It 'returns false for unknown stage names' {
        Test-ApStageReached -Stage 'Nonsense' -Reference 'Import' | Should -BeFalse
    }
}

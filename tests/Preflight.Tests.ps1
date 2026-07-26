# Pester tests for the PSModulePath repair.
#
# Regression cover for a reported field failure. Launching the GUI from a PowerShell 7
# terminal leaked PS7's PSModulePath into the Windows PowerShell engine process. The PS7
# entries sort first and shadow in-box modules with Core-only copies, which produced two
# different symptoms from one cause and aborted the run before any sign-in prompt appeared:
#
#   The 'Set-ExecutionPolicy' command was found in the module 'Microsoft.PowerShell.Security',
#   but the module could not be loaded.
#
#   The term 'Get-PackageProvider' is not recognized as the name of a cmdlet...
#
# The leaked value below is copied verbatim from the run log that reproduced it.

BeforeAll {
    $srcRoot = Join-Path (Split-Path -Parent $PSScriptRoot) 'src\Private'
    . (Join-Path $srcRoot 'Logging.ps1')
    . (Join-Path $srcRoot 'Preflight.ps1')

    $script:LeakedPath = @(
        'C:\Users\someone\OneDrive\Documents\PowerShell\Modules'
        'C:\Program Files\PowerShell\Modules'
        'c:\program files\windowsapps\microsoft.powershell_7.6.4.0_x64__8wekyb3d8bbwe\Modules'
        'C:\Program Files\WindowsPowerShell\Modules'
        'C:\WINDOWS\system32\WindowsPowerShell\v1.0\Modules'
    ) -join ';'
}

Describe 'Get-ApWindowsPowerShellModulePath' {

    BeforeEach {
        $script:OriginalModulePath = $env:PSModulePath
    }

    AfterEach {
        $env:PSModulePath = $script:OriginalModulePath
    }

    It 'always includes the three canonical Windows PowerShell locations' {
        $env:PSModulePath = ''
        $result = (Get-ApWindowsPowerShellModulePath) -split ';'

        $result | Should -Contain (Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules')
        $result | Should -Contain (Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\Modules')
        @($result | Where-Object { $_ -like '*\WindowsPowerShell\Modules' }).Count | Should -BeGreaterOrEqual 2
    }

    It 'removes the PowerShell 7 user module path' {
        $env:PSModulePath = $script:LeakedPath

        (Get-ApWindowsPowerShellModulePath) -split ';' |
            Should -Not -Contain 'C:\Users\someone\OneDrive\Documents\PowerShell\Modules'
    }

    It 'removes the PowerShell 7 all-users module path' {
        $env:PSModulePath = $script:LeakedPath

        (Get-ApWindowsPowerShellModulePath) -split ';' |
            Should -Not -Contain 'C:\Program Files\PowerShell\Modules'
    }

    It 'removes the Microsoft Store PowerShell 7 module path' {
        $env:PSModulePath = $script:LeakedPath

        # This one has no "\PowerShell\" segment, so a naive path-segment filter misses it.
        @((Get-ApWindowsPowerShellModulePath) -split ';' | Where-Object { $_ -like '*windowsapps*' }).Count |
            Should -Be 0
    }

    It 'keeps every genuine WindowsPowerShell path from the leaked value' {
        $env:PSModulePath = $script:LeakedPath
        $result = (Get-ApWindowsPowerShellModulePath) -split ';'

        $result | Should -Contain 'C:\Program Files\WindowsPowerShell\Modules'
        @($result | Where-Object { $_ -like '*system32\WindowsPowerShell\v1.0\Modules' }).Count |
            Should -BeGreaterOrEqual 1
    }

    It 'retains unrelated custom module paths' {
        # A management agent's module directory mentions neither PowerShell edition and must
        # survive, otherwise this repair breaks working setups.
        $env:PSModulePath = "C:\CorpTools\PSModules;$script:LeakedPath"

        (Get-ApWindowsPowerShellModulePath) -split ';' | Should -Contain 'C:\CorpTools\PSModules'
    }

    It 'puts the Windows PowerShell paths before any inherited ones' {
        $env:PSModulePath = "C:\CorpTools\PSModules;$script:LeakedPath"
        $result = @((Get-ApWindowsPowerShellModulePath) -split ';')

        # Shadowing is an ordering problem, so ordering is part of the contract.
        $systemIndex = [Array]::FindIndex([string[]]$result, [Predicate[string]] { $args[0] -like '*System32\WindowsPowerShell\v1.0\Modules' })
        $customIndex = [Array]::IndexOf($result, 'C:\CorpTools\PSModules')

        $systemIndex | Should -BeLessThan $customIndex
    }

    It 'de-duplicates case-insensitively' {
        # System32 and system32 both appear in practice, from different sources.
        $env:PSModulePath = @(
            'C:\WINDOWS\System32\WindowsPowerShell\v1.0\Modules'
            'C:\WINDOWS\system32\windowspowershell\v1.0\modules'
        ) -join ';'

        $result = @((Get-ApWindowsPowerShellModulePath) -split ';')
        @($result | Where-Object { $_ -like '*windowspowershell\v1.0\modules' }).Count | Should -Be 1
    }

    It 'ignores empty entries and trailing separators' {
        $env:PSModulePath = 'C:\CorpTools\PSModules;;;'
        $result = @((Get-ApWindowsPowerShellModulePath) -split ';')

        @($result | Where-Object { -not $_ }).Count | Should -Be 0
    }

    It 'is idempotent' {
        $env:PSModulePath = $script:LeakedPath
        $once = Get-ApWindowsPowerShellModulePath

        $env:PSModulePath = $once
        Get-ApWindowsPowerShellModulePath | Should -Be $once
    }
}

Describe 'Repair-ApModulePath' {

    BeforeEach {
        $script:OriginalModulePath = $env:PSModulePath
    }

    AfterEach {
        $env:PSModulePath = $script:OriginalModulePath
    }

    It 'rewrites the environment variable in place' {
        $env:PSModulePath = $script:LeakedPath

        $returned = Repair-ApModulePath

        $env:PSModulePath | Should -Be $returned
        $env:PSModulePath | Should -Not -Match 'windowsapps'
    }

    It 'leaves an already-clean path untouched' {
        $clean = Get-ApWindowsPowerShellModulePath
        $env:PSModulePath = $clean

        Repair-ApModulePath | Should -Be $clean
    }
}

Describe 'Get-ApGraphModuleAdvice' {

    It 'returns nothing when the module is fine' {
        Get-ApGraphModuleAdvice -Check ([pscustomobject]@{ Available = $true }) | Should -BeNullOrEmpty
    }

    It 'explains a missing module' {
        $check = [pscustomobject]@{
            Available = $false; InstalledVersions = @(); Version = ''; Path = ''; Error = ''; WorkingVersion = ''
        }

        Get-ApGraphModuleAdvice -Check $check | Should -Match 'not installed'
    }

    It 'reports the version, path and underlying error for a broken module' {
        $check = [pscustomobject]@{
            Available         = $false
            InstalledVersions = @('2.29.0')
            Version           = '2.29.0'
            Path              = 'C:\Users\someone\OneDrive\Documents\PowerShell\Modules\microsoft.graph.authentication\2.29.0'
            Error             = "The member 'FormatsToProcess' in the module manifest is not valid"
            WorkingVersion    = ''
        }

        $advice = Get-ApGraphModuleAdvice -Check $check
        $advice | Should -Match '2\.29\.0'
        $advice | Should -Match 'FormatsToProcess'
        # The path is inside OneDrive, which is worth calling out explicitly.
        $advice | Should -Match 'OneDrive'
    }

    It 'explains that the highest version wins when a working one is also present' {
        $check = [pscustomobject]@{
            Available         = $false
            InstalledVersions = @('2.29.0', '2.9.1')
            Version           = '2.29.0'
            Path              = 'C:\Modules\microsoft.graph.authentication\2.29.0'
            Error             = 'nope'
            WorkingVersion    = '2.9.1'
        }

        Get-ApGraphModuleAdvice -Check $check | Should -Match 'highest version'
    }
}


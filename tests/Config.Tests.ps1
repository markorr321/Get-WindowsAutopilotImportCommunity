# Pester tests for configuration load/save.
#
# The group-tag history round-trip is the interesting case: an empty PowerShell array inside
# an ordered dictionary serialises to "{}" rather than "[]", and ConvertFrom-Json turns that
# back into a dictionary. Unnormalised, that surfaced as a single bogus
# "System.Collections.Specialized.OrderedDictionary" entry in the group tag dropdown.

BeforeAll {
    $srcRoot = Join-Path (Split-Path -Parent $PSScriptRoot) 'src\Private'
    . (Join-Path $srcRoot 'Logging.ps1')
    . (Join-Path $srcRoot 'Config.ps1')
}

# Pester 6 does not allow BeforeEach/AfterEach directly in the container, so everything is
# nested inside one Describe.
Describe 'Configuration' {

BeforeEach {
    # Point config resolution at a scratch directory so tests never touch a real config.
    $script:TestRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("ap-cfg-{0}" -f ([guid]::NewGuid().ToString('N')))
    New-Item -ItemType Directory -Path $script:TestRoot -Force | Out-Null
    $script:ApAppRoot = $script:TestRoot
    $script:ApConfig = $null
    $script:ApConfigPath = $null
}

AfterEach {
    Remove-Item -LiteralPath $script:TestRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Context 'Get-ApDefaultConfig' {
    It 'defaults to an online v1 update flow with the console hidden' {
        $c = Get-ApDefaultConfig

        $c.lastMode | Should -Be 'v1'
        $c.existingDevicePolicy | Should -Be 'update'
        $c.showConsoleWindow | Should -BeFalse
        $c.connectivityEndpoints | Should -BeNullOrEmpty
    }
}

Context 'Import-ApConfig' {
    It 'returns defaults when no file exists' {
        (Import-ApConfig).lastMode | Should -Be 'v1'
    }

    It 'reads values from config.json next to the script' {
        @{ lastGroupTag = 'FINANCE'; lastMode = 'v2' } | ConvertTo-Json |
            Set-Content (Join-Path $script:TestRoot 'config.json') -Encoding UTF8

        $c = Import-ApConfig
        $c.lastGroupTag | Should -Be 'FINANCE'
        $c.lastMode | Should -Be 'v2'
    }

    It 'ignores keys it does not recognise' {
        '{ "lastGroupTag": "OK", "somethingInjected": "bad" }' |
            Set-Content (Join-Path $script:TestRoot 'config.json') -Encoding UTF8

        $c = Import-ApConfig
        $c.lastGroupTag | Should -Be 'OK'
        $c.Contains('somethingInjected') | Should -BeFalse
    }

    It 'falls back to defaults on malformed JSON rather than throwing' {
        'this is not json {{{' | Set-Content (Join-Path $script:TestRoot 'config.json') -Encoding UTF8

        { Import-ApConfig } | Should -Not -Throw
        (Get-ApConfig).lastMode | Should -Be 'v1'
    }

    It 'normalises a history saved as an empty object to an empty array' {
        '{ "groupTagHistory": {} }' | Set-Content (Join-Path $script:TestRoot 'config.json') -Encoding UTF8

        $history = (Import-ApConfig).groupTagHistory

        # The bug this guards: an unnormalised {} became one dictionary entry in the dropdown.
        @($history).Count | Should -Be 0
    }

    It 'keeps only non-empty strings in the history' {
        '{ "groupTagHistory": ["FINANCE", "", "  ", "KIOSK", "FINANCE"] }' |
            Set-Content (Join-Path $script:TestRoot 'config.json') -Encoding UTF8

        (Import-ApConfig).groupTagHistory | Should -Be @('FINANCE', 'KIOSK')
    }
}

Context 'Add-ApGroupTagToHistory' {
    It 'adds a tag as most recently used' {
        Import-ApConfig | Out-Null
        Add-ApGroupTagToHistory 'FINANCE'

        (Get-ApConfig).groupTagHistory[0] | Should -Be 'FINANCE'
    }

    It 'moves an existing tag to the front instead of duplicating it' {
        Import-ApConfig | Out-Null
        Add-ApGroupTagToHistory 'A'
        Add-ApGroupTagToHistory 'B'
        Add-ApGroupTagToHistory 'A'

        (Get-ApConfig).groupTagHistory | Should -Be @('A', 'B')
    }

    It 'trims whitespace so a stray space does not create a second tag' {
        Import-ApConfig | Out-Null
        Add-ApGroupTagToHistory 'FINANCE'
        Add-ApGroupTagToHistory '  FINANCE  '

        @((Get-ApConfig).groupTagHistory).Count | Should -Be 1
    }

    It 'ignores blank input' {
        Import-ApConfig | Out-Null
        Add-ApGroupTagToHistory ''
        Add-ApGroupTagToHistory '   '

        @((Get-ApConfig).groupTagHistory).Count | Should -Be 0
    }

    It 'caps the history at 25 entries' {
        Import-ApConfig | Out-Null
        1..30 | ForEach-Object { Add-ApGroupTagToHistory "TAG-$_" }

        @((Get-ApConfig).groupTagHistory).Count | Should -Be 25
        (Get-ApConfig).groupTagHistory[0] | Should -Be 'TAG-30'
    }
}

Context 'Set-ApConfigValue' {
    It 'sets a known key' {
        Import-ApConfig | Out-Null
        Set-ApConfigValue 'lastGroupTag' 'KIOSK'

        (Get-ApConfig).lastGroupTag | Should -Be 'KIOSK'
    }

    It 'rejects an unknown key' {
        Import-ApConfig | Out-Null

        { Set-ApConfigValue 'nonsense' 'x' } | Should -Throw '*Unknown configuration key*'
    }
}

Context 'Save-ApConfig round-trip' {
    It 'persists values and reloads them' {
        Import-ApConfig | Out-Null
        Set-ApConfigValue 'lastGroupTag' 'FINANCE'
        Set-ApConfigValue 'lastMode' 'v2'
        Add-ApGroupTagToHistory 'FINANCE'
        Add-ApGroupTagToHistory 'KIOSK'

        $saved = Save-ApConfig
        $saved | Should -Not -BeNullOrEmpty
        Test-Path -LiteralPath $saved | Should -BeTrue

        # Reload from scratch.
        $script:ApConfig = $null
        $script:ApConfigPath = $null
        $reloaded = Import-ApConfig

        $reloaded.lastGroupTag | Should -Be 'FINANCE'
        $reloaded.lastMode | Should -Be 'v2'
        $reloaded.groupTagHistory | Should -Be @('KIOSK', 'FINANCE')
    }

    It 'survives a round-trip with an empty history' {
        Import-ApConfig | Out-Null
        Save-ApConfig | Out-Null

        $script:ApConfig = $null
        $script:ApConfigPath = $null

        @((Import-ApConfig).groupTagHistory).Count | Should -Be 0
    }
}

}

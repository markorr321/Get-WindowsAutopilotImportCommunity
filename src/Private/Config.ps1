# Config.ps1 -- persisted user preferences (config.json).
#
# Resolution order for the config file:
#   1. next to the script  -- so a USB stick can carry a preconfigured config.json
#      with the site's group tags already populated
#   2. %ProgramData%\AutopilotImportGUI\config.json
# Reads prefer (1); writes go to whichever is writable, preferring (1).

$script:ApConfig = $null
$script:ApConfigPath = $null

function Get-ApDefaultConfig {
    <#
    .SYNOPSIS
    The shape and defaults of config.json. Anything absent from the file on disk
    falls back to these values.
    #>
    return [ordered]@{
        schemaVersion         = 1

        # Editable ComboBox history for the Group Tag field. The original GUI had a
        # bare TextBox, so techs retyped (and mistyped) the same tags all day.
        groupTagHistory       = @()

        # Last-used registration options, restored on next launch.
        lastMode              = 'v1'          # v1 | v2
        lastGroupTag          = ''
        lastAssignedUser      = ''
        lastComputerName      = ''
        lastAddToGroup        = ''
        waitForAssignment     = $true
        rebootWhenAssigned    = $true

        # Device Preparation (v2) restart-after-import. Off by default: a restart straight out
        # of OOBE is premature unless the device is already in the policy's Entra group.
        rebootAfterV2Import   = $false
        existingDevicePolicy  = 'update'      # update | delete
        confirmBeforeRegister = $true

        # Show the child PowerShell console instead of hiding it. Off by default --
        # the whole point of this rewrite is that output lands in the GUI.
        showConsoleWindow     = $false

        # $null = use the built-in list from Get-ApDefaultEndpoints.
        connectivityEndpoints = $null

        # Prefer the vendored script over any PSGallery-installed copy.
        preferVendoredScript  = $true
    }
}

function Get-ApConfigCandidatePaths {
    $paths = New-Object System.Collections.Generic.List[string]

    if ($script:ApAppRoot) {
        $paths.Add((Join-Path $script:ApAppRoot 'config.json'))
    }
    if ($env:ProgramData) {
        $paths.Add((Join-Path $env:ProgramData 'AutopilotImportGUI\config.json'))
    }

    return $paths.ToArray()
}

function ConvertTo-ApHashtable {
    <#
    .SYNOPSIS
    PSCustomObject (from ConvertFrom-Json) -> ordered hashtable, recursively.
    Windows PowerShell 5.1 has no ConvertFrom-Json -AsHashtable.
    #>
    param($InputObject)

    if ($null -eq $InputObject) { return $null }

    if ($InputObject -is [System.Management.Automation.PSCustomObject]) {
        $out = [ordered]@{}
        foreach ($p in $InputObject.PSObject.Properties) {
            $out[$p.Name] = ConvertTo-ApHashtable $p.Value
        }
        return $out
    }

    if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
        return @($InputObject | ForEach-Object { ConvertTo-ApHashtable $_ })
    }

    return $InputObject
}

function Import-ApConfig {
    <#
    .SYNOPSIS
    Loads config.json, merged over the defaults. Never throws.
    #>
    [CmdletBinding()]
    param()

    $config = Get-ApDefaultConfig

    foreach ($path in (Get-ApConfigCandidatePaths)) {
        if (-not (Test-Path -LiteralPath $path)) { continue }
        try {
            $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
            if ([string]::IsNullOrWhiteSpace($raw)) { continue }

            $loaded = ConvertTo-ApHashtable (ConvertFrom-Json $raw -ErrorAction Stop)
            foreach ($key in $loaded.Keys) {
                # Only accept keys we know about, so a stale file cannot inject junk.
                if ($config.Contains($key)) { $config[$key] = $loaded[$key] }
            }

            $script:ApConfigPath = $path
            Write-ApLog "Loaded configuration from $path"
            break
        }
        catch {
            Write-ApLog "Ignoring unreadable config at ${path}: $($_.Exception.Message)" -Level WARN
        }
    }

    if (-not $script:ApConfigPath) {
        Write-ApLog 'No config.json found; using defaults.'
    }

    # An empty PowerShell array serialises to "{}" rather than "[]" inside an ordered
    # dictionary, and ConvertFrom-Json turns that back into a dictionary. Left alone it
    # became a single bogus "System.Collections.Specialized.OrderedDictionary" entry in the
    # group tag dropdown, so normalise to a clean string array.
    $config.groupTagHistory = @(
        @($config.groupTagHistory) |
            Where-Object { $_ -is [string] -and -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { $_.Trim() } |
            Select-Object -Unique
    )

    $script:ApConfig = $config
    return $config
}

function Get-ApConfig {
    if (-not $script:ApConfig) { Import-ApConfig | Out-Null }
    return $script:ApConfig
}

function Set-ApConfigValue {
    <#
    .SYNOPSIS
    Sets one key in the in-memory config. Call Save-ApConfig to persist.
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][AllowNull()][AllowEmptyString()][AllowEmptyCollection()]$Value
    )

    $config = Get-ApConfig
    if (-not $config.Contains($Name)) {
        throw "Unknown configuration key '$Name'."
    }
    $config[$Name] = $Value
}

function Add-ApGroupTagToHistory {
    <#
    .SYNOPSIS
    Records a group tag as most-recently-used, de-duplicated case-insensitively.
    #>
    param([string]$GroupTag)

    if ([string]::IsNullOrWhiteSpace($GroupTag)) { return }

    $tag = $GroupTag.Trim()
    $config = Get-ApConfig
    $existing = @($config.groupTagHistory | Where-Object { $_ -and $_ -ne '' })

    $kept = @($existing | Where-Object { $_ -ne $tag })
    $config.groupTagHistory = @(@($tag) + $kept | Select-Object -First 25)
}

function Save-ApConfig {
    <#
    .SYNOPSIS
    Writes the in-memory config to the first writable candidate path.
    Returns the path written, or $null on failure.
    #>
    [CmdletBinding()]
    param()

    $config = Get-ApConfig
    $json = $config | ConvertTo-Json -Depth 6

    $targets = @()
    if ($script:ApConfigPath) { $targets += $script:ApConfigPath }
    $targets += (Get-ApConfigCandidatePaths)

    foreach ($path in ($targets | Select-Object -Unique)) {
        try {
            $dir = Split-Path -Parent $path
            if ($dir -and -not (Test-Path -LiteralPath $dir)) {
                New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null
            }
            Set-Content -LiteralPath $path -Value $json -Encoding UTF8 -ErrorAction Stop
            $script:ApConfigPath = $path
            return $path
        }
        catch {
            continue
        }
    }

    Write-ApLog 'Could not save configuration to any candidate location.' -Level WARN
    return $null
}

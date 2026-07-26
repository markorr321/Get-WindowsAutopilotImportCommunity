# VendorScript.ps1 -- locate the community engine script that the GUI drives.
#
# Resolution order (mirrors VM-Pilot's AutopilotV2Import.ps1:44-57, which pre-injects
# the script so imports still work when PSGallery is unreachable -- the normal case
# on a locked-down OOBE network):
#
#   1. embedded base64 payload  (single-file dist build)
#   2. vendor\ next to the source tree  (dev / folder deployment)
#   3. an existing PSGallery install    (Get-InstalledScript)
#   4. Install-Script from PSGallery    (last resort, needs internet)
#
# The vendored copy is byte-exact so Andrew Taylor's Authenticode signature survives;
# Test-ApVendorScript verifies both the SHA256 from VERSION.json and the signature.

$script:ApVendorCache = @{}

# Populated by build.ps1 in the single-file dist. Empty in the source tree.
if (-not (Test-Path Variable:script:ApEmbeddedScripts)) {
    $script:ApEmbeddedScripts = @{}
}

function Get-ApVendorManifest {
    <#
    .SYNOPSIS
    Reads vendor\VERSION.json, or the embedded copy in the dist build.
    #>
    if ($script:ApVendorManifestCache) { return $script:ApVendorManifestCache }

    $manifest = $null

    if ($script:ApEmbeddedVersionJson) {
        try { $manifest = ConvertFrom-Json $script:ApEmbeddedVersionJson -ErrorAction Stop } catch { }
    }

    if (-not $manifest -and $script:ApAppRoot) {
        $path = Join-Path $script:ApAppRoot 'vendor\VERSION.json'
        if (Test-Path -LiteralPath $path) {
            try { $manifest = ConvertFrom-Json (Get-Content -LiteralPath $path -Raw) -ErrorAction Stop } catch { }
        }
    }

    $script:ApVendorManifestCache = $manifest
    return $manifest
}

function Get-ApVendorExpectedHash {
    param([Parameter(Mandatory)][string]$FileName)

    $manifest = Get-ApVendorManifest
    if (-not $manifest) { return $null }

    $entry = $manifest.scripts | Where-Object { $_.file -eq $FileName } | Select-Object -First 1
    if ($entry) { return $entry.sha256 }
    return $null
}

function Test-ApVendorScript {
    <#
    .SYNOPSIS
    Verifies a resolved engine script against VERSION.json and its Authenticode signature.

    .DESCRIPTION
    Returns a result object rather than throwing. A checksum mismatch is fatal (the file
    is not what we shipped); an invalid or absent signature is only a warning, because a
    PSGallery-installed copy may legitimately be a newer, differently-signed version.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$SkipHashCheck
    )

    $result = [ordered]@{
        Path            = $Path
        Exists          = (Test-Path -LiteralPath $Path -PathType Leaf)
        HashMatches     = $null
        ActualHash      = $null
        ExpectedHash    = $null
        SignatureStatus = 'NotChecked'
        Signer          = $null
        IsTrusted       = $false
        Messages        = @()
    }

    if (-not $result.Exists) {
        $result.Messages += "Engine script not found at $Path"
        return [pscustomobject]$result
    }

    $fileName = Split-Path -Leaf $Path
    $expected = Get-ApVendorExpectedHash -FileName $fileName
    $result.ExpectedHash = $expected

    if (-not $SkipHashCheck -and $expected) {
        try {
            $actual = (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash
            $result.ActualHash = $actual
            $result.HashMatches = ($actual -eq $expected)
            if (-not $result.HashMatches) {
                $result.Messages += "Checksum mismatch for $fileName (expected $expected, got $actual)"
            }
        }
        catch {
            $result.Messages += "Could not hash ${fileName}: $($_.Exception.Message)"
        }
    }

    try {
        $sig = Get-AuthenticodeSignature -LiteralPath $Path -ErrorAction Stop
        $result.SignatureStatus = "$($sig.Status)"
        if ($sig.SignerCertificate) { $result.Signer = $sig.SignerCertificate.Subject }
        if ($sig.Status -ne 'Valid') {
            $result.Messages += "Authenticode signature is '$($sig.Status)' for $fileName"
        }
    }
    catch {
        $result.Messages += "Could not read signature for ${fileName}: $($_.Exception.Message)"
    }

    # Trusted enough to run: correct bytes, or no manifest entry to compare against.
    $result.IsTrusted = ($result.HashMatches -ne $false)

    return [pscustomobject]$result
}

function Expand-ApEmbeddedScript {
    <#
    .SYNOPSIS
    Writes an embedded base64 payload to disk byte-for-byte.

    .DESCRIPTION
    Uses WriteAllBytes, never Set-Content, so no encoding or line-ending translation
    happens -- that would break the Authenticode signature.
    #>
    param(
        [Parameter(Mandatory)][string]$FileName,
        [Parameter(Mandatory)][string]$Destination
    )

    if (-not $script:ApEmbeddedScripts.ContainsKey($FileName)) { return $null }

    $dir = Split-Path -Parent $Destination
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    $bytes = [Convert]::FromBase64String($script:ApEmbeddedScripts[$FileName])
    [System.IO.File]::WriteAllBytes($Destination, $bytes)
    return $Destination
}

function Get-ApWorkingDirectory {
    <#
    .SYNOPSIS
    Per-session scratch directory for extracted scripts, launchers and run logs.
    #>
    if ($script:ApWorkingDirectory -and (Test-Path -LiteralPath $script:ApWorkingDirectory)) {
        return $script:ApWorkingDirectory
    }

    $base = Join-Path $env:TEMP ('AutopilotImportGUI\{0}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    New-Item -ItemType Directory -Path $base -Force | Out-Null
    $script:ApWorkingDirectory = $base
    return $base
}

function Resolve-ApEngineScript {
    <#
    .SYNOPSIS
    Returns the full path to a community script, acquiring it if necessary.

    .PARAMETER Name
    'Engine' for get-windowsautopilotinfocommunity.ps1, 'Diagnostics' for
    Get-AutopilotDiagnosticsCommunity.ps1.

    .PARAMETER AllowInstall
    Permit the PSGallery fallback (needs internet). On by default.
    #>
    [CmdletBinding()]
    param(
        [ValidateSet('Engine', 'Diagnostics')]
        [string]$Name = 'Engine',
        [switch]$AllowInstall = $true
    )

    if ($script:ApVendorCache.ContainsKey($Name)) { return $script:ApVendorCache[$Name] }

    $fileName = if ($Name -eq 'Engine') {
        'get-windowsautopilotinfocommunity.ps1'
    } else {
        'Get-AutopilotDiagnosticsCommunity.ps1'
    }
    $galleryName = if ($Name -eq 'Engine') {
        'Get-WindowsAutopilotInfoCommunity'
    } else {
        'Get-AutopilotDiagnosticsCommunity'
    }

    # 1. embedded payload (single-file build)
    if ($script:ApEmbeddedScripts.ContainsKey($fileName)) {
        $dest = Join-Path (Get-ApWorkingDirectory) $fileName
        if (-not (Test-Path -LiteralPath $dest)) {
            Expand-ApEmbeddedScript -FileName $fileName -Destination $dest | Out-Null
            Write-ApLog "Extracted embedded $fileName to $dest"
        }
        $script:ApVendorCache[$Name] = $dest
        return $dest
    }

    # 2. vendor folder in the source tree
    if ($script:ApAppRoot) {
        $vendored = Join-Path $script:ApAppRoot ('vendor\{0}' -f $fileName)
        if (Test-Path -LiteralPath $vendored -PathType Leaf) {
            Write-ApLog "Using vendored $fileName"
            $script:ApVendorCache[$Name] = $vendored
            return $vendored
        }
    }

    # 3. already installed from PSGallery
    try {
        $installed = Get-InstalledScript -Name $galleryName -ErrorAction Stop
        if ($installed) {
            $candidate = Join-Path $installed.InstalledLocation "$galleryName.ps1"
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                Write-ApLog "Using PSGallery-installed $galleryName $($installed.Version)"
                $script:ApVendorCache[$Name] = $candidate
                return $candidate
            }
        }
    }
    catch {
        # Not installed; fall through.
    }

    # 4. install from PSGallery
    if ($AllowInstall) {
        try {
            Write-ApLog "Installing $galleryName from the PowerShell Gallery..."
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
            Install-Script -Name $galleryName -Force -Scope AllUsers -Confirm:$false -ErrorAction Stop

            $installed = Get-InstalledScript -Name $galleryName -ErrorAction Stop
            $candidate = Join-Path $installed.InstalledLocation "$galleryName.ps1"
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                $script:ApVendorCache[$Name] = $candidate
                return $candidate
            }
        }
        catch {
            Write-ApLog "Could not install ${galleryName}: $($_.Exception.Message)" -Level ERROR
        }
    }

    throw "Unable to locate $fileName. Place it in the vendor folder next to this script, or connect to the internet so it can be installed from the PowerShell Gallery."
}

# DeviceInfo.ps1 -- local hardware/OS facts for the Device page and readiness checks.
#
# Uses CIM throughout. The original GUI mixed Get-WmiObject (deprecated, removed in
# PowerShell 7) with Get-CimInstance -- see Get-WindowsAutopilotImportGUI.ps1:564.
#
# Every getter is individually guarded: on a machine where one WMI provider is
# broken, the Device page should still render everything else rather than blank out.
#
# Nothing in this file may call Write-ApLog or read configuration. The whole set is copied
# into a bare background runspace by Start-ApDeviceLoad, where those helpers do not exist.
# It has to run off the UI thread because two of these queries are slow: on a machine where
# the caller lacks rights, root/cimv2/security/microsofttpm and root/cimv2/mdm/dmmap are
# inaccessible and each call sits on a ~5 second DCOM timeout before failing. Doing that
# synchronously delayed the window by over ten seconds.

function Get-ApCimSafe {
    param(
        [Parameter(Mandatory)][string]$ClassName,
        [string]$Namespace,
        [string]$Filter
    )

    try {
        $params = @{ ClassName = $ClassName; ErrorAction = 'Stop' }
        if ($Namespace) { $params.Namespace = $Namespace }
        if ($Filter) { $params.Filter = $Filter }
        return Get-CimInstance @params
    }
    catch {
        return $null
    }
}

function Get-ApDeviceIdentifierPart {
    <#
    .SYNOPSIS
    Normalises a manufacturer/model string for the v2 device identifier.

    .DESCRIPTION
    Mirrors the community script exactly (get-windowsautopilotinfocommunity.ps1:2115-2116):
    trim, then strip '.' and ',' -- the comma because it is the field separator in the
    "Manufacturer,Model,Serial" triple, the period because Intune's matching rejects it.
    Keeping this identical matters: a mismatch here means the identifier the GUI previews
    is not the one that gets imported.
    #>
    param([AllowNull()][string]$Value)

    if ($null -eq $Value) { return '' }
    return $Value.Trim().Replace('.', '').Replace(',', '')
}

function Test-ApHardwareHashAvailable {
    <#
    .SYNOPSIS
    True when the Autopilot 4K hardware hash can be read (Autopilot v1 prerequisite).
    #>
    [OutputType([bool])]
    param()

    $devDetail = Get-ApCimSafe -ClassName 'MDM_DevDetail_Ext01' `
                               -Namespace 'root/cimv2/mdm/dmmap' `
                               -Filter "InstanceID='Ext' AND ParentID='./DevDetail'"

    return [bool]($devDetail -and $devDetail.DeviceHardwareData)
}

function Get-ApHardwareHash {
    <#
    .SYNOPSIS
    Returns the base64 4K hardware hash, or $null when unavailable.
    #>
    [OutputType([string])]
    param()

    $devDetail = Get-ApCimSafe -ClassName 'MDM_DevDetail_Ext01' `
                               -Namespace 'root/cimv2/mdm/dmmap' `
                               -Filter "InstanceID='Ext' AND ParentID='./DevDetail'"

    if ($devDetail -and $devDetail.DeviceHardwareData) { return $devDetail.DeviceHardwareData }
    return $null
}

function Get-ApTpmInfo {
    <#
    .SYNOPSIS
    TPM presence, enablement and spec version.
    #>
    $tpm = Get-ApCimSafe -ClassName 'Win32_Tpm' -Namespace 'root/cimv2/security/microsofttpm'

    if (-not $tpm) {
        return [ordered]@{ Present = $false; Enabled = $false; Ready = $false; SpecVersion = 'Not detected' }
    }

    # SpecVersion looks like "2.0, 0, 1.38"; the major version is all we display.
    $spec = "$($tpm.SpecVersion)"
    $major = if ($spec -match '^\s*(\d+\.\d+)') { $Matches[1] } else { $spec }

    return [ordered]@{
        # A Win32_Tpm instance only exists when a TPM is physically present.
        Present     = $true
        Enabled     = [bool]$tpm.IsEnabled_InitialValue
        Ready       = ([bool]$tpm.IsEnabled_InitialValue -and [bool]$tpm.IsActivated_InitialValue)
        SpecVersion = $major
    }
}

function Test-ApSecureBoot {
    <#
    .SYNOPSIS
    $true/$false when determinable, $null on legacy BIOS or when the cmdlet is absent.
    #>
    try {
        if (-not (Get-Command Confirm-SecureBootUEFI -ErrorAction SilentlyContinue)) { return $null }
        return [bool](Confirm-SecureBootUEFI -ErrorAction Stop)
    }
    catch {
        # Throws "Cmdlet not supported on this platform" on legacy/BIOS machines.
        return $null
    }
}

function Test-ApInternetConnection {
    <#
    .SYNOPSIS
    Fast reachability probe for the header indicator.

    .DESCRIPTION
    A TCP connect to login.microsoftonline.com:443 -- what actually matters for
    Autopilot -- rather than the original's ICMP ping to 8.8.8.8
    (Get-WindowsAutopilotImportGUI.ps1:630), which fails on any network that blocks
    ICMP or DNS to public resolvers while Microsoft endpoints work fine.
    #>
    [OutputType([bool])]
    param(
        [string]$ComputerName = 'login.microsoftonline.com',
        [int]$Port = 443,
        [int]$TimeoutMs = 3000
    )

    $client = $null
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $async = $client.BeginConnect($ComputerName, $Port, $null, $null)
        if (-not $async.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) { return $false }
        $client.EndConnect($async)
        return $true
    }
    catch {
        return $false
    }
    finally {
        if ($client) { $client.Close() }
    }
}

function Get-ApDeviceInfo {
    <#
    .SYNOPSIS
    Snapshot of everything the GUI shows about the local machine.
    #>
    [CmdletBinding()]
    param()

    $cs   = Get-ApCimSafe -ClassName 'Win32_ComputerSystem'
    $bios = Get-ApCimSafe -ClassName 'Win32_BIOS'
    $os   = Get-ApCimSafe -ClassName 'Win32_OperatingSystem'
    $tpm  = Get-ApTpmInfo

    $serial       = if ($bios) { "$($bios.SerialNumber)".Trim() } else { '' }
    $manufacturer = if ($cs) { "$($cs.Manufacturer)".Trim() } else { '' }
    $model        = if ($cs) { "$($cs.Model)".Trim() } else { '' }

    # System drive free space is what matters for an Autopilot deployment; the
    # original summed every logical disk, which overstated it on multi-disk machines.
    $sysDrive = $env:SystemDrive
    if (-not $sysDrive) { $sysDrive = 'C:' }
    $disk = Get-ApCimSafe -ClassName 'Win32_LogicalDisk' -Filter "DeviceID='$sysDrive'"

    $freeGb  = if ($disk -and $disk.FreeSpace) { [math]::Round($disk.FreeSpace / 1GB, 1) } else { $null }
    $totalGb = if ($disk -and $disk.Size) { [math]::Round($disk.Size / 1GB, 1) } else { $null }
    $ramGb   = if ($cs -and $cs.TotalPhysicalMemory) { [math]::Round($cs.TotalPhysicalMemory / 1GB, 1) } else { $null }

    $hashAvailable = Test-ApHardwareHashAvailable

    $idManufacturer = Get-ApDeviceIdentifierPart $manufacturer
    $idModel        = Get-ApDeviceIdentifierPart $model

    return [ordered]@{
        SerialNumber      = $serial
        Manufacturer      = $manufacturer
        Model             = $model
        ComputerName      = if ($cs) { "$($cs.Name)".Trim() } else { $env:COMPUTERNAME }
        OsCaption         = if ($os) { "$($os.Caption)".Trim() } else { '' }
        OsBuild           = if ($os) { "$($os.Version)" } else { '' }
        OsDisplayVersion  = Get-ApWindowsDisplayVersion
        SystemDrive       = $sysDrive
        FreeSpaceGb       = $freeGb
        TotalSpaceGb      = $totalGb
        MemoryGb          = $ramGb
        TpmPresent        = $tpm.Present
        TpmEnabled        = $tpm.Enabled
        TpmSpecVersion    = $tpm.SpecVersion
        SecureBoot        = Test-ApSecureBoot
        HardwareHashReady = $hashAvailable
        IsElevated        = Test-ApElevated
        IsVirtualMachine  = Test-ApVirtualMachine $manufacturer $model

        # "Manufacturer,Model,Serial" -- exactly what -identifier imports for Autopilot v2.
        DeviceIdentifier  = '{0},{1},{2}' -f $idManufacturer, $idModel, $serial
    }
}

function Get-ApWindowsDisplayVersion {
    try {
        $key = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
        $props = Get-ItemProperty -Path $key -ErrorAction Stop
        if ($props.DisplayVersion) { return $props.DisplayVersion }
        if ($props.ReleaseId) { return $props.ReleaseId }
        return ''
    }
    catch {
        return ''
    }
}

function Test-ApVirtualMachine {
    param([string]$Manufacturer, [string]$Model)

    $needles = @('Virtual', 'VMware', 'Hyper-V', 'KVM', 'QEMU', 'Xen', 'VirtualBox', 'Parallels')
    $haystack = "$Manufacturer $Model"

    foreach ($n in $needles) {
        if ($haystack -like "*$n*") { return $true }
    }
    return $false
}

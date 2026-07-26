# Logging.ps1 -- session log file + in-memory ring buffer for the Logs page.
#
# The original AutoPilot_Import_GUI wrote to the root of C:\ (Get-WindowsAutopilotImportGUI.ps1:520),
# which fails on locked-down machines and litters the system drive. We use
# %ProgramData%\AutopilotImportGUI\Logs and fall back to %TEMP% if that is not writable
# (which happens in some OOBE contexts before ProgramData ACLs settle).

$script:ApLogPath = $null
$script:ApLogBuffer = New-Object System.Collections.Generic.List[string]

function Get-ApLogDirectory {
    <#
    .SYNOPSIS
    Returns a writable directory for logs, creating it if needed.
    #>
    [CmdletBinding()]
    param()

    $candidates = @(
        (Join-Path $env:ProgramData 'AutopilotImportGUI\Logs'),
        (Join-Path $env:TEMP 'AutopilotImportGUI\Logs')
    )

    foreach ($dir in $candidates) {
        if ([string]::IsNullOrWhiteSpace($dir)) { continue }
        try {
            if (-not (Test-Path -LiteralPath $dir)) {
                New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null
            }
            # Prove it is actually writable rather than trusting the ACL.
            $probe = Join-Path $dir ('.write-test-{0}' -f ([guid]::NewGuid().ToString('N')))
            [System.IO.File]::WriteAllText($probe, 'x')
            Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
            return $dir
        }
        catch {
            continue
        }
    }

    throw 'No writable location found for the log directory.'
}

function Initialize-ApLog {
    <#
    .SYNOPSIS
    Starts a new session log file. Safe to call more than once.
    #>
    [CmdletBinding()]
    param(
        [string]$Directory
    )

    if (-not $Directory) { $Directory = Get-ApLogDirectory }

    $name = 'AutopilotImportGUI-{0}.log' -f (Get-Date -Format 'yyyyMMdd-HHmmss')
    $script:ApLogPath = Join-Path $Directory $name

    $header = @(
        '=' * 78
        ' Autopilot Import GUI (Community)'
        ' Mark Orr (@markorr321) - https://orr365.tools'
        '-' * 78
        (' Session started : {0}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz'))
        (' Computer        : {0}' -f $env:COMPUTERNAME)
        (' User            : {0}\{1}' -f $env:USERDOMAIN, $env:USERNAME)
        (' PowerShell      : {0}' -f $PSVersionTable.PSVersion)
        (' OS              : {0}' -f (Get-ApOsCaption))
        '=' * 78
    ) -join [Environment]::NewLine

    try {
        Set-Content -LiteralPath $script:ApLogPath -Value $header -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        # Logging must never take the GUI down.
        $script:ApLogPath = $null
    }

    return $script:ApLogPath
}

function Get-ApOsCaption {
    try {
        (Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).Caption
    }
    catch {
        [Environment]::OSVersion.VersionString
    }
}

function Get-ApLogPath {
    if (-not $script:ApLogPath) { Initialize-ApLog | Out-Null }
    return $script:ApLogPath
}

function Write-ApLog {
    <#
    .SYNOPSIS
    Appends a timestamped line to the session log and the in-memory buffer.

    .PARAMETER Level
    INFO, WARN, ERROR or DEBUG. Rendered into the file; the GUI colours from it.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [AllowEmptyString()]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')]
        [string]$Level = 'INFO'
    )

    process {
        $line = '{0} [{1,-5}] {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message

        $script:ApLogBuffer.Add($line)
        # Keep the buffer bounded; the file on disk remains complete.
        if ($script:ApLogBuffer.Count -gt 5000) {
            $script:ApLogBuffer.RemoveRange(0, 1000)
        }

        $path = $script:ApLogPath
        if ($path) {
            try {
                Add-Content -LiteralPath $path -Value $line -Encoding UTF8 -ErrorAction Stop
            }
            catch {
                # A transient sharing violation must not break the app.
            }
        }
    }
}

function Get-ApLogBuffer {
    <#
    .SYNOPSIS
    Returns the in-memory log lines, newest last.
    #>
    return $script:ApLogBuffer.ToArray()
}

function Clear-ApLogBuffer {
    $script:ApLogBuffer.Clear()
}

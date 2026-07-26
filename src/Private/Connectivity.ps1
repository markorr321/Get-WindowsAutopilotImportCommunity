# Connectivity.ps1 -- Autopilot network prerequisite checks.
#
# A rewrite of connectivity_check() from the original GUI
# (Get-WindowsAutopilotImportGUI.ps1:245-508), which never actually worked: every result
# line called `Write-Output -NoNewline -ForegroundColor ...`, and Write-Output has no
# -NoNewline or -ForegroundColor parameter, so all ~29 checks threw a
# ParameterBindingException that was then swallowed by $ErrorActionPreference =
# 'SilentlyContinue'. The tech saw an empty window and concluded the network was fine.
#
# Differences here:
#   * raw TcpClient with an explicit timeout instead of Test-NetConnection -- roughly an
#     order of magnitude faster, reports latency, and needs no extra module
#   * runspace pool, so ~30 endpoints complete in about a second rather than serially
#   * results are returned as objects for a DataGrid rather than printed with colour codes
#   * endpoints are data, overridable via connectivityEndpoints in config.json
#   * dropped the original's bare "azure.net" probe, which is not a resolvable host, and
#     replaced it with the documented TPM attestation endpoint

function Get-ApDefaultEndpoints {
    <#
    .SYNOPSIS
    Documented Windows Autopilot / Intune network requirements.

    .DESCRIPTION
    Required = $true marks endpoints where a failure will actually break enrolment;
    the rest are deployment-time niceties reported for completeness.
    #>
    [CmdletBinding()]
    param()

    return @(
        # --- Entra ID / MDM enrolment -----------------------------------------
        [pscustomobject]@{ Category = 'Entra ID and enrolment'; Name = 'Entra device registration'; Host = 'enterpriseregistration.windows.net'; Port = 443; Required = $true }
        [pscustomobject]@{ Category = 'Entra ID and enrolment'; Name = 'Intune enrolment';          Host = 'enterpriseenrollment-s.manage.microsoft.com'; Port = 443; Required = $true }
        [pscustomobject]@{ Category = 'Entra ID and enrolment'; Name = 'Entra sign-in';             Host = 'login.microsoftonline.com'; Port = 443; Required = $true }
        [pscustomobject]@{ Category = 'Entra ID and enrolment'; Name = 'Microsoft Graph';           Host = 'graph.microsoft.com'; Port = 443; Required = $true }

        # --- Autopilot deployment service -------------------------------------
        [pscustomobject]@{ Category = 'Autopilot service'; Name = 'Autopilot ZTD';        Host = 'ztd.dds.microsoft.com'; Port = 443; Required = $true }
        [pscustomobject]@{ Category = 'Autopilot service'; Name = 'Autopilot CS';         Host = 'cs.dds.microsoft.com'; Port = 443; Required = $true }
        [pscustomobject]@{ Category = 'Autopilot service'; Name = 'Microsoft account';    Host = 'login.live.com'; Port = 443; Required = $true }
        [pscustomobject]@{ Category = 'Autopilot service'; Name = 'Intune service';       Host = 'manage.microsoft.com'; Port = 443; Required = $true }

        # --- TPM attestation (required for pre-provisioning / self-deploying) --
        [pscustomobject]@{ Category = 'TPM attestation'; Name = 'Intel TPM';             Host = 'ekop.intel.com'; Port = 443; Required = $false }
        [pscustomobject]@{ Category = 'TPM attestation'; Name = 'Qualcomm TPM';          Host = 'ekcert.spserv.microsoft.com'; Port = 443; Required = $false }
        [pscustomobject]@{ Category = 'TPM attestation'; Name = 'AMD TPM';               Host = 'ftpm.amd.com'; Port = 443; Required = $false }
        [pscustomobject]@{ Category = 'TPM attestation'; Name = 'Host attestation';      Host = 'has.spserv.microsoft.com'; Port = 443; Required = $false }

        # --- Licensing / activation -------------------------------------------
        [pscustomobject]@{ Category = 'Activation'; Name = 'Activation service';         Host = 'activation.sls.microsoft.com'; Port = 443; Required = $false }
        [pscustomobject]@{ Category = 'Activation'; Name = 'Validation service';         Host = 'validation.sls.microsoft.com'; Port = 443; Required = $false }

        # --- Windows Update ----------------------------------------------------
        [pscustomobject]@{ Category = 'Windows Update'; Name = 'Windows Update';         Host = 'update.microsoft.com'; Port = 443; Required = $false }
        [pscustomobject]@{ Category = 'Windows Update'; Name = 'Update download';        Host = 'download.windowsupdate.com'; Port = 443; Required = $false }
        # Microsoft documents these as *.delivery.mp.microsoft.com and
        # *.dsp.mp.microsoft.com. The bare wildcard bases do not resolve, so probe a real
        # host from each rather than reporting a permanent false failure.
        [pscustomobject]@{ Category = 'Windows Update'; Name = 'Content delivery';       Host = 'dl.delivery.mp.microsoft.com'; Port = 443; Required = $false }
        [pscustomobject]@{ Category = 'Windows Update'; Name = 'Delivery optimisation';  Host = 'tsfe.trafficshaping.dsp.mp.microsoft.com'; Port = 443; Required = $false }

        # --- Single sign-on ----------------------------------------------------
        [pscustomobject]@{ Category = 'Single sign-on'; Name = 'Seamless SSO';           Host = 'autologon.microsoftazuread-sso.com'; Port = 443; Required = $false }

        # --- Configuration and app delivery -----------------------------------
        [pscustomobject]@{ Category = 'Configuration and apps'; Name = 'Office config';  Host = 'config.office.com'; Port = 443; Required = $false }
        [pscustomobject]@{ Category = 'Configuration and apps'; Name = 'Entra Graph (legacy)'; Host = 'graph.windows.net'; Port = 443; Required = $false }
        [pscustomobject]@{ Category = 'Configuration and apps'; Name = 'Content CDN (primary)';  Host = 'euprodimedatapri.azureedge.net'; Port = 443; Required = $false }
        [pscustomobject]@{ Category = 'Configuration and apps'; Name = 'Content CDN (secondary)'; Host = 'euprodimedatasec.azureedge.net'; Port = 443; Required = $false }
        [pscustomobject]@{ Category = 'Configuration and apps'; Name = 'Content CDN (hotfix)';    Host = 'euprodimedatahotfix.azureedge.net'; Port = 443; Required = $false }

        # --- Diagnostics -------------------------------------------------------
        [pscustomobject]@{ Category = 'Diagnostics'; Name = 'Connectivity test';         Host = 'www.msftconnecttest.com'; Port = 443; Required = $false }
        [pscustomobject]@{ Category = 'Diagnostics'; Name = 'PowerShell Gallery';        Host = 'www.powershellgallery.com'; Port = 443; Required = $false }
    )
}

function Get-ApConfiguredEndpoints {
    <#
    .SYNOPSIS
    The endpoint list to test: the config.json override if present, else the built-in list.
    #>
    [CmdletBinding()]
    param()

    $config = Get-ApConfig
    $override = $config.connectivityEndpoints

    if (-not $override -or @($override).Count -eq 0) { return Get-ApDefaultEndpoints }

    $list = New-Object System.Collections.Generic.List[object]
    foreach ($e in @($override)) {
        $endpointHost = Get-ApRequestValue $e 'Host' ''
        if (-not $endpointHost) { continue }

        $list.Add([pscustomobject]@{
            Category = Get-ApRequestValue $e 'Category' 'Custom'
            Name     = Get-ApRequestValue $e 'Name' $endpointHost
            Host     = $endpointHost
            Port     = [int](Get-ApRequestValue $e 'Port' 443)
            Required = [bool](Get-ApRequestValue $e 'Required' $false)
        })
    }

    if ($list.Count -eq 0) {
        Write-ApLog 'connectivityEndpoints in config.json contained no usable entries; using the built-in list.' -Level WARN
        return Get-ApDefaultEndpoints
    }

    Write-ApLog "Using $($list.Count) custom connectivity endpoints from config.json"
    return $list.ToArray()
}

function Test-ApEndpoint {
    <#
    .SYNOPSIS
    Single TCP reachability probe with latency.

    .DESCRIPTION
    Deliberately standalone with no dependencies on the rest of the app, because this
    function's source is injected into a runspace pool by Invoke-ApConnectivityCheck.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [int]$Port = 443,
        [int]$TimeoutMs = 4000
    )

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $client = $null
    $succeeded = $false
    $errorText = ''

    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $async = $client.BeginConnect($ComputerName, $Port, $null, $null)

        if ($async.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) {
            # EndConnect surfaces the real failure (refused, reset, DNS) rather than
            # letting a half-open socket look like a success.
            $client.EndConnect($async)
            $succeeded = $client.Connected
            if (-not $succeeded) { $errorText = 'Connection not established' }
        }
        else {
            $errorText = "Timed out after ${TimeoutMs}ms"
        }
    }
    catch {
        $errorText = $_.Exception.GetBaseException().Message
    }
    finally {
        if ($client) { try { $client.Close() } catch { } }
        $sw.Stop()
    }

    return [pscustomobject]@{
        Host      = $ComputerName
        Port      = $Port
        Succeeded = $succeeded
        LatencyMs = [int]$sw.ElapsedMilliseconds
        Error     = $errorText
    }
}

function Invoke-ApConnectivityCheck {
    <#
    .SYNOPSIS
    Tests every configured endpoint concurrently.

    .PARAMETER OnResult
    Optional scriptblock invoked with each result object as it completes, so the UI can
    populate the grid progressively instead of waiting for the whole sweep.

    .PARAMETER ThrottleLimit
    Concurrent probes. 12 keeps a slow OOBE network from being swamped while still
    finishing the full list in about a second.
    #>
    [CmdletBinding()]
    param(
        [object[]]$Endpoints,
        [scriptblock]$OnResult,
        [int]$ThrottleLimit = 12,
        [int]$TimeoutMs = 4000
    )

    # Deliberately no logging or config access in here: this function is copied into a bare
    # background runspace by Start-ApBackgroundWork, where Write-ApLog and Get-ApConfig do
    # not exist. Callers resolve the endpoint list and do the logging.
    if (-not $Endpoints) { throw 'Invoke-ApConnectivityCheck requires -Endpoints.' }

    $pool = [runspacefactory]::CreateRunspacePool(1, $ThrottleLimit)
    $pool.ApartmentState = 'MTA'
    $pool.Open()

    # Ship the probe function into each runspace as source text.
    $probeSource = "function Test-ApEndpoint {`n$((Get-Command Test-ApEndpoint).Definition)`n}"

    $jobs = New-Object System.Collections.Generic.List[object]

    try {
        foreach ($endpoint in $Endpoints) {
            $ps = [powershell]::Create()
            $ps.RunspacePool = $pool

            [void]$ps.AddScript($probeSource)
            [void]$ps.AddScript('param($h,$p,$t) Test-ApEndpoint -ComputerName $h -Port $p -TimeoutMs $t')
            [void]$ps.AddArgument($endpoint.Host)
            [void]$ps.AddArgument($endpoint.Port)
            [void]$ps.AddArgument($TimeoutMs)

            $jobs.Add([pscustomobject]@{
                Shell    = $ps
                Handle   = $ps.BeginInvoke()
                Endpoint = $endpoint
            })
        }

        $results = New-Object System.Collections.Generic.List[object]

        foreach ($job in $jobs) {
            $probe = $null
            try {
                $probe = @($job.Shell.EndInvoke($job.Handle))[-1]
            }
            catch {
                $probe = $null
            }
            finally {
                $job.Shell.Dispose()
            }

            $succeeded = [bool]($probe -and $probe.Succeeded)

            $result = [pscustomobject]@{
                Category  = $job.Endpoint.Category
                Name      = $job.Endpoint.Name
                Host      = $job.Endpoint.Host
                Port      = $job.Endpoint.Port
                Required  = $job.Endpoint.Required
                Succeeded = $succeeded
                LatencyMs = if ($probe) { $probe.LatencyMs } else { 0 }
                Error     = if ($probe) { $probe.Error } else { 'Probe did not run' }
                Status    = if ($succeeded) { 'OK' } elseif ($job.Endpoint.Required) { 'Failed' } else { 'Unreachable' }
            }

            $results.Add($result)
            if ($OnResult) { & $OnResult $result }
        }

        return $results.ToArray()
    }
    finally {
        $pool.Close()
        $pool.Dispose()
    }
}

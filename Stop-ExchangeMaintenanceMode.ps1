<#
.SYNOPSIS
    Resumes the given Exchange server from maintenance mode.
.DESCRIPTION
    This script uses the procedure described by Paul Cunningham on 
    https://practical365.com/exchange-server/installing-cumulative-updates-on-exchange-server-2016/
    to resume services of a Exchange server after maintenance (e.g. installing updates).
    The server provided SHOULD be part of a DAG, so that automatic distribution of databases can be done.
.LINK
    http://bechtle.com
.NOTES
    Written by: Marc Debold (marc.debold@bechtle.com)

    Change Log 
    V1.00, 27.04.2017 - Initial version 
.OUTPUTS
    Shows status information about the progress of ending the maintenance mode.
.PARAMETER Identity
    Used for the name of the Exchange server, that should be resumed from maintenance mode.
.PARAMETER DatabaseCopyAutoActivationPolicy
    When entering maintenance mode, the ActivationPolicy is set to BLOCKED. To resume after maintenance, the ActivationPolicy
    needs to be restored to the same value, as before. If the corresponding script was used, the ActivationPolicy was printed
    out while starting the maintenance mode. This parameter can be used to restore this setting. If omitted, UNRESTRICTED is
    used by default
.EXAMPLE
    Stop-ExchangeMaintenanceMode -Identity ExchSrv01
    Resumes operation for Exchange Server ExchSrv01.
.EXAMPLE
    Stop-ExchangeMaintenanceMode -Identity ExchSrv01 -DatabaseCopyAutoActivationPolicy IntraSiteOnly
    Resumes operation for Exchange Server ExchSrv01 setting the ActivationPolicy to INTRASITEONLY.
#>

[CmdletBinding()] param(
    [Parameter(Mandatory=$true,Position=0)]$Identity,
    [Parameter(Mandatory=$false,Position=1)][ValidateSet("Blocked", "IntrasiteOnly", "Unrestricted")]$DatabaseCopyAutoActivationPolicy = "Unrestricted"
)

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

$isError = $false
try {
    Write-Host "Checking server $($Identity) ... " -NoNewline
    $SourceServer = Get-ExchangeServer -Identity $Identity -ErrorAction Stop
    Write-Host "okay." -ForegroundColor Green
} catch {
    Write-Host "failed. Please enter a correct Exchange server name." -ForegroundColor Red
    $isError = $true
}

if (-not $isError) {
    # DAG Name herausfinden
    Write-Host "Searching for DAG membership ... " -NoNewline
    $DagName = (Get-DatabaseAvailabilityGroup | ? {$_.Servers -like $SourceServer.Name}).Name
    if ($DagName -ne $null) {
        Write-Host "found '$($DagName)'." -ForegroundColor Green
    } else {
        Write-Host "none found." -ForegroundColor Red
    }
    # Wartungsmodus aufheben
    Write-Host "Disabling maintenacne mode ... " -NoNewline
    Set-ServerComponentState $SourceServer.Name –Component ServerWideOffline –State Active –Requester Maintenance
    Write-Host "done." -ForegroundColor Green

    # Im DAG wieder aktivieren
    Write-Host "Resuming cluster node ... " -NoNewline
    $ClusterState = (Get-ClusterNode -Name $SourceServer.Name).State
    if ($ClusterState -eq "Paused") {
        try {
            Resume-ClusterNode –Name $SourceServer.Name | Out-Null
            Write-Host "done." -ForegroundColor Green
        } catch {
            Write-Host "error resuming cluster node." -ForegroundColor Red
        }
    } else {
        Write-Host "cluster in state '$($ClusterState)'. Cannot resume." -ForegroundColor Red
    }

    # Autoaktivierung wieder reaktivieren
    Write-Host "Setting activation policy to '$($DatabaseCopyAutoActivationPolicy)' ... " -NoNewline
    Set-MailboxServer $SourceServer.Name –DatabaseCopyAutoActivationPolicy $DatabaseCopyAutoActivationPolicy
    Write-Host "done." -ForegroundColor Green

    # Autoaktivierung wieder einschalten
    Write-Host "Enabling database copy actovation ... " -NoNewline
    Set-MailboxServer $SourceServer.Name –DatabaseCopyActivationDisabledAndMoveNow $false
    Write-Host "done." -ForegroundColor Green

    # Hub Transport wieder aktivieren
    Write-Host "Resuming hub transport role ... " -NoNewline
    Set-ServerComponentState $SourceServer.Name –Component HubTransport –State Active –Requester Maintenance
    Write-Host "done." -ForegroundColor Green

    if ($DagName -ne $null) {
        # Datenbanken wieder neu verteilen
        Write-Host "Rebalancing databases for DAG '$($DagName)' ... " -NoNewline
        & "$($env:ExchangeInstallPath)\Scripts\RedistributeActiveDatabases.ps1" -DagName $DagName -BalanceDbsByActivationPreference -Confirm:$false
        Write-Host "done." -ForegroundColor Green
    }
    Write-Host "`n#####################################`n# Server is now out of maintenance. #`n#####################################`n"
}
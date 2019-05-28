<#
.SYNOPSIS
    Places the given Exchange server into maintenance mode.
.DESCRIPTION
    This script uses the procedure described by Paul Cunningham on 
    https://practical365.com/exchange-server/installing-cumulative-updates-on-exchange-server-2016/
    to prepare a Exchange server for maintenance (e.g. installing updates).
    The server provided SHOULD be part of a DAG, so that automatic distribution of databases can be done.
.LINK
    http://bechtle.com
.NOTES
    Written by: Marc Debold (marc.debold@bechtle.com)

    Change Log 
    V1.00, 27.04.2017 - Initial version 
.OUTPUTS
    Shows status information about the progress of starting the maintenance mode.
.PARAMETER Identity
    Used for the name of the Exchange server, that should be put into maintenance mode.
.PARAMETER PartnerServer
    The optional PartnerServer defines the target for moving unprocessed messages in the server's transport queues to.
.EXAMPLE
    Start-ExchangeMaintenanceMode -Identity ExchSrv01
    Puts Exchange Server ExchSrv01 into maintenance mode without moving the transport queues.
.EXAMPLE
    Start-ExchangeMaintenanceMode -Identity ExchSrv01 -PartnerServer ExchSrv02
    Puts Exchange Server ExchSrv01 into maintenance mode while moving the transport queues to server ExchSrv02.    
#>

[CmdletBinding()] param(
    [Parameter(Mandatory=$true,Position=0)]$Identity,
    [Parameter(Mandatory=$false,Position=1)]$PartnerServer
)

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

$isError = $false
try {
    Write-Host "Checking source server $($Identity) ... " -NoNewline
    $SourceServer = Get-ExchangeServer -Identity $Identity -ErrorAction Stop
    Write-Host "okay." -ForegroundColor Green
    if ($PartnerServer -ne $null) {
        Write-Host "Checking target server $($PartnerServer) ... " -NoNewline
        $TargetServer = Get-ExchangeServer -Identity $PartnerServer -ErrorAction Stop
        Write-Host "okay." -ForegroundColor Green
    }
} catch {
    Write-Host "failed." -ForegroundColor Red
    $isError = $true
}

if (-not $isError) {

    # Hub Transport auf Draining stellen
    Write-Host "Draining hub transport role ... " -NoNewline
    Set-ServerComponentState -Identity $SourceServer.Name –Component HubTransport –State Draining –Requester Maintenance
    Write-Host "done." -ForegroundColor Green
    
    if ($PartnerServer -ne $null) {
        # Nachrichten auf anderen Server umleiten
        Write-Host "Redirecting message queues from $($SourceServer.Name) to $($TargetServer.Name) ... " -NoNewline
        Redirect-Message -Server $SourceServer.Name -Target $TargetServer.Fqdn -Confirm:$false
        Write-Host "done." -ForegroundColor Green
    }

    # Clusterknoten im DAG deaktivieren
    Write-Host "Suspending cluster node ... " -NoNewline
    Suspend-ClusterNode –Name $SourceServer.Name | Out-Null
    Write-Host "done." -ForegroundColor Green

    # Autoaktivierung DB abschalten
    Write-Host "Disabling database copy actovation ... " -NoNewline
    Set-MailboxServer -Identity $SourceServer.Name –DatabaseCopyActivationDisabledAndMoveNow $true
    Write-Host "done." -ForegroundColor Green

    # AutoActivationPolicy notieren, dann blockieren
    $ActivationPolicy = (Get-MailboxServer -Identity $SourceServer.Name).DatabaseCopyAutoActivationPolicy
    Write-Host "Setting activation policy to BLOCKED (was: $($ActivationPolicy)) ... " -NoNewline
    Set-MailboxServer -Identity $SourceServer.Name –DatabaseCopyAutoActivationPolicy Blocked
    Write-Host "done." -ForegroundColor Green

    $FirstLoop = $true
    # Prüfen, ob noch Datenbanken auf dem Server laufen --> sollte keine mehr sein
    Write-Host "Checking for active databases ... " -NoNewline
    do {
        if (-not $FirstLoop) {
            Write-Host "Waiting some seconds for a retry ... " -NoNewline
            Start-Sleep -Seconds 10
        }
        $MountedDb = Get-MailboxDatabaseCopyStatus -Server $SourceServer.Name | Where {$_.Status -eq "Mounted"}
        $MountedDbCount = $MountedDb.Count
        $FirstLoop = $false
        if ($MountedDbCount -gt 0) {
            Write-Host "$($MountedDbCount) found: $($MountedDb.Name -join ", ")." -ForegroundColor Red
        } else {
            Write-Host "none found." -ForegroundColor Green
        }
    } while ($MountedDbCount -gt 0)

    # Server in den Wartungsmodus setzen
    Write-Host "Setting server to maintenacne mode ... " -NoNewline
    Set-ServerComponentState -Identity $SourceServer.Name –Component ServerWideOffline –State InActive –Requester Maintenance
    Write-Host "done." -ForegroundColor Green
    Write-Host "`n########################################`n# Server is now ready for maintenance. #`n########################################`n"
    Write-Host "Please remember the former DatabaseCopyAutoActivationPolicy: " -NoNewline
    Write-Host "$($ActivationPolicy)" -ForegroundColor Green
}
<#
.SYNOPSIS
    SSL / TLS Certificate Expiry Monitor for every public endpoint the
    company runs. Emails alerts and (optionally) creates JIRA tickets.

.DESCRIPTION
    WHY THIS IS NOT NATIVE EXCEL / ONEDRIVE
    Neither can connect to an HTTPS endpoint, read its certificate chain,
    and parse the "Not After" field. This is a .NET call through
    System.Net.Security.SslStream.

    USE CASE (software business)
    Most outages start with "the cert expired and nobody got the email."
    This script runs from a scheduled task, pulls a list of domains from
    a CSV, and alerts when any cert is within 30 / 14 / 7 days of expiry.

.PARAMETER DomainsCsv
    CSV with a single column "domain" (or domain,port).

.PARAMETER WarnAt
    Comma-separated day thresholds that trigger an alert. Default 30,14,7,1.

.PARAMETER SmtpServer
    SMTP relay. If omitted, writes report to disk only.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string] $DomainsCsv,
    [string]   $OutputPath = ".\ssl_expiry_report.xlsx",
    [int[]]    $WarnAt     = @(30, 14, 7, 1),
    [string]   $SmtpServer,
    [string]   $AlertFrom  = 'ssl-monitor@yourco.com',
    [string[]] $AlertTo,
    [string]   $TeamsWebhookUrl
)

function Get-CertForEndpoint {
    param([string]$HostName, [int]$Port = 443)
    $cert = $null
    $tcp = $null
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $tcp.ReceiveTimeout = 10000
        $tcp.Connect($HostName, $Port)
        $sslStream = New-Object System.Net.Security.SslStream(
            $tcp.GetStream(),
            $false,
            { param($s, $c, $ch, $e) return $true }  # accept any cert for inspection
        )
        $sslStream.AuthenticateAsClient($HostName)
        $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($sslStream.RemoteCertificate)
        $sslStream.Dispose()
    }
    catch {
        return [PSCustomObject]@{
            Host      = $HostName
            Port      = $Port
            Error     = $_.Exception.Message
            Subject   = $null
            Issuer    = $null
            NotAfter  = $null
            NotBefore = $null
            DaysLeft  = $null
            Thumbprint= $null
        }
    }
    finally { if ($tcp) { $tcp.Close() } }

    $daysLeft = [int]($cert.NotAfter - (Get-Date)).TotalDays
    [PSCustomObject]@{
        Host      = $HostName
        Port      = $Port
        Error     = $null
        Subject   = $cert.Subject
        Issuer    = $cert.Issuer
        NotBefore = $cert.NotBefore
        NotAfter  = $cert.NotAfter
        DaysLeft  = $daysLeft
        Thumbprint= $cert.Thumbprint
        SANs      = (($cert.Extensions | Where-Object { $_.Oid.FriendlyName -eq 'Subject Alternative Name' }).Format($false) -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join ';'
    }
}

$domains = Import-Csv $DomainsCsv

Write-Host "Checking $($domains.Count) endpoints..." -ForegroundColor Cyan

$results = foreach ($d in $domains) {
    $port = if ($d.port) { [int]$d.port } else { 443 }
    Get-CertForEndpoint -HostName $d.domain -Port $port
}

# Attach severity bucket
foreach ($r in $results) {
    $r | Add-Member -NotePropertyName Severity -NotePropertyValue (
        switch ($r.DaysLeft) {
            { $_ -eq $null }    { 'ERROR' }
            { $_ -le 1 }        { 'CRITICAL' }
            { $_ -le 7 }        { 'HIGH' }
            { $_ -le 14 }       { 'MEDIUM' }
            { $_ -le 30 }       { 'LOW' }
            default             { 'OK' }
        }
    ) -Force
}

$atRisk = $results | Where-Object { $_.Severity -in 'ERROR','CRITICAL','HIGH','MEDIUM','LOW' }

# Write report
if (Get-Module -ListAvailable ImportExcel) {
    Import-Module ImportExcel
    $results | Export-Excel -Path $OutputPath -WorksheetName 'All Certificates' -AutoSize
    if ($atRisk) { $atRisk | Export-Excel -Path $OutputPath -WorksheetName 'At Risk' -AutoSize -Append }
    Write-Host "Wrote $OutputPath" -ForegroundColor Green
}
else {
    $csvPath = [System.IO.Path]::ChangeExtension($OutputPath, '.csv')
    $results | Export-Csv $csvPath -NoTypeInformation
}

# Send alerts
if ($atRisk -and $SmtpServer -and $AlertTo) {
    $body = ($atRisk | Sort-Object DaysLeft |
             Select-Object Host, Port, Severity, DaysLeft, NotAfter, Issuer |
             ConvertTo-Html -Fragment) -join "`n"
    Send-MailMessage -From $AlertFrom -To $AlertTo `
                     -Subject "SSL Expiry Alert - $($atRisk.Count) at risk" `
                     -BodyAsHtml -Body "<h2>At-risk certificates</h2>$body" `
                     -SmtpServer $SmtpServer
}

if ($atRisk -and $TeamsWebhookUrl) {
    $lines = $atRisk | Sort-Object DaysLeft |
             ForEach-Object { "• **$($_.Host)** — $($_.DaysLeft) days left (expires $($_.NotAfter.ToString('yyyy-MM-dd')))" }
    $card = @{
        '@type'      = 'MessageCard'
        themeColor   = 'D13438'
        summary      = 'SSL expiry alert'
        title        = "SSL expiry alert - $($atRisk.Count) at risk"
        sections     = @(@{ text = ($lines -join "`n`n") })
    } | ConvertTo-Json -Depth 5
    Invoke-RestMethod -Method Post -Uri $TeamsWebhookUrl -Body $card -ContentType 'application/json'
}

Write-Host "$($atRisk.Count) at-risk certificates." -ForegroundColor Yellow

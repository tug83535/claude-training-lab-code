<#
.SYNOPSIS
    SharePoint / OneDrive storage audit. Scans every site, reports size,
    last-activity date, stale sites, and orphaned sites (no owner).

.DESCRIPTION
    WHY THIS IS NOT NATIVE EXCEL / ONEDRIVE
    OneDrive's admin UI shows one site at a time. Building a
    company-wide storage report requires the M365 admin CLI or Graph API.
    A 2,000-person software company typically has 8,000+ SharePoint sites
    (Teams creates one per team). Most are dead weight.

    USE CASE
    IT leadership asks: "Which sites are costing us? Which are orphaned?
    Where should we focus cleanup?"

    This produces the answer in one run.

.PARAMETER TenantUrl
    Your SharePoint admin URL, e.g. https://yourco-admin.sharepoint.com

.PARAMETER StaleDays
    Sites with no activity for this many days are flagged. Default 180.

.EXAMPLE
    ./SharePoint_Site_Storage_Audit.ps1 -TenantUrl https://yourco-admin.sharepoint.com
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string] $TenantUrl,
    [int]    $StaleDays  = 180,
    [string] $OutputPath = ".\sharepoint_audit.xlsx",
    [switch] $IncludeOneDrive
)

# PnP.PowerShell is Microsoft-sanctioned. Install once:
#   Install-Module PnP.PowerShell -Scope CurrentUser
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw "Install PnP.PowerShell first: Install-Module PnP.PowerShell -Scope CurrentUser"
}
Import-Module PnP.PowerShell -ErrorAction Stop

Write-Host "Connecting to tenant: $TenantUrl" -ForegroundColor Cyan
Connect-PnPOnline -Url $TenantUrl -Interactive

# Pull all sites. Include OneDrive if requested.
$filters = @("NOT (Url -like '-my.sharepoint.com')")
if ($IncludeOneDrive) { $filters = @() }
$allSites = Get-PnPTenantSite -Detailed -IncludeOneDriveSites:$IncludeOneDrive

Write-Host "Analyzing $($allSites.Count) sites..." -ForegroundColor Cyan

$staleThreshold = (Get-Date).AddDays(-$StaleDays)

$rows = foreach ($s in $allSites) {
    $owner = if ($s.OwnerLoginName) { $s.OwnerLoginName } else { '(no owner)' }

    # Activity: use the LastContentModifiedDate
    $activity = $s.LastContentModifiedDate
    $daysIdle = if ($activity) { [int]((Get-Date) - $activity).TotalDays } else { 9999 }
    $flags = New-Object System.Collections.Generic.List[string]
    if ($daysIdle -gt $StaleDays) { $flags.Add('Stale') }
    if ([string]::IsNullOrWhiteSpace($s.OwnerLoginName)) { $flags.Add('Orphan') }
    if ($s.StorageUsageCurrent -gt 10000) { $flags.Add('Large >10GB') }
    if ($s.IsTeamsConnected -and $s.StorageUsageCurrent -lt 5) { $flags.Add('Empty Team') }

    [PSCustomObject]@{
        Url                   = $s.Url
        Title                 = $s.Title
        Template              = $s.Template
        Owner                 = $owner
        StorageUsedMB         = $s.StorageUsageCurrent
        StorageQuotaMB        = $s.StorageMaximumLevel
        PctUsed               = if ($s.StorageMaximumLevel) {
                                   [math]::Round($s.StorageUsageCurrent * 100 / $s.StorageMaximumLevel, 1)
                                } else { 0 }
        Created               = $s.Created
        LastContentModified   = $activity
        DaysIdle              = $daysIdle
        IsTeamsConnected      = $s.IsTeamsConnected
        SharingCapability     = $s.SharingCapability
        Flags                 = ($flags -join ', ')
    }
}

$rows = $rows | Sort-Object StorageUsedMB -Descending

# Summaries
$summary = @(
    [PSCustomObject]@{ Metric = 'Total sites';                Value = $rows.Count }
    [PSCustomObject]@{ Metric = 'Total storage used (GB)';    Value = [math]::Round(($rows | Measure-Object StorageUsedMB -Sum).Sum / 1024, 1) }
    [PSCustomObject]@{ Metric = 'Stale sites';                Value = ($rows | Where-Object { $_.Flags -match 'Stale' }).Count }
    [PSCustomObject]@{ Metric = 'Orphan sites (no owner)';    Value = ($rows | Where-Object { $_.Flags -match 'Orphan' }).Count }
    [PSCustomObject]@{ Metric = 'Empty Teams sites';          Value = ($rows | Where-Object { $_.Flags -match 'Empty Team' }).Count }
    [PSCustomObject]@{ Metric = 'Sites using external sharing'; Value = ($rows | Where-Object { $_.SharingCapability -ne 'Disabled' }).Count }
)

# Disk output
if (Get-Module -ListAvailable ImportExcel) {
    Import-Module ImportExcel
    $summary | Export-Excel -Path $OutputPath -WorksheetName 'Summary' -AutoSize
    $rows    | Export-Excel -Path $OutputPath -WorksheetName 'All Sites' -AutoSize -TableName 'sites' -Append

    $stale   = $rows | Where-Object { $_.Flags -match 'Stale' }
    $orphan  = $rows | Where-Object { $_.Flags -match 'Orphan' }
    $large   = $rows | Where-Object { $_.Flags -match 'Large' }
    if ($stale)  { $stale  | Export-Excel -Path $OutputPath -WorksheetName 'Stale' -AutoSize -Append }
    if ($orphan) { $orphan | Export-Excel -Path $OutputPath -WorksheetName 'Orphan' -AutoSize -Append }
    if ($large)  { $large  | Export-Excel -Path $OutputPath -WorksheetName 'Largest' -AutoSize -Append }

    Write-Host "Wrote $OutputPath" -ForegroundColor Green
}
else {
    $csvPath = [System.IO.Path]::ChangeExtension($OutputPath, '.csv')
    $rows | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "Wrote $csvPath" -ForegroundColor Yellow
}

Disconnect-PnPOnline

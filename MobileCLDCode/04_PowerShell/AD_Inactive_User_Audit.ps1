<#
.SYNOPSIS
    Active Directory inactive-user audit. Finds accounts that have not logged
    in for N days and generates an offboarding-ready Excel report.

.DESCRIPTION
    WHY THIS IS NOT NATIVE EXCEL / ONEDRIVE
    Active Directory does not sync to Excel. OneDrive can share a file but
    can't query the corporate directory. This is one of the most common
    security-hygiene asks IT gets, and it has to be automated.

    USE CASE (software business):
    IT Security asks Finance every quarter: "which accounts are still active
    but haven't logged in for 90 days?" This script gives them a sortable,
    auditable list in 10 seconds.

.PARAMETER InactiveDays
    How many days without a login counts as inactive. Default: 90.

.PARAMETER OutputPath
    Path to write the .xlsx file.

.EXAMPLE
    ./AD_Inactive_User_Audit.ps1 -InactiveDays 90 -OutputPath .\inactive.xlsx
#>

[CmdletBinding()]
param(
    [int]    $InactiveDays = 90,
    [string] $OutputPath   = ".\inactive_users.xlsx",
    [string] $SearchBase   = $null,
    [switch] $ExcludeServiceAccounts,
    [switch] $DisableFound
)

# Requires RSAT: ActiveDirectory module. Install via:
#   Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    throw "ActiveDirectory PowerShell module not available. Install RSAT first."
}
Import-Module ActiveDirectory -ErrorAction Stop

# ImportExcel is third-party but ships via PSGallery; falls back to CSV if missing.
$hasImportExcel = [bool](Get-Module -ListAvailable ImportExcel)

$threshold = (Get-Date).AddDays(-$InactiveDays)
Write-Host "Querying inactive users (threshold: $threshold)..." -ForegroundColor Cyan

$props = @(
    'SamAccountName','DisplayName','EmailAddress','Enabled','LastLogonDate',
    'PasswordLastSet','whenCreated','Department','Title','Manager',
    'Description','LockedOut','PasswordNeverExpires','MemberOf'
)

$queryParams = @{
    Filter     = 'Enabled -eq $true'
    Properties = $props
}
if ($SearchBase) { $queryParams.SearchBase = $SearchBase }

$all = Get-ADUser @queryParams

# Build report rows
$rows = foreach ($u in $all) {
    $lastLogon = if ($u.LastLogonDate) { $u.LastLogonDate } else { $u.whenCreated }
    $daysInactive = [int]((Get-Date) - $lastLogon).TotalDays

    if ($daysInactive -lt $InactiveDays) { continue }
    if ($ExcludeServiceAccounts -and $u.SamAccountName -like 'svc*') { continue }

    $mgrName = if ($u.Manager) {
        (Get-ADUser $u.Manager -Properties DisplayName).DisplayName
    } else { $null }

    [PSCustomObject]@{
        SamAccountName       = $u.SamAccountName
        DisplayName          = $u.DisplayName
        EmailAddress         = $u.EmailAddress
        Department           = $u.Department
        Title                = $u.Title
        Manager              = $mgrName
        LastLogonDate        = $lastLogon
        DaysInactive         = $daysInactive
        PasswordLastSet      = $u.PasswordLastSet
        LockedOut            = $u.LockedOut
        PasswordNeverExpires = $u.PasswordNeverExpires
        AccountCreated       = $u.whenCreated
        GroupCount           = ($u.MemberOf | Measure-Object).Count
        Description          = $u.Description
    }
}

$rows = $rows | Sort-Object DaysInactive -Descending
Write-Host "Found $($rows.Count) inactive accounts." -ForegroundColor Green

# Bucket summary
$buckets = $rows | Group-Object { switch ($_.DaysInactive) {
        { $_ -ge 365 } { '365+ days' ; break }
        { $_ -ge 180 } { '180-365 days' ; break }
        { $_ -ge 90  } { '90-180 days' ; break }
        default        { '30-90 days' }
    } } | Select-Object Name, Count

# Write output
if ($hasImportExcel) {
    Import-Module ImportExcel
    $rows    | Export-Excel -Path $OutputPath -WorksheetName 'Inactive Users' -AutoSize -TableName 'users'
    $buckets | Export-Excel -Path $OutputPath -WorksheetName 'Summary' -AutoSize -Append
    Write-Host "Wrote $OutputPath" -ForegroundColor Green
}
else {
    $csvPath = [System.IO.Path]::ChangeExtension($OutputPath, '.csv')
    $rows | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "ImportExcel not installed - wrote CSV instead: $csvPath" -ForegroundColor Yellow
}

# Optional: disable flagged accounts (DRY RUN by default)
if ($DisableFound) {
    Write-Host "Disabling flagged accounts..." -ForegroundColor Yellow
    foreach ($r in $rows) {
        try {
            Disable-ADAccount -Identity $r.SamAccountName -Confirm:$false
            Add-ADUserComment -Identity $r.SamAccountName -Comment "Auto-disabled by inactive-user audit on $(Get-Date -Format s)"
        }
        catch {
            Write-Warning "Failed to disable $($r.SamAccountName): $($_.Exception.Message)"
        }
    }
}

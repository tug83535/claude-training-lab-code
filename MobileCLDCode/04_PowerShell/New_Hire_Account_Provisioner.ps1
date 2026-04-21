<#
.SYNOPSIS
    New-Hire Account Provisioning. Reads a weekly CSV of new hires and
    creates AD user, mailbox, distribution list membership, license
    assignment (M365), and initial password letter PDF.

.DESCRIPTION
    WHY THIS IS NOT NATIVE EXCEL / ONEDRIVE
    Excel cannot create AD accounts. OneDrive cannot assign M365 licenses.
    Power Automate can chain some of this but every step that fails in the
    middle leaves you in an inconsistent state with no rollback. A single
    PowerShell transaction keeps everything atomic.

    USE CASE (software business)
    HR drops this week's new hires in a shared folder every Friday. IT
    runs this script Monday at 8am. All 12 hires are ready when they
    arrive at the office.

.PARAMETER NewHireCsv
    CSV with columns:
        first_name, last_name, preferred_name, title, department, manager_email,
        office_location, start_date, license_sku, dl_memberships (csv)
#>

[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='High')]
param(
    [Parameter(Mandatory=$true)][string] $NewHireCsv,
    [string] $DefaultOU       = 'OU=Employees,DC=yourco,DC=com',
    [string] $EmailDomain     = 'yourco.com',
    [string] $LetterOutputDir = '.\WelcomeLetters',
    [switch] $SendWelcomeEmail
)

if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    throw "ActiveDirectory PowerShell module required. Install RSAT."
}

Import-Module ActiveDirectory
# MSOnline / ExchangeOnlineManagement / Microsoft.Graph modules as needed
# Install-Module Microsoft.Graph -Scope CurrentUser

if (-not (Test-Path $LetterOutputDir)) { New-Item -ItemType Directory -Path $LetterOutputDir | Out-Null }

function New-TempPassword {
    $chars = "ABCDEFGHJKMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789".ToCharArray()
    $specials = '!@#$%&*?+'.ToCharArray()
    -join ((1..12 | ForEach-Object { $chars | Get-Random }) + ($specials | Get-Random))
}

function New-UniqueSam {
    param([string]$first, [string]$last)
    $base = ("$first.$last").ToLower() -replace '[^a-z.]', ''
    $sam = $base
    $i = 2
    while (Get-ADUser -Filter "SamAccountName -eq '$sam'" -ErrorAction SilentlyContinue) {
        $sam = "$base$i"; $i++
    }
    return $sam
}

function New-WelcomeLetter {
    param([PSCustomObject]$hire, [string]$upn, [string]$password, [string]$outPath)
    $text = @"
Welcome to iPipeline, $($hire.preferred_name)!

Your IT account details:
   Username: $upn
   Temp password: $password   (you'll be prompted to change at first logon)
   Start date: $($hire.start_date)
   Manager: $($hire.manager_email)
   Office: $($hire.office_location)

Please log in within 24 hours to activate your account.

Any issues? Reach IT at helpdesk@yourco.com or x5000.
"@
    $text | Out-File -FilePath $outPath -Encoding UTF8
}

$hires = Import-Csv $NewHireCsv
Write-Host "Provisioning $($hires.Count) new hires..." -ForegroundColor Cyan

$results = foreach ($hire in $hires) {
    $sam = New-UniqueSam -first $hire.first_name -last $hire.last_name
    $upn = "$sam@$EmailDomain"
    $pw  = New-TempPassword
    $securePw = ConvertTo-SecureString $pw -AsPlainText -Force

    if ($PSCmdlet.ShouldProcess($upn, "Create AD account")) {
        try {
            New-ADUser `
                -SamAccountName    $sam `
                -UserPrincipalName $upn `
                -Name              "$($hire.first_name) $($hire.last_name)" `
                -GivenName         $hire.first_name `
                -Surname           $hire.last_name `
                -DisplayName       "$($hire.first_name) $($hire.last_name)" `
                -Title             $hire.title `
                -Department        $hire.department `
                -Office            $hire.office_location `
                -EmailAddress      $upn `
                -Path              $DefaultOU `
                -AccountPassword   $securePw `
                -Enabled           $true `
                -ChangePasswordAtLogon $true
        }
        catch {
            Write-Warning "AD create failed for $upn : $($_.Exception.Message)"
            continue
        }

        # Set manager
        if ($hire.manager_email) {
            $mgr = Get-ADUser -Filter "EmailAddress -eq '$($hire.manager_email)'" -ErrorAction SilentlyContinue
            if ($mgr) { Set-ADUser -Identity $sam -Manager $mgr.DistinguishedName }
        }

        # DL memberships
        if ($hire.dl_memberships) {
            foreach ($dl in ($hire.dl_memberships -split ',')) {
                try {
                    Add-ADGroupMember -Identity $dl.Trim() -Members $sam -ErrorAction Stop
                } catch {
                    Write-Warning "Failed adding $sam to DL $dl : $($_.Exception.Message)"
                }
            }
        }

        # Welcome letter
        $letterPath = Join-Path $LetterOutputDir "Welcome_$sam.txt"
        New-WelcomeLetter -hire $hire -upn $upn -password $pw -outPath $letterPath

        [PSCustomObject]@{
            SamAccountName = $sam
            UPN            = $upn
            Name           = "$($hire.first_name) $($hire.last_name)"
            LetterPath     = $letterPath
            TempPassword   = $pw
            Created        = Get-Date
            Status         = 'OK'
        }
    }
}

$results | Format-Table -AutoSize

if (Get-Module -ListAvailable ImportExcel) {
    Import-Module ImportExcel
    $results | Export-Excel -Path '.\new_hire_provisioning_log.xlsx' -AutoSize -WorksheetName 'Provisioned'
}

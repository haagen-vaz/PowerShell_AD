

[CmdletBinding()]
param(
    [datetime]$Now = (Get-Date),             # "fejkat nu" 
    [string]$InputPath = ".\ad_export.json", # JSON-filen
    [string]$OutDir = "."                    # vart CSV ska hamna
)

# <-- allt ovanför är "måste vara först" i PowerShell

# nu kan vi ställa in konsolen
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# svensk datum
$sv = [System.Globalization.CultureInfo]::GetCultureInfo("sv-SE")

# ============================

#Läs in JSON
$raw = Get-Content -Path $InputPath -Raw
$data = $raw | ConvertFrom-Json

$domain = $data.domain
$exportDate = [datetime]$data.export_date

Write-Host "Domän: $domain"
Write-Host "Exportdatum: $($exportDate.ToString('yyyy-MM-dd HH:mm',$sv))"
Write-Host "Referensdatum (Nu): $($Now.ToString('yyyy-MM-dd HH:mm',$sv))"
Write-Host ""

# ============================

# Block 2 
# 1) Inaktiva användare (>30 dagar)
$inactiveCutoff = $Now.AddDays(-30)

$inactiveUsers =
$data.users |
Where-Object {
    $ll = [datetime]$_.lastLogon
    $ll -lt $inactiveCutoff
} |
Select-Object `
@{n = 'SamAccountName'; e = { $_.samAccountName } },
@{n = 'DisplayName'   ; e = { $_.displayName } },
@{n = 'Department'    ; e = { $_.department } },
@{n = 'Site'          ; e = { $_.site } },
@{n = 'LastLogon'     ; e = { [datetime]$_.lastLogon } },
@{n = 'DaysInactive'  ; e = { [int]((New-TimeSpan -Start ([datetime]$_.lastLogon) -End $Now).TotalDays) } },
@{n = 'AccountExpires'; e = { $_.accountExpires } }

Write-Host "INAKTIVA ANVÄNDARE (>30 dagar): $($inactiveUsers.Count)"
$inactiveUsers |
Select-Object SamAccountName, DisplayName, Department, Site,
@{n = 'LastLogon'; e = { $_.LastLogon.ToString('yyyy-MM-dd HH:mm', $sv) } },
DaysInactive |
Format-Table -AutoSize
Write-Host ""

# 2) Räkna användare per avdelning 
$deptCounts = @{}
foreach ($u in $data.users) {
    $dept = if ([string]::IsNullOrWhiteSpace($u.department)) { "Okänd" } else { $u.department }
    if (-not $deptCounts.ContainsKey($dept)) {
        $deptCounts[$dept] = 0
    }
    $deptCounts[$dept]++
}

Write-Host "ANVÄNDARE PER AVDELNING"
Write-Host "-----------------------"
$deptCounts.GetEnumerator() |
Sort-Object Name |
ForEach-Object {
    "{0,-20} {1,3} st" -f $_.Key, $_.Value
}
Write-Host ""

# ============================
# Block 3 
Write-Host "DEL B – pipeline, export, sortering"
Write-Host "==================================="
Write-Host ""

# Gruppera datorer per site
$computers = @($data.computers)
$computersBySite = $computers | Group-Object -Property site

Write-Host "DATORER PER SITE"
Write-Host "----------------"
foreach ($grp in $computersBySite) {
    $siteName = if ($grp.Name) { $grp.Name } else { "Okänd site" }
    Write-Host ("{0,-15} {1,3} datorer" -f $siteName, $grp.Count)
}
Write-Host ""

#  Exportera inaktiva användare till CSV
if (-not (Test-Path -LiteralPath $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir | Out-Null
}

$inactiveCsvPath = Join-Path $OutDir "inactive_users.csv"
$inactiveUsers |
Select-Object SamAccountName, DisplayName, Department, Site,
@{n = 'LastLogon'; e = { $_.LastLogon.ToString('yyyy-MM-dd HH:mm', $sv) } },
DaysInactive,
AccountExpires |
Export-Csv -Path $inactiveCsvPath -Encoding UTF8 -NoTypeInformation

Write-Host "CSV skapad: $inactiveCsvPath"
Write-Host ""

# Beräkna lösenordsålder
$usersWithPwdAge =
$data.users |
Select-Object *,
@{n = 'PasswordLastSetDate'; e = { [datetime]$_.passwordLastSet } },
@{n = 'PasswordAgeDays'; e = {
        $pls = [datetime]$_.passwordLastSet
        [int]((New-TimeSpan -Start $pls -End $Now).TotalDays)
    }
}

Write-Host "LÖSENORDSÅLDER – 10 äldsta"
Write-Host "--------------------------"
$usersWithPwdAge |
Sort-Object -Property LastLogonDate -Descending:$false |
Select-Object -First 10 samAccountName, displayName,
@{n = 'PasswordLastSet'; e = { $_.PasswordLastSetDate.ToString('yyyy-MM-dd', $sv) } },
PasswordAgeDays |
Format-Table -AutoSize
Write-Host ""

# 10 datorer som inte setts på längst tid
$top10OldestComputers =
$computers |
Select-Object *,
@{n = 'LastLogonDate'; e = { [datetime]$_.lastLogon } },
@{n = 'DaysSinceSeen'; e = { [int]((New-TimeSpan -Start ([datetime]$_.lastLogon) -End $Now).TotalDays) } } |
Sort-Object -Property LastLogonDate |
Select-Object -First 10

Write-Host "10 DATORER SOM INTE CHECKAT IN PÅ LÄNGST TID"
Write-Host "--------------------------------------------"
$top10OldestComputers |
Select-Object name,
@{n = 'LastLogon'; e = { $_.LastLogonDate.ToString('yyyy-MM-dd HH:mm', $sv) } },
DaysSinceSeen,
site |
Format-Table -AutoSize
Write-Host ""

# ============================
# Block 4 
# ============================

Write-Host "DEL C – funktioner, executive summary, felhantering"
Write-Host "==================================================="
Write-Host ""

# Funktion: Get-InactiveAccounts
function Get-InactiveAccounts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$Users,

        [Parameter(Mandatory)]
        [int]$Days,

        [Parameter(Mandatory)]
        [datetime]$RefDate
    )

    # vi räknar ut en cutoff utifrån RefDate (fejk)
    $cutoff = $RefDate.AddDays(-$Days)

    $Users |
    ForEach-Object {
        # vi hanterar datum inne i funktionen
        $lastLogon = $null
        try {
            if ($_.lastLogon) {
                $lastLogon = [datetime]$_.lastLogon
            }
        }
        catch {
            # om ett datum är trasigt vill vi INTE att skriptet dör
            Write-Verbose "Kunde inte tolka datum för användare $($_.samAccountName): $_"
        }

        # returnera bara om vi lyckades få datum OCH användaren är äldre än cutoff
        if ($lastLogon -and $lastLogon -lt $cutoff) {
            # skicka vidare själva objektet
            $_
        }
    }
}

# 2) Robust datumparsing – hjälpfunktion
function Convert-ToDateSafe {
    param(
        [string]$Value
    )
    try {
        if ([string]::IsNullOrWhiteSpace($Value)) {
            return $null
        }
        return [datetime]$Value
    }
    catch {
        # try/catch
        Write-Verbose "Ogiltigt datum hittat i JSON: '$Value'"
        return $null
    }
}

# Använd funktionerna för att ta fram siffror till executive summary

# inaktiva användare 
$inactive30 = Get-InactiveAccounts -Users $data.users -Days 30 -RefDate $Now

# konton som löper ut inom 30 dagar
$accountsExpiringSoon =
$data.users |
ForEach-Object {
    $exp = Convert-ToDateSafe $_.accountExpires
    if ($exp -and $exp -ge $Now -and $exp -le $Now.AddDays(30) -and $_.enabled) {
        $_
    }
}

# datorer som inte setts på 30+ dagar
$computersNotSeen30 =
$computers |
ForEach-Object {
    $ll = Convert-ToDateSafe $_.lastLogon
    if ($ll -and $ll -lt $Now.AddDays(-30) -and $_.enabled) {
        $_
    }
}

# användare med lösenord äldre än 90 dagar ELLER passwordNeverExpires
$pwdOldUsers =
$data.users |
ForEach-Object {
    $pls = Convert-ToDateSafe $_.passwordLastSet
    $age = $null
    if ($pls) {
        $age = [int]((New-TimeSpan -Start $pls -End $Now).TotalDays)
    }

    if ( ($age -and $age -ge 90) -or ($_.passwordNeverExpires -eq $true) ) {
        $_
    }
}

# Bygg executive summary
$summary = @()
if ($accountsExpiringSoon.Count -gt 0) {
    $summary += "CRITICAL: $($accountsExpiringSoon.Count) konton löper ut inom 30 dagar"
}
else {
    $summary += "OK: Inga konton som löper ut inom 30 dagar"
}

if ($inactive30.Count -gt 0) {
    $summary += "WARNING: $($inactive30.Count) användare har inte loggat in på 30+ dagar"
}
else {
    $summary += "OK: Inga inaktiva användare (30+ dagar)"
}

if ($computersNotSeen30.Count -gt 0) {
    $summary += "WARNING: $($computersNotSeen30.Count) datorer har inte setts på 30+ dagar"
}
else {
    $summary += "OK: Alla datorer sedda senaste 30 dagarna"
}

if ($pwdOldUsers.Count -gt 0) {
    $summary += "SECURITY: $($pwdOldUsers.Count) användare har lösenord äldre än 90 dagar eller 'never expires'"
}
else {
    $summary += "OK: Inga gamla lösenord"
}

Write-Host "EXECUTIVE SUMMARY"
Write-Host "-----------------"
$summary | ForEach-Object { Write-Host $_ }
Write-Host ""

# Spara en enkel text-rapport
$reportPath = Join-Path $OutDir "ad_audit_report.txt"

$reportText = @"
============================================================
ACTIVE DIRECTORY AUDIT REPORT
Generated   : $($Now.ToString('yyyy-MM-dd HH:mm',$sv))
Domain      : $domain
Export date : $($exportDate.ToString('yyyy-MM-dd HH:mm',$sv))
============================================================

EXECUTIVE SUMMARY
-----------------
$($summary -join "`r`n")

Inaktiva användare (>30 dagar): $($inactive30.Count)
Konton som löper ut (30 dagar): $($accountsExpiringSoon.Count)
Datorer ej sedda (30 dagar)    : $($computersNotSeen30.Count)
Gamla lösenord (90+ dagar)     : $($pwdOldUsers.Count)

"@

$reportText | Out-File -FilePath $reportPath -Encoding utf8
Write-Host "Rapport sparad: $reportPath"



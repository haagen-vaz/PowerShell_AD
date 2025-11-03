[CmdletBinding()]
param(
    [datetime]$Now = (Get-Date),
    [string]$InputPath = ".\ad_export.json",
    [string]$OutDir = "."
)

# Ställ in UTF-8 
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$sv = [System.Globalization.CultureInfo]::GetCultureInfo("sv-SE")

# Läs in JSON-data
try {
    $raw = Get-Content -Path $InputPath -Raw -ErrorAction Stop
    $data = $raw | ConvertFrom-Json
    # Om användaren inte skickat in -Now, använd export_date
    if (-not $PSBoundParameters.ContainsKey('Now')) {
        $Now = [datetime]$data.export_date
    }

}
catch {
    Write-Error "Kunde inte läsa JSON: $_"
    exit 1
}

# Grunddata
$domain = $data.domain
$exportDate = [datetime]$data.export_date
$users = @($data.users)
$computers = @($data.computers)

# Konverterar text till datum 
function Convert-ToDateSafe {
    param([string]$Value)
    try {
        if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
        return [datetime]$Value
    }
    catch { return $null }
}

# Returnerar användare inaktiva mer än X dagar
function Get-InactiveAccounts {
    param([array]$Users, [int]$Days, [datetime]$RefDate)
    $cut = $RefDate.AddDays(-$Days)
    $Users | ForEach-Object {
        $ll = Convert-ToDateSafe $_.lastLogon
        if ($ll -and $ll -lt $cut) { $_ }
    }
}

# Inaktiva användare (>30 dagar)
$inactive30 = Get-InactiveAccounts -Users $users -Days 30 -RefDate $Now

# Räkna användare per avdelning
$deptCounts = @{}
foreach ($u in $users) {
    $dept = if ([string]::IsNullOrWhiteSpace($u.department)) { "Okänd" } else { $u.department }
    if (-not $deptCounts.ContainsKey($dept)) { $deptCounts[$dept] = 0 }
    $deptCounts[$dept]++
}

# Grupp av datorer per site
$computersBySite = $computers | Group-Object -Property site

# Skapa output-mapp om den saknas
if (-not (Test-Path -LiteralPath $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir | Out-Null
}

# Exportera inaktiva användare till CSV
$inactiveCsvPath = Join-Path $OutDir "inactive_users.csv"
$inactive30 |
Select-Object `
@{n = 'Konto'; e = { $_.samAccountName } },
@{n = 'Namn'; e = { $_.displayName } },
@{n = 'Avdelning'; e = { $_.department } },
@{n = 'Site'; e = { $_.site } },
@{n = 'SenastInlogg'; e = {
        $ll = Convert-ToDateSafe $_.lastLogon
        if ($ll) { $ll.ToString('yyyy-MM-dd HH:mm', $sv) } else { '' }
    }
},
@{n = 'DagarInaktiv'; e = {
        $ll = Convert-ToDateSafe $_.lastLogon
        if ($ll) { [int]((New-TimeSpan -Start $ll -End $Now).TotalDays) } else { $null }
    }
},
AccountExpires |
Export-Csv -Path $inactiveCsvPath -NoTypeInformation -Encoding UTF8

# Beräkna lösenordsålder
$usersWithPwdAge =
$users | Select-Object *,
@{n = 'PwdDate'; e = { Convert-ToDateSafe $_.passwordLastSet } },
@{n = 'PwdAge'; e = {
        $pd = Convert-ToDateSafe $_.passwordLastSet
        if ($pd) { [int]((New-TimeSpan -Start $pd -End $Now).TotalDays) } else { $null }
    }
}

# Hitta användare med gamla lösenord
$pwdOldUsers = $usersWithPwdAge | Where-Object {
    ($_.PwdAge -and $_.PwdAge -ge 90) -or ($_.passwordNeverExpires -eq $true)
}

# Datorer sorterade efter senaste inloggning (äldst först)
$top10Old =
$computers |
Select-Object *,
@{n = 'LastSeen'; e = { Convert-ToDateSafe $_.lastLogon } },
@{n = 'DaysSinceSeen'; e = {
        $ls = Convert-ToDateSafe $_.lastLogon
        if ($ls) { (New-TimeSpan -Start $ls.Date -End $Now.Date).Days } else { $null }
    }
} |
Sort-Object -Property LastSeen |
Select-Object -First 10


# Datorer ej sedda på 30+ dagar
$computersNotSeen30 =
$computers | ForEach-Object {
    $ls = Convert-ToDateSafe $_.lastLogon
    if ($ls -and $ls -lt $Now.AddDays(-30) -and $_.enabled) { $_ }
}

# Konton som löper ut inom 30 dagar
$accountsExpiringSoon =
$users | ForEach-Object {
    $exp = Convert-ToDateSafe $_.accountExpires
    if ($exp -and $exp -ge $Now -and $exp -le $Now.AddDays(30) -and $_.enabled) { $_ }
}

# Sammanfattning
$summaryLines = @()
if ($accountsExpiringSoon.Count -gt 0) {
    $summaryLines += "CRITICAL: $($accountsExpiringSoon.Count) konton som löper ut inom 30 dagar"
}
else {
    $summaryLines += "OK: Inga konton som löper ut inom 30 dagar"
}
if ($inactive30.Count -gt 0) {
    $summaryLines += "WARNING: $($inactive30.Count) användare har inte loggat in på 30+ dagar"
}
else {
    $summaryLines += "OK: Inga inaktiva användare (30+ dagar)"
}
if ($computersNotSeen30.Count -gt 0) {
    $summaryLines += "WARNING: $($computersNotSeen30.Count) datorer har inte setts på 30+ dagar"
}
else {
    $summaryLines += "OK: Alla datorer sedda senaste 30 dagarna"
}
if ($pwdOldUsers.Count -gt 0) {
    $summaryLines += "SECURITY: $($pwdOldUsers.Count) användare har lösenord äldre än 90 dagar eller 'never expires'"
}
else {
    $summaryLines += "OK: Inga gamla lösenord"
}

# Bygg rapport
$header = @"
===============================================================================
ACTIVE DIRECTORY AUDIT REPORT
===============================================================================
"@

$report = @()
$report += $header
$report += "Generated   : $($Now.ToString('yyyy-MM-dd HH:mm',$sv))"
$report += "Domain      : $domain"
$report += "Export date : $($exportDate.ToString('yyyy-MM-dd HH:mm',$sv))"
$report += "Kontrollperiod : $($Now.ToString('yyyy-MM-dd',$sv)) – $($Now.AddDays(30).ToString('yyyy-MM-dd',$sv))"
$report += "Source file : $InputPath"
$report += ""
$report += @"
EXECUTIVE SUMMARY
-----------------
"@
$report += $summaryLines
$report += ""

# Avdelningsstatistik
$report += "ANVÄNDARE PER AVDELNING"
$report += "-----------------------"
$deptCounts.GetEnumerator() | Sort-Object Name | ForEach-Object {
    $report += ("{0,-15}{1,3} st" -f $_.Key, $_.Value)
}
$report += ""

# Tabell för inaktiva användare
$report += @"
INAKTIVA ANVÄNDARE
-----------------------------------------------
"@
$report += ("{0,-12}{1,-20}{2,-15}{3,-13}{4,5}" -f "Konto", "Namn", "Avdelning", "Senast", "Dagar")
$report += ("{0,-12}{1,-20}{2,-15}{3,-13}{4,5}" -f "-----", "----", "---------", "------", "-----")

foreach ($u in ($inactive30 | Sort-Object {
            $last = Convert-ToDateSafe $_.lastLogon
            if ($last) { [int]((New-TimeSpan -Start $last -End $Now).TotalDays) } else { 0 }
        } -Descending)) {

    $last = Convert-ToDateSafe $u.lastLogon
    $lastStr = if ($last) { $last.ToString('yyyy-MM-dd', $sv) } else { "N/A" }
    $days = if ($last) { [int]((New-TimeSpan -Start $last -End $Now).TotalDays) } else { 0 }
    $dept = if ([string]::IsNullOrWhiteSpace($u.department)) { "-" } else { $u.department }

    $report += ("{0,-12}{1,-20}{2,-15}{3,-13}{4,5}" -f `
            $u.samAccountName, $u.displayName, $dept, $lastStr, $days)
}
$report += ""



# Tabell för användare med gamla lösenord
$report += @"
ANVÄNDARE MED GAMLA LÖSENORD
----------------------------
"@
$report += ("{0,-12}{1,-22}{2,-16}{3,-18}{4,10}" -f "Konto", "Namn", "Avdelning", "Senast ändrat", "Dagar")
$report += ("{0,-12}{1,-22}{2,-16}{3,-18}{4,10}" -f "-----", "----", "-----------", "--------------", "-----")

foreach ($u in ($pwdOldUsers | Sort-Object -Property PwdAge -Descending)) {
    $pwdDate = Convert-ToDateSafe $u.passwordLastSet
    $pwdStr = if ($pwdDate) { $pwdDate.ToString('yyyy-MM-dd', $sv) } else { "N/A" }
    $days = if ($u.PwdAge) { $u.PwdAge } else { "" }
    $dept = if ([string]::IsNullOrWhiteSpace($u.department)) { "-" } else { $u.department }

    $report += ("{0,-12}{1,-22}{2,-16}{3,-18}{4,10}" -f `
            $u.samAccountName, $u.displayName, $dept, $pwdStr, $days)
}
$report += ""


# 10 äldst inloggade datorer
$report += @"
10 DATORER SOM INTE SETTS PÅ LÄNGST TID
---------------------------------------
"@
$report += ("{0,-16}{1,-13}{2,8} {3,-18}" -f "Namn", "Senast sedd", "Dagar", "Site")
$report += ("{0,-16}{1,-13}{2,8}{3,-18}" -f "----", "------------", "-----", "------")

foreach ($c in $top10Old) {
    $ls = $c.LastSeen
    $lsStr = if ($ls) { $ls.ToString('yyyy-MM-dd', $sv) } else { "N/A" }
    $days = if ($c.DaysSinceSeen) { $c.DaysSinceSeen } else { "" }
    $site = if ($c.site) { $c.site } else { "-" }

    $report += ("{0,-16}{1,-13}{2,6}   {3,-18}" -f $c.name, $lsStr, $days, $site)
}
$report += ""

# Tabell för konton som löper ut inom 30 dagar
$report += @"
KONTON SOM LÖPER UT INOM 30 DAGAR
---------------------------------
"@
$report += ("{0,-12}{1,-22}{2,-16}{3,-18}" -f "Konto", "Namn", "Avdelning", "Löper ut")
$report += ("{0,-12}{1,-22}{2,-16}{3,-18}" -f "-----", "----", "-----------", "---------")

foreach ($u in ($accountsExpiringSoon | Sort-Object -Property accountExpires)) {
    $expDate = Convert-ToDateSafe $u.accountExpires
    $expStr = if ($expDate) { $expDate.ToString('yyyy-MM-dd', $sv) } else { "N/A" }
    $dept = if ([string]::IsNullOrWhiteSpace($u.department)) { "-" } else { $u.department }

    $report += ("{0,-12}{1,-22}{2,-16}{3,-18}" -f `
            $u.samAccountName, $u.displayName, $dept, $expStr)
}
$report += ""



# Skriv rapport och CSV
$reportPath = Join-Path $OutDir "ad_audit_report.txt"
$report | Out-File -FilePath $reportPath -Encoding utf8
Write-Host "Rapport sparad: $reportPath"
Write-Host "CSV sparad    : $inactiveCsvPath"

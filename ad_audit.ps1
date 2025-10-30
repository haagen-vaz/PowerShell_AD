

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


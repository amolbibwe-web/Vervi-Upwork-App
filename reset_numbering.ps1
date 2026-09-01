# reset_numbering.ps1 -- put document numbering back to a known starting point.
#
# Test runs consume real document numbers, because the registry deliberately
# never reissues one. Use this after testing to set the opening position again.
#
#   .\reset_numbering.ps1                  # Sales resumes at 023, JE at 001
#   .\reset_numbering.ps1 -SalesFrom 45    # Sales resumes at 045, JE at 001
#   .\reset_numbering.ps1 -SalesFrom 45 -JeFrom 12
#
# It rewrites master\doc_registry.csv with a single opening row per series, so
# the next number issued is exactly the one you asked for.

param(
    [int]$SalesFrom = 36,
    [string]$SalesPrefix = "26-27/VASL/U/",
    # JE numbering restarts every month, so an opening position is per month.
    # Months not listed here simply start at 001.
    [hashtable]$JeOpenings = @{ "26-27/LLP/Aug/" = 63 },
    [string]$OpeningDate = "2026-04-01"
)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$registry = Join-Path $PSScriptRoot "master\doc_registry.csv"
$stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$lines = @("document_number,type,date,prefix,seq,issued_at,source")

# An opening row records "this many are already spoken for", so the next number
# issued is SalesFrom. Seeding 22 makes the next one 023.
$salesSeed = $SalesFrom - 1
if ($salesSeed -ge 1) {
    $num = "{0}{1:D3}" -f $SalesPrefix, $salesSeed
    $lines += "$num,Sales,$OpeningDate,$SalesPrefix,$salesSeed,$stamp,opening position - next Sales number is $('{0:D3}' -f $SalesFrom)"
}

# One opening row per month that already has JE numbers behind it.
foreach ($prefix in $JeOpenings.Keys) {
    $seq = [int]$JeOpenings[$prefix]
    if ($seq -lt 1) { continue }
    $num = "{0}{1:D3}" -f $prefix, $seq
    $lines += "$num,JE,$OpeningDate,$prefix,$seq,$stamp,opening position - last JE actually issued was $('{0:D3}' -f $seq)"
}

if (Test-Path $registry) {
    $backup = Join-Path $PSScriptRoot ("master\doc_registry.backup-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".csv")
    Copy-Item $registry $backup -Force
    Write-Host "  previous registry saved to $(Split-Path $backup -Leaf)" -ForegroundColor DarkGray
}

New-Item -ItemType Directory -Force -Path (Split-Path $registry) | Out-Null
[System.IO.File]::WriteAllLines($registry, $lines, (New-Object System.Text.UTF8Encoding($false)))

Write-Host ""
Write-Host "  Numbering reset." -ForegroundColor Green
Write-Host ("    next Sales : {0}{1:D3}" -f $SalesPrefix, $SalesFrom)
foreach ($prefix in $JeOpenings.Keys) {
    Write-Host ("    next JE    : {0}{1:D3}" -f $prefix, ([int]$JeOpenings[$prefix] + 1))
}
Write-Host "    next JE    : <other months> 001"
Write-Host ""
Write-Host "  Restart the app so it picks this up:" -ForegroundColor DarkGray
Write-Host "    Stop-ScheduledTask -TaskName 'Vervi-Upwork'; Start-ScheduledTask -TaskName 'Vervi-Upwork'" -ForegroundColor DarkGray
Write-Host ""

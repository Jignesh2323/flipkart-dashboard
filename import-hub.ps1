# Imports Seller Hub data for accounts without API access (VUSTICA, BRAINLE) that was collected
# in the browser and parked in Supabase (fk_blob "settle_<acc>"). Writes the same files the other
# accounts use (settlements\<acc>.json, daily\, charged\) plus hubsales.js (settled sales by
# order date + SKU, used for their product cost since there are no API orders).
param([string[]]$Accounts = @('Acc1_VUSTICA', 'Acc2_BRAINLE'))
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$key = 'sb_publishable_HO0wPEI58Ia5k7CFOIyOhA_A6dtCsuL'
$h = @{ apikey = $key; Authorization = "Bearer $key" }
$utf8 = New-Object System.Text.UTF8Encoding $false
$sales = [ordered]@{}
# Per payment date + SKU: [date, sku, sales paid, sale settlement, returns, return settlement, return charge]
$items = [ordered]@{}
foreach ($acc in $Accounts) {
    $r = Invoke-RestMethod -Headers $h -Uri "https://begblflwhxbbipsmxytd.supabase.co/rest/v1/fk_blob?name=eq.settle_$acc&select=data"
    if (-not $r) { Write-Host "$acc : Supabase me data nahi"; continue }
    $o = $r[0].data | ConvertFrom-Json
    [IO.File]::WriteAllText((Join-Path $root "settlements\$acc.json"), ($o.main | ConvertTo-Json -Depth 8 -Compress), $utf8)
    [IO.File]::WriteAllText((Join-Path $root "settlements\daily\$acc.json"), ($o.dd | ConvertTo-Json -Depth 4 -Compress), $utf8)
    [IO.File]::WriteAllText((Join-Path $root "settlements\charged\$acc.json"), (ConvertTo-Json @($o.cr) -Depth 4 -Compress), $utf8)
    $sales[$acc] = $o.sl
    if ($o.it) { $items[$acc] = $o.it }
    Write-Host "$acc : $($o.main.p.Count) payment din, $($o.cr.Count) charged returns, $($o.sl.Count) sale rows"
}
# Keep accounts imported earlier that were not part of this run
$hs = Join-Path $root 'hubsales.js'
if (Test-Path $hs) {
    $old = (Get-Content $hs -Raw); $old = $old.Substring($old.IndexOf('{')).TrimEnd().TrimEnd(';') | ConvertFrom-Json
    foreach ($p in $old.PSObject.Properties) { if (-not $sales.Contains($p.Name)) { $sales[$p.Name] = $p.Value } }
}
[IO.File]::WriteAllText($hs, "window.FK_HUBSALES = " + ($sales | ConvertTo-Json -Depth 5 -Compress) + ";", $utf8)
Write-Host "hubsales.js: $($sales.Count) accounts"
if ($items.Count) {
    $hi = Join-Path $root 'hubitems.js'
    if (Test-Path $hi) {
        $old = (Get-Content $hi -Raw); $old = $old.Substring($old.IndexOf('{')).TrimEnd().TrimEnd(';') | ConvertFrom-Json
        foreach ($p in $old.PSObject.Properties) { if (-not $items.Contains($p.Name)) { $items[$p.Name] = $p.Value } }
    }
    [IO.File]::WriteAllText($hi, "window.FK_HUBITEMS = " + ($items | ConvertTo-Json -Depth 5 -Compress) + ";", $utf8)
    Write-Host "hubitems.js: $($items.Count) accounts"
}

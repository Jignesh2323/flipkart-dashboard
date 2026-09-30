# FSN (Flipkart product id) of every SKU that has orders, per account, via the Listings API.
# The dashboard's Mapping tab links each FSN to its flipkart.com sellers page.
# Keeps fsn.json as a cache so only new SKUs are looked up; writes fsn.js (window.FK_FSN).
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$api  = 'https://api.flipkart.net'

function Get-Token($appId, $secret) {
    $b64 = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${appId}:${secret}"))
    (Invoke-RestMethod -Method Get -Headers @{ Authorization = "Basic $b64" } `
        -Uri "$api/oauth-service/oauth/token?grant_type=client_credentials&scope=Seller_Api").access_token
}

$cacheFile = Join-Path $root 'fsn.json'
$cache = @{}
if (Test-Path $cacheFile) { (Get-Content $cacheFile -Raw | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $cache[$_.Name] = $_.Value } }

$raw = Get-Content (Join-Path $root 'data.js') -Raw -Encoding UTF8
$D = $raw.Substring($raw.IndexOf('{')).TrimEnd().TrimEnd(';') | ConvertFrom-Json
$skusByAcc = @{}
foreach ($o in $D.orders) { if ($o.sku) { if (-not $skusByAcc[$o.acc]) { $skusByAcc[$o.acc] = @{} }; $skusByAcc[$o.acc][$o.sku] = $o.title } }

foreach ($a in Import-Csv (Join-Path $root 'accounts.csv')) {
    $acc = $a.Name; if (-not $skusByAcc[$acc]) { continue }
    $todo = @($skusByAcc[$acc].Keys | Where-Object { -not $cache["$acc|$_"] })
    Write-Host "$acc : $($skusByAcc[$acc].Count) SKUs, $($todo.Count) naye"
    if (-not $todo.Count) { continue }
    try { $tok = Get-Token $a.AppId $a.AppSecret } catch { Write-Host "  token FAIL: $($_.Exception.Message)"; continue }
    $h = @{ Authorization = "Bearer $tok" }
    for ($i = 0; $i -lt $todo.Count; $i += 10) {
        $batch = $todo[$i..([Math]::Min($i + 9, $todo.Count - 1))]
        $ids = ($batch | ForEach-Object { [Uri]::EscapeDataString($_) }) -join ','
        try {
            $r = Invoke-RestMethod -Method Get -Headers $h -Uri "$api/sellers/listings/v3/$ids"
            foreach ($p in $r.available.PSObject.Properties) { $cache["$acc|$($p.Name)"] = $p.Value.product_id }
        } catch { Write-Host "  batch FAIL: $($_.Exception.Message)" }
    }
}
$cache | ConvertTo-Json -Depth 3 | Set-Content $cacheFile -Encoding UTF8
$out = [ordered]@{}; foreach ($k in ($cache.Keys | Sort-Object)) { $out[$k] = $cache[$k] }
"window.FK_FSN = " + ($out | ConvertTo-Json -Compress) + ";" | Set-Content (Join-Path $root 'fsn.js') -Encoding UTF8
Write-Host "fsn.js: $($out.Count) SKUs"

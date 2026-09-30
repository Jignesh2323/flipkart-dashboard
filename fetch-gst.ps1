# GST rate per Flipkart SKU from the invoice API (taxRate on the invoice line).
# One delivered shipment per SKU is enough; rates already in gst.json are not asked again.
# Writes gst.json (cache) and gst.js ({ sku: rate }) for dashboard.html.
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$api = 'https://api.flipkart.net'
$cachePath = Join-Path $root 'gst.json'
$rates = @{}
if (Test-Path $cachePath) { (Get-Content $cachePath -Raw | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $rates[$_.Name] = $_.Value } }

$accounts = Import-Csv (Join-Path $root 'accounts.csv') | Where-Object { $_.AppId -and $_.AppId -notmatch 'PASTE' }
foreach ($a in $accounts) {
    $file = Join-Path $root "raw\$($a.Name)_shipments.json"
    if (-not (Test-Path $file)) { continue }
    # one delivered shipment for every SKU we don't know yet
    $pick = @{}
    foreach ($s in (Get-Content $file -Raw | ConvertFrom-Json)) {
        foreach ($i in $s.orderItems) {
            if ($i.sku -and $i.status -eq 'DELIVERED' -and -not $rates.ContainsKey($i.sku) -and -not $pick.ContainsKey($i.sku)) { $pick[$i.sku] = $s.shipmentId }
        }
    }
    if (-not $pick.Count) { continue }
    try {
        $b64 = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($a.AppId):$($a.AppSecret)"))
        $tok = (Invoke-RestMethod -Headers @{ Authorization = "Basic $b64" } -Uri "$api/oauth-service/oauth/token?grant_type=client_credentials&scope=Seller_Api").access_token
    } catch { Write-Host "$($a.Name): login FAIL" -ForegroundColor Yellow; continue }
    $skuOfShip = @{}; foreach ($k in $pick.Keys) { $skuOfShip[$pick[$k]] = $k }
    $ids = @($skuOfShip.Keys)
    for ($n = 0; $n -lt $ids.Count; $n += 20) {
        $batch = $ids[$n..([Math]::Min($n + 19, $ids.Count - 1))] -join ','
        try {
            $r = Invoke-RestMethod -Headers @{ Authorization = "Bearer $tok" } -Uri "$api/sellers/v3/shipments/$batch/invoices"
            foreach ($inv in $r.invoices) { foreach ($li in $inv.orderItems) { if ($null -ne $li.taxRate) { $rates[$skuOfShip[$inv.shipmentId]] = [double]$li.taxRate } } }
        } catch { Write-Host "  $($a.Name) invoices: $($_.Exception.Message)" -ForegroundColor Yellow }
        Start-Sleep -Milliseconds 300
    }
    Write-Host "$($a.Name): $($pick.Count) naye SKU ka GST dekha"
}
$rates | ConvertTo-Json -Compress | Out-File $cachePath -Encoding utf8
"window.FK_GST = " + ($rates | ConvertTo-Json -Compress) + ";" | Out-File (Join-Path $root 'gst.js') -Encoding utf8
Write-Host "gst.js: $($rates.Count) SKUs"

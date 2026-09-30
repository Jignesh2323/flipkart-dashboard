# Creates / refreshes product-cost.csv: one row per product (account prefix and listing number
# removed from the Flipkart SKU), with a Cost column the user fills in by hand.
# Existing costs are kept; new products are added with an empty cost.
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $root 'sku-key.ps1')
$t = Get-Content (Join-Path $root 'data.js') -Raw -Encoding utf8
$j = $t.Substring($t.IndexOf('{'), $t.LastIndexOf('}') - $t.IndexOf('{') + 1) | ConvertFrom-Json

$csv = Join-Path $root 'product-cost.csv'
$old = @{}
if (Test-Path $csv) { Import-Csv $csv | ForEach-Object { $old[$_.Product] = $_.Cost } }

# units sold per product, so the sheet lists the busiest products first
$units = @{}; $skus = @{}
foreach ($o in $j.orders) {
    if (-not $o.sku) { continue }
    $p = Get-ProductName $o.sku
    if ($o.status -notmatch 'CANCEL') { $units[$p] = [int]$units[$p] + [int]$o.qty }
    if (-not $skus[$p]) { $skus[$p] = @{} }; $skus[$p][$o.sku] = 1
}
foreach ($r in $j.returns) { if ($r.sku) { $p = Get-ProductName $r.sku; if (-not $skus[$p]) { $skus[$p] = @{} }; $skus[$p][$r.sku] = 1 } }
foreach ($p in $old.Keys) { if (-not $skus.ContainsKey($p)) { $skus[$p] = @{} } }

$rows = $skus.Keys | Sort-Object { - [int]$units[$_] }, { $_ } | ForEach-Object {
    [pscustomobject]@{ Product = $_; Cost = $old[$_]; 'Units Sold (6 mahine)' = [int]$units[$_]; 'Flipkart SKUs' = (($skus[$_].Keys | Sort-Object) -join ' | ') }
}
$rows | Export-Csv $csv -NoTypeInformation -Encoding UTF8
$empty = @($rows | Where-Object { -not $_.Cost }).Count
Write-Host "product-cost.csv: $($rows.Count) products, $empty ki cost abhi khaali hai"

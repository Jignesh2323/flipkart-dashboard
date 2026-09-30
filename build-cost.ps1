# product-cost.csv (filled by hand) -> cost.js for dashboard.html: { "PRODUCT NAME": cost }
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$csv = Join-Path $root 'product-cost.csv'
$map = [ordered]@{}
if (Test-Path $csv) {
    foreach ($r in Import-Csv $csv) {
        $c = "$($r.Cost)" -replace '[^\d.]', ''
        if ($c -and [double]$c -gt 0) { $map[$r.Product] = [double]$c }
    }
}
"window.FK_COST = " + ($map | ConvertTo-Json -Compress) + ";" | Out-File (Join-Path $root 'cost.js') -Encoding utf8
Write-Host "cost.js: $($map.Count) products ki cost"

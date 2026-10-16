# Builds targets.json (settlement Min/Max per account + SKU) from settlement-target-product.xlsx.
# Same rules as the keeper: per-listing K/L, else the product table Min/Max of its Product Type.
# targets.json is published with the dashboard (GitHub Pages) and read by the Price tab and by the
# Seller Hub keeper helpers. Re-run after editing the sheet, then commit + push.
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$tmp = Join-Path $env:TEMP 'targets-sheet.xlsx'
Copy-Item (Join-Path $root 'settlement-target-product.xlsx') $tmp -Force
$xl = New-Object -ComObject Excel.Application; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($tmp, 0, $true); $v = $wb.Worksheets.Item(1).UsedRange.Value2
$wb.Close($false); $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
$num = { param($x) $d = 0.0; if ($null -ne $x -and [double]::TryParse("$x", [ref]$d)) { $d } else { $null } }
$types = @{}
for ($r = 3; $r -le 13; $r++) { $p = "$($v[$r,1])"; if ($p) { $types[$p] = @{ min = & $num $v[$r,5]; max = & $num $v[$r,6] } } }
$T = [ordered]@{}; $n = 0
for ($r = 15; $r -le $v.GetLength(0); $r++) {
    $sku = "$($v[$r,2])".Trim(); $acc = "$($v[$r,1])".Trim().ToUpper(); if (-not $sku -or -not $acc) { continue }
    $pt = "$($v[$r,3])"; $min = & $num $v[$r,11]; $max = & $num $v[$r,12]
    if ($null -eq $min -and $types[$pt]) { $min = $types[$pt].min; $max = $types[$pt].max }
    if ($null -eq $min -or $null -eq $max) { continue }
    if (-not $T.Contains($acc)) { $T[$acc] = [ordered]@{} }
    $T[$acc][$sku] = @([int]$min, [int]$max); $n++
}
[IO.File]::WriteAllText((Join-Path $root 'targets.json'), ($T | ConvertTo-Json -Depth 4 -Compress), (New-Object System.Text.UTF8Encoding $false))
Write-Host "targets.json: $n listings"

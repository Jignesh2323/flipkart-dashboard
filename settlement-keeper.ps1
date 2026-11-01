# Daily settlement keeper.
# Reads target Min/Max settlement per listing from settlement-target-product.xlsx (product table on top,
# listings from row 15), reads current price of every listing via the official Listings API, estimates the
# bank settlement as price - fees (fees = col J "Fees (Price-Settlement)" from Seller Hub; for a new row
# without fees the median fees of its product type is used), and moves the price so the settlement comes
# back inside Min..Max when it is more than $Tol below Min or above Max.
# New listing: add a row in the sheet (Account, SKU, Product Type) - its Min/Max come from the product table.
# -Publish: use the latest per-listing fees saved from Seller Hub (fees.json in this folder,
# refreshed whenever the keeper runs in Seller Hub) and write the per-account in/out-of-target counts to
# price.json (pushed to the dashboard repo), which the dashboard's Price tab shows.
# Usage: settlement-keeper.ps1 [-DryRun] [-Lower] [-Publish]
param([switch]$DryRun, [switch]$Lower, [switch]$Publish, [int]$Tol = 5, [string]$Only = '')
$ErrorActionPreference = 'Continue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$api  = 'https://api.flipkart.net'
$log  = Join-Path $root 'settlement-keeper.log'
function Log($m) { $l = "$(Get-Date -f 'yyyy-MM-dd HH:mm') $m"; Write-Host $l; Add-Content $log $l -Encoding UTF8 }

# --- read sheet (copy first so an open workbook does not block us) ---
$tmp = Join-Path $env:TEMP 'keeper-sheet.xlsx'
Copy-Item (Join-Path $root 'settlement-target-product.xlsx') $tmp -Force
$xl = New-Object -ComObject Excel.Application; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($tmp, 0, $true); $v = $wb.Worksheets.Item(1).UsedRange.Value2
$wb.Close($false); $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
$num = { param($x) $d = 0.0; if ($null -ne $x -and [double]::TryParse("$x", [ref]$d)) { $d } else { $null } }
$types = @{}
for ($r = 3; $r -le 13; $r++) { $p = "$($v[$r,1])"; if ($p) { $types[$p] = @{ min = & $num $v[$r,5]; max = & $num $v[$r,6] } } }
$rows = @()
for ($r = 15; $r -le $v.GetLength(0); $r++) {
    $sku = "$($v[$r,2])".Trim(); if (-not $sku) { continue }
    $pt = "$($v[$r,3])"; $min = & $num $v[$r,11]; $max = & $num $v[$r,12]
    if ($null -eq $min -and $types[$pt]) { $min = $types[$pt].min; $max = $types[$pt].max }
    if ($null -eq $min -or $null -eq $max) { continue }
    $fee = & $num $v[$r,10]; $esp = & $num $v[$r,9]
    if ($null -ne $esp -and $esp -le 0) { $fee = $null }   # no settlement known for this row
    $rows += [pscustomobject]@{ acc = "$($v[$r,1])".Trim(); sku = $sku; type = $pt; min = $min; max = $max; fee = $fee }
}
# targets.json (dashboard repo, kept current) wins over the sheet's Min/Max; its listings missing from the sheet
# are added too (their fees come from fees.json below, else they are skipped)
try {
    $tj = Get-Content (Join-Path $root 'targets.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($ap in $tj.PSObject.Properties) {
        foreach ($sp in $ap.Value.PSObject.Properties) {
            $x = $rows | Where-Object { $_.acc -eq $ap.Name -and $_.sku -eq $sp.Name } | Select-Object -First 1
            if ($x) { $x.min = [double]$sp.Value[0]; $x.max = [double]$sp.Value[1] }
            else { $rows += [pscustomobject]@{ acc = $ap.Name; sku = $sp.Name; type = ''; min = [double]$sp.Value[0]; max = [double]$sp.Value[1]; fee = $null } }
        }
    }
} catch { Log "targets.json load FAIL $($_.Exception.Message)" }
foreach ($g in $rows | Where-Object type | Group-Object type) {
    $f = @($g.Group | Where-Object { $null -ne $_.fee } | ForEach-Object fee | Sort-Object)
    if ($f.Count) { $med = $f[[int][Math]::Floor($f.Count / 2)]; $g.Group | Where-Object { $null -eq $_.fee } | ForEach-Object { $_.fee = $med } }
}
$feesAt = $null
if ($Publish -and (Test-Path (Join-Path $root 'fees.json'))) {
    try {
        $fb = Get-Content (Join-Path $root 'fees.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        $feesAt = $fb.at; $n = 0
        foreach ($x in $rows) { $f = $fb.fees.($x.acc.ToUpper()); if ($f) { $v2 = $f.($x.sku); if ($null -ne $v2) { $x.fee = [double]$v2; $n++ } } }
        Log "Seller Hub fees from $feesAt applied to $n listings"
    } catch { Log "fees load FAIL $($_.Exception.Message)" }
}
Log "---- run $(if($DryRun){'DRY'}else{'LIVE'}) : $($rows.Count) listings with targets"

$report = @(); $stat = [ordered]@{}
foreach ($a in Import-Csv (Join-Path $root 'accounts.csv')) {
    $mine = @($rows | Where-Object { $a.Name -match "_$([regex]::Escape($_.acc))$" -and $null -ne $_.fee -and (-not $Only -or $_.sku -eq $Only) })
    if (-not $mine.Count) { continue }
    try {
        $b64 = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($a.AppId):$($a.AppSecret)"))
        $tok = (Invoke-RestMethod -Headers @{ Authorization = "Basic $b64" } -Uri "$api/oauth-service/oauth/token?grant_type=client_credentials&scope=Seller_Api").access_token
    } catch { Log "$($a.Name) token FAIL"; continue }
    $h = @{ Authorization = "Bearer $tok" }
    for ($i = 0; $i -lt $mine.Count; $i += 10) {
        $bt = $mine[$i..([Math]::Min($i + 9, $mine.Count - 1))]
        $res = $null
        try { $res = Invoke-RestMethod -Headers $h -Uri "$api/sellers/listings/v3/$(($bt | ForEach-Object { [Uri]::EscapeDataString($_.sku) }) -join ',')" }
        catch {   # one bad SKU fails the whole batch - read them one by one
            $res = [pscustomobject]@{ available = [pscustomobject]@{} }
            foreach ($x in $bt) { try { $one = Invoke-RestMethod -Headers $h -Uri "$api/sellers/listings/v3/$([Uri]::EscapeDataString($x.sku))"; $p = $one.available.($x.sku); if ($p) { $res.available | Add-Member -NotePropertyName $x.sku -NotePropertyValue $p -Force } } catch { Log "$($a.Name) $($x.sku) read FAIL" } }
        }
        $upd = @{}
        foreach ($x in $bt) {
            $L = $res.available.($x.sku); if (-not $L) { continue }
            $sp = [double]$L.price.selling_price; $mrp = [double]$L.price.mrp
            $esp = $sp - $x.fee; $new = $null
            $ak = $x.acc.ToUpper(); if (-not $stat[$ak]) { $stat[$ak] = [ordered]@{ total = 0; inR = 0; low = 0; high = 0; noT = 0; zero = 0; lows = @(); highs = @() } }
            $s = $stat[$ak]; $s.total++
            if ($esp -lt $x.min - $Tol) { $s.low++; $s.lows += "$($x.sku) $([Math]::Round($esp))/$($x.min)" }
            elseif ($esp -gt $x.max + $Tol) { $s.high++; $s.highs += "$($x.sku) $([Math]::Round($esp))/$($x.max)" }
            else { $s.inR++ }
            if ($esp -lt $x.min - $Tol) {
                $want = [Math]::Ceiling(($x.min + $x.max) / 2)
                $cap = [Math]::Floor($esp * 1.05); if ($esp -gt 0 -and $want -gt $cap) { $want = $cap }   # Flipkart: max +5% per change
                $new = [Math]::Round($want + $x.fee)
            } elseif ($Lower -and $esp -gt $x.max + $Tol) {
                $new = [Math]::Round([Math]::Floor(($x.min + $x.max) / 2) + $x.fee)
            }
            if ($null -eq $new -or $new -eq $sp) { continue }
            if ($new -gt $mrp) { Log "$($a.Name) $($x.sku) skip: new price $new > MRP $mrp"; continue }
            $report += [pscustomobject]@{ acc = $a.Name; sku = $x.sku; type = $x.type; price = $sp; settle = $esp; min = $x.min; max = $x.max; newPrice = $new; newSettle = $new - $x.fee }
            $upd[$x.sku] = @{ product_id = $L.product_id; price = @{ mrp = $mrp; selling_price = $new; currency = 'INR' } }
        }
        if ($upd.Count -and -not $DryRun) {
            try {
                $r = Invoke-RestMethod -Method Post -Headers $h -ContentType 'application/json' -Uri "$api/sellers/listings/v3/update/price" -Body ($upd | ConvertTo-Json -Depth 5)
                foreach ($k in $upd.Keys) { $st = $r.$k.status; Log "$($a.Name) $k -> $($upd[$k].price.selling_price) : $st $(if($st -ne 'SUCCESS'){$r.$k | ConvertTo-Json -Compress -Depth 5})" }
            } catch { Log "$($a.Name) update FAIL $($_.Exception.Message)" }
        }
    }
}
$report | Export-Csv (Join-Path $root 'settlement-keeper-report.csv') -NoTypeInformation -Encoding UTF8
Log "changes: $($report.Count)"
# Seller Hub status (real settlement, written when the keeper runs in Seller Hub) is better than this
# price - fees estimate: do not overwrite it while it is less than 20 hours old
$keepHub = $false
try { $old = Get-Content (Join-Path $root 'price.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($old.source -eq 'sellerhub' -and ((Get-Date).ToUniversalTime() - ([datetime]$old.at).ToUniversalTime()).TotalHours -lt 20) { $keepHub = $true; Log "price.json from Seller Hub ($($old.at)) is newer, not overwritten" }
} catch { }
if ($Publish -and $stat.Count -and -not $keepHub) {
    # Price tab data = price.json in the dashboard repo (GitHub Pages); commit + push it
    try {
        $data = @{ at = (Get-Date).ToUniversalTime().ToString('o'); source = 'api'; feesAt = $feesAt; accounts = $stat } | ConvertTo-Json -Depth 6 -Compress
        [IO.File]::WriteAllText((Join-Path $root 'price.json'), $data, (New-Object System.Text.UTF8Encoding $false))
        Push-Location $root; git add price.json; git commit -q -m "price status update" 2>&1 | Out-Null; git push -q 2>&1 | Out-Null; Pop-Location
        Log "Price tab status published (price.json pushed)"
    } catch { Log "publish FAIL $($_.Exception.Message)" }
}



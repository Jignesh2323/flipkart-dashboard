# Flipkart multi-account data fetcher
# Reads accounts.csv, pulls orders (shipments) + returns via Flipkart Seller API,
# writes data.js which dashboard.html reads.
param([int]$Days = 30)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$api  = 'https://api.flipkart.net'
$rawDir = Join-Path $root 'raw'
New-Item -ItemType Directory -Force $rawDir | Out-Null

function Get-Token($appId, $secret) {
    $b64 = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${appId}:${secret}"))
    $r = Invoke-RestMethod -Method Get -Headers @{ Authorization = "Basic $b64" } `
        -Uri "$api/oauth-service/oauth/token?grant_type=client_credentials&scope=Seller_Api"
    $r.access_token
}

# Retries flaky calls (dropped connections, 5xx) a few times before giving up
function Invoke-Retry([scriptblock]$call) {
    for ($i = 1; ; $i++) {
        try { return & $call }
        catch {
            $code = 0; if ($_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode }
            if ($i -ge 4 -or ($code -ge 400 -and $code -lt 500 -and $code -ne 429)) { throw }
            Start-Sleep -Seconds (2 * $i)
        }
    }
}

# $prefix: path the API's relative next-page links hang off (they omit it)
function Invoke-Paged($token, $firstCall, $listProp, $prefix) {
    $h = @{ Authorization = "Bearer $token"; 'Content-Type' = 'application/json' }
    $all = @()
    $r = Invoke-Retry { & $firstCall $h }
    while ($true) {
        if ($r.$listProp) { $all += $r.$listProp }
        $next = $r.nextPageUrl; if (-not $next) { $next = $r.nextUrl }
        if (-not $r.hasMore -or -not $next) { break }
        if ($next -notmatch '^https?://') { $next = "$api$prefix$next" }
        $r = Invoke-Retry { Invoke-RestMethod -Method Get -Headers $h -Uri $next }
    }
    $all
}

function Get-Shipments($token, $type, $states, $from, $to) {
    $dateField = if ($type -eq 'cancelled') { 'cancellationDate' } else { 'orderDate' }
    $body = @{
        filter = @{ type = $type; states = $states; $dateField = @{ from = $from; to = $to } }
        pagination = @{ pageSize = 20 }
    } | ConvertTo-Json -Depth 6
    Invoke-Paged $token { param($h) Invoke-RestMethod -Method Post -Headers $h -Body $body -Uri "$api/sellers/v3/shipments/filter" } 'shipments' '/sellers'
}

function Get-Returns($token, $source, $from, $to) {
    Invoke-Paged $token { param($h) Invoke-RestMethod -Method Get -Headers $h `
        -Uri "$api/sellers/v2/returns?source=$source&createdAfter=$($from.Substring(0,10))&createdBefore=$($to.Substring(0,10))" } 'returnItems' '/sellers/v2'
}

$accounts = Import-Csv (Join-Path $root 'accounts.csv') | Where-Object { $_.AppId -and $_.AppId -notmatch 'PASTE' }
if (-not $accounts) { Write-Host 'accounts.csv me koi AppId nahi mila.' -ForegroundColor Red; exit 1 }

$to   = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss.000+05:30')
$from = (Get-Date).AddDays(-$Days).ToString('yyyy-MM-ddTHH:mm:ss.000+05:30')

$orders = @(); $returns = @(); $status = @()
foreach ($a in $accounts) {
    Write-Host "== $($a.Name) ==" -ForegroundColor Cyan
    try {
        $t = Get-Token $a.AppId $a.AppSecret
        $groups = @(
            @{ type = 'preDispatch';  states = @('APPROVED','PACKING_IN_PROGRESS','PACKED','FORM_FAILED','READY_TO_DISPATCH') },
            @{ type = 'postDispatch'; states = @('SHIPPED','DELIVERED') },
            @{ type = 'cancelled';    states = @('CANCELLED') }
        )
        $ships = @()
        foreach ($g in $groups) {
            try { $ships += Get-Shipments $t $g.type $g.states $from $to }
            catch { Write-Host "  $($g.type): $($_.Exception.Message)" -ForegroundColor Yellow }
        }
        $ships | ConvertTo-Json -Depth 12 | Out-File (Join-Path $rawDir "$($a.Name)_shipments.json") -Encoding utf8
        foreach ($s in $ships) {
            foreach ($i in $s.orderItems) {
                $p = $i.priceComponents
                $orders += [pscustomobject]@{
                    acc = $a.Name; orderId = $i.orderId; itemId = $i.orderItemId
                    date = "$($i.orderDate)".Substring(0, [Math]::Min(10, "$($i.orderDate)".Length))
                    sku = $i.sku; title = $i.title; qty = [int]$i.quantity
                    amount = [double]$(if ($p.totalPrice) { $p.totalPrice } else { $p.sellingPrice })
                    status = $i.status; cancelReason = $i.cancellationReason
                }
            }
        }
        foreach ($src in 'customer_return','courier_return') {
            try {
                $rs = Get-Returns $t $src $from $to
                $rs | ConvertTo-Json -Depth 12 | Out-File (Join-Path $rawDir "$($a.Name)_$src.json") -Encoding utf8
                foreach ($r in $rs) {
                    $returns += [pscustomobject]@{
                        acc = $a.Name; returnId = $r.returnId; itemId = $r.orderItemId
                        date = "$($r.createdDate)".Substring(0, [Math]::Min(10, "$($r.createdDate)".Length))
                        type = $src; reason = $r.reason; subReason = $r.subReason
                        qty = [int]$(if ($r.quantity) { $r.quantity } else { 1 }); sku = $r.sku; status = $r.status
                    }
                }
            } catch { Write-Host "  returns $src : $($_.Exception.Message)" -ForegroundColor Yellow }
        }
        # Returns API has no SKU; returns of orders older than the window need a lookup by item id
        $skuOf = @{}; $orders | Where-Object { $_.acc -eq $a.Name } | ForEach-Object { $skuOf[$_.itemId] = $_.sku }
        $need = @($returns | Where-Object { $_.acc -eq $a.Name -and -not $_.sku -and -not $skuOf[$_.itemId] } | ForEach-Object { $_.itemId } | Select-Object -Unique)
        for ($i = 0; $i -lt $need.Count; $i += 25) {
            $ids = ($need[$i..([Math]::Min($i + 24, $need.Count - 1))]) -join ','
            try {
                $r = Invoke-Retry { Invoke-RestMethod -Headers @{ Authorization = "Bearer $t" } -Uri "$api/sellers/v3/shipments?orderItemIds=$ids" }
                foreach ($s in $r.shipments) { foreach ($i2 in $s.orderItems) { $skuOf[$i2.orderItemId] = $i2.sku } }
            } catch { Write-Host "  sku lookup: $($_.Exception.Message)" -ForegroundColor Yellow }
        }
        $returns | Where-Object { $_.acc -eq $a.Name -and -not $_.sku } | ForEach-Object { $_.sku = $skuOf[$_.itemId] }
        $status += [pscustomobject]@{ acc = $a.Name; ok = $true; msg = "$($ships.Count) shipments" }
        Write-Host "  OK: $($ships.Count) shipments" -ForegroundColor Green
    } catch {
        $status += [pscustomobject]@{ acc = $a.Name; ok = $false; msg = $_.Exception.Message }
        Write-Host "  FAIL: $($_.Exception.Message)" -ForegroundColor Red
    }
}

$orders  = $orders  | Sort-Object acc, itemId -Unique
$returns = $returns | Sort-Object acc, returnId -Unique
$data = [pscustomobject]@{
    generated = (Get-Date).ToString('yyyy-MM-dd HH:mm'); days = $Days
    orders = @($orders); returns = @($returns); status = @($status)
}
"window.FK_DATA = " + ($data | ConvertTo-Json -Depth 6 -Compress) + ";" |
    Out-File (Join-Path $root 'data.js') -Encoding utf8
Write-Host "`nDone. dashboard.html kholiye." -ForegroundColor Green

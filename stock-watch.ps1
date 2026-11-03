# Stock watch: reads every listing in targets.json (all 9 accounts), reads its stock via the official Listings API,
# and shows a Windows notification when a listing goes below $Min units (or out of stock).
# Only NEW low-stock listings are notified (state kept in stock-watch-state.json); the full current list is in stock-low.txt.
# Scheduled task "Flipkart Stock Watch" runs it hourly via stock-watch-hidden.vbs.
param([int]$Min = 10)
$ErrorActionPreference = 'Continue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$api  = 'https://api.flipkart.net'
$tj = Get-Content (Join-Path $root 'targets.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$low = [ordered]@{}
foreach ($a in Import-Csv (Join-Path $root 'accounts.csv')) {
    $acc = ($a.Name -split '_', 2)[1].ToUpper()
    $skus = @($tj.$acc.PSObject.Properties.Name); if (-not $skus.Count) { continue }
    try {
        $b64 = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($a.AppId):$($a.AppSecret)"))
        $tok = (Invoke-RestMethod -Headers @{ Authorization = "Basic $b64" } -Uri "$api/oauth-service/oauth/token?grant_type=client_credentials&scope=Seller_Api").access_token
    } catch { continue }
    $h = @{ Authorization = "Bearer $tok" }
    for ($i = 0; $i -lt $skus.Count; $i += 10) {
        $bt = $skus[$i..([Math]::Min($i + 9, $skus.Count - 1))]
        $res = $null
        try { $res = (Invoke-RestMethod -Headers $h -Uri "$api/sellers/listings/v3/$(($bt | ForEach-Object { [Uri]::EscapeDataString($_) }) -join ',')").available }
        catch { $res = [pscustomobject]@{}; foreach ($s in $bt) { try { $one = (Invoke-RestMethod -Headers $h -Uri "$api/sellers/listings/v3/$([Uri]::EscapeDataString($s))").available.$s; if ($one) { $res | Add-Member -NotePropertyName $s -NotePropertyValue $one -Force } } catch { } } }
        foreach ($s in $bt) {
            $L = $res.$s; if (-not $L -or $L.listing_status -ne 'ACTIVE') { continue }   # archived/inactive listings are skipped
            $q = ($L.locations | Measure-Object inventory -Sum).Sum
            if ($q -lt $Min) { $low["$acc | $s"] = [int]$q }
        }
    }
}
$lines = $low.Keys | ForEach-Object { "$_ = $($low[$_])" }
Set-Content (Join-Path $root 'stock-low.txt') (@("Checked $(Get-Date -f 'dd MMM HH:mm') - stock below $Min") + $lines) -Encoding UTF8
# dashboard Stock tab: stock.json in the repo (GitHub Pages), pushed only when the list changed
$items = @($low.Keys | ForEach-Object { $p = $_ -split ' \| ', 2; @{ acc = $p[0]; sku = $p[1]; q = $low[$_] } })
$sj = Join-Path $root 'stock.json'
$prev = if (Test-Path $sj) { try { (Get-Content $sj -Raw | ConvertFrom-Json).items | ConvertTo-Json -Compress } catch { '' } } else { '' }
$data = @{ at = (Get-Date).ToUniversalTime().ToString('o'); items = $items }
[IO.File]::WriteAllText($sj, ($data | ConvertTo-Json -Depth 4 -Compress), (New-Object System.Text.UTF8Encoding $false))
if (($items | ConvertTo-Json -Compress) -ne $prev) {
    Push-Location $root; git add stock.json; git commit -q -m "stock status update" 2>&1 | Out-Null; git push -q 2>&1 | Out-Null; Pop-Location
}
$stf = Join-Path $root 'stock-watch-state.json'
$old = @{}; if (Test-Path $stf) { (Get-Content $stf -Raw | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $old[$_.Name] = $_.Value } }
$new = @($low.Keys | Where-Object { -not $old.ContainsKey($_) -or ($low[$_] -eq 0 -and $old[$_] -ne 0) })
$low | ConvertTo-Json | Set-Content $stf -Encoding UTF8
if ($new.Count) {
    $oos = @($new | Where-Object { $low[$_] -eq 0 }).Count
    $title = "Flipkart stock: $($new.Count) listing low ($oos out of stock)"
    $body = (($new | Select-Object -First 6 | ForEach-Object { "$_ = $($low[$_])" }) -join "`n")
    [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
    $x = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)
    $t = $x.GetElementsByTagName('text'); $t.Item(0).InnerText = $title; $t.Item(1).InnerText = $body
    $app = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
    [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($app).Show([Windows.UI.Notifications.ToastNotification]::new($x))
}

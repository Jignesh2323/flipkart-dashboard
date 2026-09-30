# Bundles settlements\*.json (collected from Seller Hub) into settle.js for dashboard.html.
# settlements\daily\<acc>.json (per payment date totals) is attached to its account as "dd".
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$parts = Get-ChildItem (Join-Path $root 'settlements') -Filter *.json | ForEach-Object {
    $body = (Get-Content $_.FullName -Raw).Trim()
    $daily = Join-Path $root "settlements\daily\$($_.Name)"
    if (Test-Path $daily) { $body = $body.Substring(0, $body.LastIndexOf('}')) + ',"dd":' + (Get-Content $daily -Raw).Trim() + '}' }
    '"' + $_.BaseName + '":' + $body
}
$stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm')
"window.FK_SETTLE = {" + ($parts -join ',') + "}; window.FK_SETTLE_AT = '$stamp';" |
    Out-File (Join-Path $root 'settle.js') -Encoding utf8
Write-Host "settle.js: $($parts.Count) accounts"

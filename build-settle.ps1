# Bundles settlements\*.json (collected from Seller Hub) into settle.js for dashboard.html
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$parts = Get-ChildItem (Join-Path $root 'settlements') -Filter *.json | ForEach-Object {
    '"' + $_.BaseName + '":' + (Get-Content $_.FullName -Raw).Trim()
}
$stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm')
"window.FK_SETTLE = {" + ($parts -join ',') + "}; window.FK_SETTLE_AT = '$stamp';" |
    Out-File (Join-Path $root 'settle.js') -Encoding utf8
Write-Host "settle.js: $($parts.Count) accounts"

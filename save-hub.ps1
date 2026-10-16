# Saves a Seller Hub keeper run for the dashboard: takes the JSON from __dump() in the Seller Hub tab
# (saved to a file), writes price.json (Price tab counts) and fees.json (per-listing fee = price - settlement,
# used by the 9:30 morning check), then commits and pushes both.
# Usage: save-hub.ps1 <path-to-dump.json>
param([Parameter(Mandatory)][string]$Dump)
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$d = Get-Content $Dump -Raw -Encoding UTF8 | ConvertFrom-Json
$utf8 = New-Object System.Text.UTF8Encoding $false
[IO.File]::WriteAllText((Join-Path $root 'price.json'), (@{ at = $d.at; source = 'sellerhub'; accounts = $d.accounts } | ConvertTo-Json -Depth 6 -Compress), $utf8)
[IO.File]::WriteAllText((Join-Path $root 'fees.json'), (@{ at = $d.at; fees = $d.fees } | ConvertTo-Json -Depth 4 -Compress), $utf8)
Push-Location $root
git add price.json fees.json; git commit -q -m "settlement status + fees from Seller Hub" 2>&1 | Out-Null; git push -q 2>&1 | Out-Null
Pop-Location
Write-Host "price.json + fees.json saved and pushed ($($d.at))"

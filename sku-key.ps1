# Flipkart SKU -> product name shared by all accounts: drop the account prefix (KDC-, TD-, ZT-, ...)
# and the trailing listing number (-01, -10, ...), which only counts repeat listings.
function Get-ProductName([string]$sku) {
    $s = $sku.Trim()
    $s = $s -replace '^(ARZ|ARKZ|AVD|AVSP|AV|KDCSP|KDC|PVDSP|PVD|TD|WSKSP|WSK|ZTSP|ZT|ZO|FO|VU|BRL|BR)\s*-\s*(SP-)?', ''
    $s = $s -replace '[\s_-]+0*\d{1,2}[a-zA-Z]?$', ''
    $s = ($s -replace '\s+', ' ').Trim()
    return $s.ToUpper()
}

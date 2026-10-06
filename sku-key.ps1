# Flipkart SKU -> product name shared by all accounts: drop the account prefix (KDC-, TD-, ZT-, ...)
# and the trailing listing number (-01, -10, ...), which only counts repeat listings.
function Get-ProductName([string]$sku) {
    $s = $sku.Trim()
    $s = $s -replace '^(ARZSP|AEZSP|ARZ|ARKZ|AVD|AVSP|AV|KDCSP|KDC|PVDSP|PVD|TDSP|TD|WSKSP|WSK|ZTSP|ZT|ZO|FOV|FO|VUSP|VU|BRSP|BRL|BRE|BR)\s*-\s*(SP-)?', ''
    $s = $s -replace '[\s_-]+0*\d{1,2}[a-zA-Z]?$', ''
    $s = ($s -replace '\s+', ' ').Trim()
    return $s.ToUpper()
}

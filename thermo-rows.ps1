# Builds + fills the kitchen_thermometer bulk template for one account: TP49 mini hygrometer and the digital
# hygrometer clock (copied from TAUDI listings TD-ThermoPro-TP49 / TD-Digital Hygrometer), Flipkart + Shopsy SKU.
# Usage: thermo-rows.ps1 -Template <downloaded xls, kept under its own name> -Acc WISKA
param([Parameter(Mandatory)][string]$Template, [Parameter(Mandatory)][string]$Acc)
$cfg = @{ AVEDIA = 'AV', 'Avedia'; ZOTIMO = 'ZT', 'ZOTIMO'; PIVDO = 'PVD', 'Pivdo'; WISKA = 'WSK', 'Wiska'; ARKZO = 'ARZ', 'ARKZO'
    BRAINLE = 'BR', 'brainle'; KIDOCA = 'KDC', 'KIDOCA'; VUSTICA = 'VU', 'VUSTICA' }[$Acc]
$pfx, $brand = $cfg
$img = 'https://jignesh2323.github.io/flipkart-dashboard/listing-images/thermo/'
# title-like Model Number: different wording per account and channel (Flipkart builds the title from it)
$tpTitles = @{
    AV = 'Mini Room Thermometer Hygrometer Digital Temperature Humidity Meter With Comfort Indicator'
    ZT = 'Digital Hygrometer Room Thermometer Indoor Humidity Monitor Mini LCD Meter'
    PVD = 'Mini Digital Indoor Thermometer Hygrometer Room Temperature And Humidity Gauge'
    WSK = 'Mini Room Temperature Humidity Monitor Digital Hygrometer Thermometer Meter'
    ARZ = 'Indoor Digital Thermometer Hygrometer Mini Humidity Meter For Home Office Baby Room'
    BR = 'Room Thermometer And Hygrometer Mini Digital Humidity Temperature Monitor'
    KDC = 'Mini Hygrometer Thermometer Digital Room Humidity Meter For Baby Room Bedroom'
    VU = 'Digital Room Hygrometer Thermometer Mini Temperature Humidity Sensor Meter' }
$hgTitles = @{
    AV = 'Digital Room Thermometer Hygrometer Clock Large LCD Temperature Humidity Monitor With Backlight'
    ZT = 'Indoor Hygrometer Thermometer Clock Digital Humidity Meter With Max Min Record Backlight'
    PVD = 'Digital Hygrometer Room Thermometer With Clock Date Backlit LCD Humidity Monitor'
    WSK = 'Digital Room Temperature And Humidity Monitor Hygrometer Clock Meter With Backlight'
    ARZ = 'Room Temperature Humidity Meter Digital Hygrometer Clock With Date Max Min Backlight'
    BR = 'Large Display Digital Thermometer Hygrometer Clock Indoor Humidity Monitor Backlit'
    KDC = 'Digital Hygrometer Thermometer Clock For Home Baby Room Humidity Meter With Backlight'
    VU = 'Backlit Digital Room Thermometer Hygrometer With Clock Date Humidity Temperature Monitor' }
$common = @{ 9 = 'Active'; 10 = 1999; 12 = 'Seller'; 13 = 'express'; 14 = 1; 15 = 100; 16 = 'Flipkart'; 17 = 0; 18 = 0; 19 = 0
    20 = 11.5; 21 = 4; 22 = 11.5; 23 = 0.1; 24 = '9025'; 26 = 'India'; 27 = "$($brand.ToUpper()) INDIA"; 28 = "$($brand.ToUpper()) INDIA"
    30 = 'GST_18'; 31 = 1; 32 = $brand; 34 = 'Instant Read Thermocouple'; 35 = 'Plastic'; 36 = 'Digital'; 38 = 'Yes'; 39 = 'AAA'
    42 = 'cm'; 44 = 'cm'; 61 = 6; 62 = 'Months'; 70 = 1 }
$tp = @{ 11 = 799; 37 = '-50°C to 70°C'; 40 = 1; 41 = 5.5; 43 = 6.5; 50 = '1 x Digital Thermometer Hygrometer, 1 x AAA Battery'
    51 = 'Mini digital room thermometer and hygrometer that shows indoor temperature and humidity on a large, clear LCD. Accuracy of +-1 degree C and +-2-3% RH with a fresh reading every 10 seconds. The face icon shows DRY, COMFORT or WET at a glance so you can keep a healthy room for your family. Three ways to place it: tabletop stand, strong back magnet (fridge) or wall hanging hole. AAA battery included, lasts 18-24 months. Ideal for bedroom, baby room, living room, office, greenhouse, cellar and incubator.'
    52 = 'room thermometer::hygrometer::humidity meter::digital thermometer for room::temperature and humidity monitor::indoor thermometer::baby room thermometer::mini hygrometer'
    57 = 'High Accuracy and 10s Refresh::Comfort Level Indicator::Large Clear LCD Digits::Stand, Magnet and Wall Mount::AAA Battery Included'
    45 = "${img}tp49-1.jpg"; 46 = "${img}tp49-2.jpg"; 47 = "${img}tp49-3.jpg"; 48 = "${img}tp49-4.jpg"; 49 = "${img}tp49-5.jpg" }
$hg = @{ 11 = 1000; 37 = '-10°C to 50°C'; 40 = 2; 41 = 9; 43 = 8.5; 50 = '1 x Digital Hygrometer Clock'
    51 = 'Digital indoor thermometer and hygrometer with clock. The large 3.34 inch backlit LCD shows room temperature, humidity, MAX/MIN records, time (12/24H) and date together. The colour comfort bar (DRY / COMFORT / WET) helps you keep the right air condition at home. Fast 10 second refresh, calibration setting, and 3 ways to place it: table stand, wall hanging or magnetic back. Runs on 2 AAA batteries. Ideal for bedroom, baby nursery, living room, office and greenhouse.'
    52 = 'hygrometer clock::room thermometer with clock::digital hygrometer::humidity meter::temperature and humidity monitor::indoor thermometer::backlight thermometer::max min thermometer'
    57 = 'Temperature, Humidity, Time and Date::MAX / MIN Records::Backlit Large LCD::Comfort Level Bar::Stand, Wall or Magnet Mount'
    45 = "${img}hyg-1.jpg"; 46 = "${img}hyg-2.jpg"; 47 = "${img}hyg-3.jpg"; 48 = "${img}hyg-4.jpg"; 49 = "${img}hyg-5.jpg" }
$rows = @(
    @{ 7 = "$pfx-ThermoPro-TP49"; 33 = $tpTitles[$pfx] } + $tp
    @{ 7 = "${pfx}SP-ThermoPro-TP49"; 33 = "Premium $($tpTitles[$pfx])" } + $tp
    @{ 7 = "$pfx-Digital Hygrometer"; 33 = $hgTitles[$pfx] } + $hg
    @{ 7 = "${pfx}SP-Digital Hygrometer"; 33 = "Premium $($hgTitles[$pfx])" } + $hg )
$x = New-Object -ComObject Excel.Application; $x.Visible = $false; $x.DisplayAlerts = $false
try {
    $wb = $x.Workbooks.Open($Template); $ws = $wb.Worksheets.Item('kitchen_thermometer')
    $r = 5
    foreach ($row in $rows) { foreach ($k in $common.Keys) { $ws.Cells($r, $k).Value2 = "$($common[$k])" }
        foreach ($k in $row.Keys) { $v = "$($row[$k])"; if ($k -eq 33) { while ($v.Length -gt 80) { $v = $v.Substring(0, $v.LastIndexOf(' ')) } }; $ws.Cells($r, $k).Value2 = $v }; "row $r $($row[7])"; $r++ }
    $wb.Save(); $wb.Close($false)
} finally { $x.Quit() }

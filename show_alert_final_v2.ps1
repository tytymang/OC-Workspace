Add-Type -AssemblyName System.Windows.Forms
$m1 = -join @([char]51452, [char]51064, [char]45784, " 8:30", [char]50640, " ", [char]51068, [char]51221, [char]51060, " ", [char]51080, [char]49845, [char]45768, [char]45796, ".")
$m2 = -join @("`n`n", [char]52572, [char]44540, " ", [char]51473, [char]50836, " ", [char]47700, [char]51068, ":")
$m3 = -join @("`n1. ", [char]51060, [char]50696, [char]51652, " ", [char]51452, [char]51076, ": ", [char]51076, [char]50896, " KPI ", [char]54217, [char]44032)
$m4 = -join @("`n2. ", [char]49552, [char]48124, [char]54712, " ", [char]49345, [char]44036, ": G2(8GR) ", [char]54924, [char]51032, " ", [char]45236, [char]50857)
$msg = $m1 + $m2 + $m3 + $m4
$title = -join @([char]46028, [char]49604, " ", [char]49692, [char]52272, " ", [char]50508, [char]47548)
[System.Windows.Forms.MessageBox]::Show($msg, $title)

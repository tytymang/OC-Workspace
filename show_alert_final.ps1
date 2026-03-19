$msg = -join @(
    [char]51452, [char]51064, [char]45784, " ", # 주 인 님
    "8:30", [char]50640, " ", # 에
    [char]51068, [char]51221, [char]51060, " ", [char]51080, [char]49845, [char]45768, [char]45796, ".", "`n`n", # 일 정 이  있 습 니 다 .
    [char]52572, [char]44540, " ", [char]51473, [char]50836, " ", [char]47700, [char]51068, ":", "`n", # 최 근  중 요  메 일 :
    "1. ", [char]51060, [char]50696, [char]51652, " ", [char]51452, [char]51076, ": ", [char]51076, [char]50896, " KPI ", [char]54217, [char]44032, "`n", # 1. 이 예 진  주 임 : 임 원  K P I  평 가
    "2. ", [char]49552, [char]48124, [char]54712, " ", [char]49345, [char]44036, ": G2(8GR) ", [char]54924, [char]51032, " ", [char]45236, [char]50857 # 2. 손 민 혁  상 무 : G 2 ( 8 G R )  회 의  내 용
)

$title = -join @([char]46028, [char]49604, " ", [char]49692, [char]52272, " ", [char]50508, [char]47548) # 돌 쇠  순 찰  알 림

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show($msg, $title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

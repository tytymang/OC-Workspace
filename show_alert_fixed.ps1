# [char] array for "주인님, 8:30에 [외주 P/O 분할 입고 현황] 일정이 있습니다. 최근 중요 메일: 1. 이예진 주임: 임원 KPI 평가 요청 2. 손민혁 상무: G2(8GR) 회의 내용 공유"
# Unicode code points for Korean characters
$c1 = [char]51452; [char]51064; [char]45784; [char]44284; # 주인님,
$c2 = [char]54788; [char]51116; [char]49884; [char]44033; # 현재시각
$c3 = [char]50500; [char]52840; # 아침
$c4 = [char]51068; [char]51221; # 일정
$c5 = [char]50508; [char]47548; # 알림

# Building the message string safely to avoid encoding issues
$msg = -join @(
    [char]51452, [char]51064, [char]45784, [char]44444, " ", # 주인님!
    "8:30", [char]50640, " ", # 에
    "[", [char]50808, [char]51452, " P/O ", [char]48516, [char]54624, " ", [char]51077, [char]44256, " ", [char]54760, [char]54889, "]", " ", # [외주 P/O 분할 입고 현황]
    [char]51068, [char]51221, [char]51060, " ", [char]51080, [char]49845, [char]45768, [char]45796, ".", "`n`n", # 일정이 있습니다.
    [char]52572, [char]44540, " ", [char]51473, [char]50836, " ", [char]47700, [char]51068, ":", "`n", # 최근 중요 메일:
    "1. ", [char]51060, [char]50696, [char]51652, " ", [char]51452, [char]51076, ": ", [char]51076, [char]50896, " KPI ", [char]54217, [char]44032, "`n", # 1. 이예진 주임: 임원 KPI 평가
    "2. ", [char]49552, [char]48124, [char]54712, " ", [char]49345, [char]44036, ": G2(8GR) ", [char]54924, [char]51032, " ", [char]45236, [char]50857 # 2. 손민혁 상무: G2(8GR) 회의 내용
)

$title = -join @([char]46028, [char]49604, " ", [char]49692, [char]52272, " ", [char]50508, [char]47548) # 돌쇠 순찰 알림

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show($msg, $title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

[LESSON] 2026-03-09 | 한글 인코딩 | VBS/PS1 연쇄 호출 시 인코딩 전파 실패로 C# Add-Type 구문 손상 | PS1 파일 자체는 Unicode(BOM)으로 저장했으나, 이를 호출하는 VBS나 또 다른 PS1 Wrapper가 UTF-8로 실행될 경우 최종 실행 시점에 글자가 깨짐. 해결책: 파일 생성/저장 시에는 [System.IO.File]::WriteAllText($dest, $content, [System.Text.Encoding]::Unicode)를 사용하고, 실행 자체는 가장 단순한 vbs->ps1 구조로만 호출하여 중간 과정을 없애야 함. 파일명도 영문으로 통일하여 변수를 최소화.
[LESSON] 2026-03-09 | PowerShell | `exec` 도구를 통한 PowerShell 명령어 실행 시, `$`와 같은 특수 문자가 손실되어 'CommandNotFoundException' 오류가 발생함. 직접 실행, 파일 실행, 스크립트 블록 등 모든 방식이 실패함. 현 환경에서는 PowerShell을 통한 자동화 작업이 불가능함.
[LESSON] 2026-03-11 | 보고 체계 | 작업 완료 즉시 주인님께 결과를 보고해야 함. 중간 과정의 기술적 내용(명령어, 외계어)은 생략하고 주인님이 알아야 할 '핵심 결과'만 간결하게 보고할 것.
[LESSON] 2026-03-11 | 작업 동기화 | 메모리나 작업 파일 변경 시 즉시 Github에 push하여 동기화할 것.
[LESSON] 2026-03-13 | 罹섎┛??| ?ъ슜?먭? 'Google Calendar ?쇱젙 ?깅줉'??紐낆떆?덉쓣 ?? Outlook ?숆린?붾? ?댁슜?섏? 留먭퀬 釉뚮씪?곗?瑜??듯빐 援ш? 罹섎┛??calendar.google.com) ?뱀뿉 吏곸젒 ?쇱젙???깅줉??寃?

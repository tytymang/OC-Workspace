[LESSON] 2026-03-09 | 한글 인코딩 | VBS/PS1 연쇄 호출 시 인코딩 전파 실패로 C# Add-Type 구문 손상 | PS1 파일 자체는 Unicode(BOM)으로 저장했으나, 이를 호출하는 VBS나 또 다른 PS1 Wrapper가 UTF-8로 실행될 경우 최종 실행 시점에 글자가 깨짐. 해결책: 파일 생성/저장 시에는 [System.IO.File]::WriteAllText($dest, $content, [System.Text.Encoding]::Unicode)를 사용하고, 실행 자체는 가장 단순한 vbs->ps1 구조로만 호출하여 중간 과정을 없애야 함. 파일명도 영문으로 통일하여 변수를 최소화.
[LESSON] 2026-03-09 | PowerShell | `exec` 도구를 통한 PowerShell 명령어 실행 시, `$`와 같은 특수 문자가 손실되어 'CommandNotFoundException' 오류가 발생함. 직접 실행, 파일 실행, 스크립트 블록 등 모든 방식이 실패함. 현 환경에서는 PowerShell을 통한 자동화 작업이 불가능함.
[LESSON] 2026-03-11 | 보고 체계 | 작업 완료 즉시 주인님께 결과를 보고해야 함. 중간 과정의 기술적 내용(명령어, 외계어)은 생략하고 주인님이 알아야 할 '핵심 결과'만 간결하게 보고할 것.
[LESSON] 2026-03-11 | 작업 동기화 | 메모리나 작업 파일 변경 시 즉시 Github에 push하여 동기화할 것.
[LESSON] 2026-03-13 | Google Calendar | 일정 등록 시 'Google Calendar 설정 파일'을 직접 사용하지 말고 Outlook 일정에 직접 등록 후 calendar.google.com 에 직접 설정하여 등록할 것.

# MEMORY.md - 돌쇠의 장기 기억 (Long-Term Log)

> ⚠️ 메인 세션에서만 읽고 쓸 것. 그룹/공유 세션에서는 절대 열지 말 것.
> 이 파일은 주인님과의 주요 대화에서 도출된 장기 업무 규칙을 영구 보존하는 공간.
> 돌쇠는 주기적으로 새로운 교훈을 이곳에 커밋해야 한다.

## 고정 업무 지침
- 보고 시 기술 용어(외계어) 배제, 핵심 결과만 간결하게
- 숫자/계산 포함 보고는 반드시 이중 검증 후
- 바로 붙여넣기 가능한 깔끔한 결과물 제공
- 서론/아부 없이 결론부터

## 고정 소통 지침
- 주인님께는 머슴체 사용
- 외부 문서/메일은 주인님 지시에 따른 격식 사용

## 기술 환경
- OpenClaw Gateway: localhost:18792
- PowerShell 한글: UTF-16 LE BOM 필수
- Outlook 인코딩: UTF-16 LE BOM PowerShell 스크립트
- Google Calendar: Outlook 경유 등록

## 축적된 교훈 (LESSON)
- [2026-03-09] VBS→PS1 인코딩 체인 문제 → 단순 구조 + 영문 파일명
- [2026-03-09] exec 도구에서 $ 특수문자 손실 → 파일 기반 실행
- [2026-03-11] 보고는 즉시 + 핵심만 + 외계어 금지
- [2026-03-11] 변경 즉시 GitHub push
- [2026-03-13] Google Calendar 직접 등록 금지 → Outlook 경유

## 진행 중인 프로젝트
(작업 발생 시 기록)

## 주인님 업무 패턴
(시간이 지나면서 파악한 내용 축적)
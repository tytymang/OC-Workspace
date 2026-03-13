---
name: korean-encoding
description: PowerShell, Outlook, 파일 저장 등에서 한글 인코딩 문제를 방지하는 스킬. 한글이 포함된 모든 작업에서 자동으로 트리거된다. UTF-8, UTF-16, Unicode BOM, 코드포인트 변환을 처리한다.
---

# Korean Encoding (한글 인코딩 스킬)

## 언제 사용하는가
아래 조건 중 하나라도 해당되면 이 스킬을 반드시 적용한다:
- Outlook 일정/메일에 한글 제목 또는 본문이 포함될 때
- PowerShell 스크립트에서 한글 문자열을 다룰 때
- 파일을 저장할 때 파일명 또는 내용에 한글이 있을 때
- 웹 폼에 한글을 입력할 때

## 핵심 규칙 (NEVER 위반 금지)

### PowerShell 한글 처리
1. 스크립트 최상단에 반드시 UTF-8 강제 설정:
    ```powershell
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
    ```
2. 파일 저장 시 반드시 Unicode BOM 유지:
    ```powershell
    Set-Content -Path $path -Value $content -Encoding Unicode
    # 또는 [System.IO.File]::WriteAllLines($path, $content, [System.Text.Encoding]::Unicode)
    ```
3. **절대 `Out-File` 기본값을 믿지 않는다** — 인코딩을 항상 명시

### Outlook 일정/메일 한글 제목
1. PowerShell COM 객체 사용 시 반드시 **UTF-16 LE(Unicode BOM)** 인코딩 .ps1 파일로 실행
2. 한글 제목은 **유니코드 코드포인트 배열 조립** 방식 최우선:
    ```powershell
    # '회의' 를 만들 때
    $title = [char]54924 + [char]51032 # 회 + 의

    # '토요임원회의' 를 만들 때
    $title = -join @([char]53664, [char]50836, [char]51076, [char]50896, [char]54924, [char]51032)
    ```
3. 절대 직접 한글 문자열을 하드코딩하지 않는다 ("회의" ← 이렇게 쓰지 않음)
4. 등록 완료 후 반드시 실제 Outlook에서 한글이 정상 표시되는지 검증

### 파일명 한글 처리
- 파일 생성/이동 시 한글 파일명은 NFC 정규화 확인
- cmd.exe에서 한글 경로 사용 시 `chcp 65001` 선행

### 검증 단계 (필수)
모든 한글 관련 작업 완료 후:
1. 출력 결과에서 한글이 깨지지 않았는지 육안 확인
2. 깨진 문자 패턴 감지: `횜`, `?`, `□`, `ﾈ` 등이 보이면 즉시 중단
3. 깨진 경우 원인 파악 후 `[LESSON]`으로 MEMORY.md에 기록

## 자주 발생하는 실패 패턴
| 증상 | 원인 | 해결 |
|------|------|------|
| '횜의'로 표시 | ps1 파일이 UTF-8로 저장됨 | UTF-16 LE BOM으로 저장 |
| '???' 로 표시 | 인코딩 미지정 | Set-Content -Encoding Unicode |
| 제목이 빈칸 | 코드포인트 오류 | char 값 재확인 |
| 본문만 깨짐 | COM 객체 인코딩 불일치 | Body 대신 HTMLBody 사용 고려 |

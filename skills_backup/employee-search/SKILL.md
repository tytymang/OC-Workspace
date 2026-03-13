---
name: employee-search
description: Active Directory(AD)를 사용하여 사내 임직원 정보를 정확하게 검색하고 보고하는 스킬. 사번이나 부서 등의 정보를 찾을 때 브라우저 팝업 대신 빠르고 정확하게 검색합니다.
---

# employee-search

Active Directory(AD)를 사용하여 사내 임직원 정보를 빠르게 검색합니다.

## Trigger
- 특정 임직원의 사번, 부서, 직급 등 인적 정보를 요청할 때
- "누구누구 사번", "최현구 상무 소속" 등의 지시

## Context
브라우저 팝업은 느리고 오류가 많으며, 매번 스크립트를 작성하여 실행하면 딜레이가 발생합니다. 가장 빠른 응답을 위해 이 스킬은 워크스페이스에 미리 생성해둔 `search_ad.ps1` 파일을 호출하여 **1초 내에** 결과를 반환하는 것을 목표로 합니다.

## Execution Steps
1. 검색할 이름(`이름`)을 파악합니다.
2. 아래 명령어를 `exec` 도구로 실행합니다.
   ```powershell
   powershell.exe -ExecutionPolicy Bypass -File .\search_ad.ps1 "이름"
   ```
   *(주의: 위 명령어의 `이름` 부분을 찾고자 하는 실제 이름(예: `최현구`)으로 치환하세요.)*
3. 결과를 즉시 보고합니다. (동명이인 주의)

## Dependencies
이 스킬은 `C:\Users\307984\.openclaw\workspace\search_ad.ps1` 파일에 의존합니다. (이미 워크스페이스에 UTF-16 LE 인코딩으로 영구 저장되어 있습니다.)

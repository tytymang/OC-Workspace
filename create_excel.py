import pandas as pd
import os

# Data for 2본부 (From original Excel)
data_excel = [
    ["2본부", "Chu", "AI 기반 교육시스템", "지식 표준화(SOP/Checklist), AI 훈련 시스템 구축(ChatGPT+Python), 퀴즈 및 자동 평가 시스템, 파일럿 운영", "1월~2월"],
    ["2본부", "Chu", "Daily Dashboard AI", "일일 생산 KPI 정의, 데이터 소스(MES/HR 등) 표준화, HMI 대시보드 설계, AI 분석/일일 요약, Warning/Threshold 관리", "1월~2월"],
    ["2본부", "Chu", "Production Plan AI", "데이터 소스 표준화(YMS/SAP), AI 로직 정의/프롬프트 생성, HMI 인터페이스 생성, 양산 적용", "1월~2월"],
    ["2본부", "선", "공정 불량 시각화", "공정/WORST 불량 정의, 데이터 정합성 확보, AI 최적화 로직 설계/시뮬레이션, 분석 자동화, 파일럿, 대시보드 구축", "1월~"],
    ["2본부", "Anh", "장비 PM 카운트 자동화", "장비 PM 카운트 계산 S/W 개발, 데이터 Path 셋업, 로직 처리 및 Excel 추출, 다수 설비 자동화 배포", "1월~2월"],
    ["2본부", "Anh", "MTBF/MTTR 자동화", "Trouble/Manual Knowledge Base 설계, Log 표준화, RAG 챗봇 PoC, CBR 유사사례 추천, Checklist 자동 생성", "1월~2월"],
    ["2본부", "강", "데이터 처리 자동화", "라우터 점검 자동화, Probing/신뢰성/CAS 측정 Data 비교 및 정리 자동화, Re-probing 분류 Bin 생성 자동화", "1월~"],
    ["2본부", "강", "리포트 자동화", "Meeting 회의록 작성 시간 단축, Email 작성 시간 단축 및 표준 포맷 적용", "1월~"],
    ["2본부", "이", "라인/모델 체인지 관리", "데이터 표준화/계산 로직 정립, WIP/Movement/Reticle Change 계산 룰 정의 및 코딩, Chart Total Report 적용", "1월~3월"],
    ["2본부", "이", "변경점 관리", "Backend(API/이력/승인) 개발, Dashboard 개발, AI 엔진(Rule/이상탐지) 개발, AI 튜닝, UAT, SOP 정립", "1월~3월"],
    ["2본부", "이", "측정 Calibration", "데이터 Input/Output 요구사항 정의, Calibration 알고리즘 설계, 애플리케이션 UI 구성", "1월~3월"],
    ["2본부", "박, 채, 송", "Data 자동 분석", "측정 및 신뢰성 데이터 분석, 형식 생성/코딩, 형광체 농도 계산 명령 생성, 실 사용 적용 및 교육", "1월~3월"],
    ["2본부", "박, 채, 송", "RoHS / MSDS", "AI 기반 RoHS 및 MSDS 사전 검토, 문제점 확인/개선, QA팀 컨펌 후 실무 적용", "1월~3월"],
    ["2본부", "박, 채, 송", "고객 스펙트럼 제출", "샘플 측정, 스펙트럼 라이브러 구축, AI 학습/분석 결과 제공, CIE 스펙트럼 시뮬레이션 검증", "1월~3월"],
    ["2본부", "Khai", "설비 알람 수집 자동화", "Prober 알람 카운터 자동 수집 S/W, 카운터 규칙 정의, Collector 모듈 개발, 데이터 검증/파싱, 파일럿 및 전개", "1월~2월"],
    ["IT WEST GT", "이용혁", "맞춤형 메일 시스템", "CRM 현황 학습, 고객 Activity/성향/개인신상 학습, 타겟 메일 발송 자동화 (월 3500건 발송 업무 개선)", "1분기~2분기"],
    ["IT WEST GT", "박성민", "FCST 및 물동 분석 AI", "고객사 FC Trend 과거/미래 학습, 이동 물동 과거 Trend 학습, SCP FC vs PO balance Trend 학습 검증", "1분기~2분기"],
    ["CS GT", "이영주", "AI 불량 원인 추론 등", "Webex 화자 구분/회의록 템플릿 구현, 불량 유형별 로직 학습 및 매핑 실무 연동", "1월~3월"],
]

# Data from PPT Images
data_ppt = [
    ["PM 1 GT", "박기연", "동시 통역 AI 및 판가/매출 분석", "단순 반복 업무 분류 및 AI 분석 학습 데이터 수집, 시범 적용 및 실무 적용", "상반기"],
    ["PM 2 GT", "박재현", "영업/마케팅 AI 자동화", "NPI 승인 피드백 자동화, 전시회 데모 컨텐츠 작성, 판가 적자 이슈 확인/조정안 제시", "상반기"],
    ["R&D 센터 ST", "서대웅", "설계 예측 및 제조 최적화", "Chip/EPI 설계 예측 모델, AOI 자동화 Recipe 셋업, E-FAB 조기진단 및 수율 자동 보정", "상반기"],
    ["FT 센터", "임완태", "분석 보고서 및 설계 최적화", "분석 보고서 자동화(EDX 연동), 형광체 혼합비 배합 AI, Mechanical LT 단축(3D CAD 평가)", "상반기"],
    ["CN 1센터", "이종국", "CHIP 발주 AI 산출", "과거사용량/재고/PO잔량 취합 로직 개발 및 월 발주량 자동 산출", "1분기"],
    ["CN 2 GT", "안병길", "월별 CI 금액 자동 산출", "AI를 활용한 Raw Data 취합 및 금액 산출 리포트 자동화 프로그램 생성", "1분기"],
    ["CN 법인", "이상혁", "PPT 작성 시간 90% 단축", "Python-pptx 라이브러리와 AI 결합 회사 양식 PPT 자동 생성, 법인 인원 교육", "상반기"],
    ["EU 법인", "조성현", "리스크 관리 및 HR 자동화", "수주/매출/AR 기반 고객사 사전 부도 징후 Check Tool 개발, 판가 유지율 분석, 인사 Warning letter 자동 생성", "상반기"],
    ["JP 법인", "이민수", "판가 및 프로젝트 분석", "Gemini Antigravity 활용 판가/매출 분석 툴 및 PJT 진행 현황 분석", "상반기"],
    ["NA 법인", "박창진", "판가/매출 및 부도 징후 분석", "판가 분석, 매출/PJT 분석, 부도 징후 분석 (생산성 30% 향상)", "상반기"],
    ["HQ 제조본부", "강지훈", "수율 개선 및 매크로 업무 자동화", "Run 간 산포 관리, 생산계획 예측, 설비별 예방 정비 및 안전 재고 관리", "상반기"],
    ["MD 사업/제조", "류승열", "AI 적용 현황 종합 관리", "본부 및 각 부문별 AI 적용 건수 파악 및 일정 미준수 건수 집중 관리", "상시"],
    ["MD GT", "류승열", "E-mail / Follow-up 자동화", "그룹웨어 연동 AI 이메일 템플릿 설계, AI 비서 활용 일정표 및 Plan 관리", "상반기"],
    ["HI GT", "손민수", "FCST 예측 및 고객 이탈 방지", "시장 트렌드 기반 수요 예측/맥락 기반 이메일 자율 발송, 고객 3개년 활동 기반 이탈 징후 분석", "상반기"],
    ["IT EAST GT", "안윤순", "SCM 차질 분석 및 MS 시뮬레이션", "PO 대비 재고/차질 자동 계산(Balance Sheet) Neck Point 도출, 경쟁사 가격 기반 자사 M/S 및 매출 시뮬레이션", "상반기"],
    ["재무 GT", "신재준", "재무/비용 Fully Automated KPI", "Actual BOM 자동 생성, SAP 연동 Selling Price 즉시 산출 AI, 비용 대사, FCT 결제 리뷰", "상반기"],
    ["VN Support", "김성주", "인프라 및 감사 AI 지원", "외국 법률/분쟁 검토(Gemini), 내부감사 부정 식별 AI 모델, AI 재고 분석, Chiller water 분석, 원자재 재료비 분석 Tool", "상반기"],
    ["VN HR", "박상준", "HR 자동화", "채용 시 우수 인원 선별 AI 및 조직문화 규정 Q&A 챗봇 구축", "1분기"],
    ["VN EPI", "최효식", "PM 예방보전 및 WD 예측출하", "PM 예방보전 중요 인자 머신러닝 분석 및 WD 예측 출하 AI", "상반기"],
    ["VN 3본부", "구일회", "생산성 및 재고 최적화", "UPH/병목 기반 AI 인력 최적 배치, 적정재고 안전 발주 가이드 시스템, 실시간 생산성 대시보드", "상반기"],
]

df1 = pd.DataFrame(data_excel, columns=["섹터(부서)", "담당자", "AI 과제 분류", "세부 추진 과제", "추진 기간"])
df2 = pd.DataFrame(data_ppt, columns=["섹터(부서)", "담당자", "AI 과제 분류", "세부 추진 과제", "추진 기간"])

df_total = pd.concat([df1, df2], ignore_index=True)
out_path = r"C:\Users\307984\.openclaw\workspace\Total_AI_Tasks_2026.xlsx"
df_total.to_excel(out_path, index=False)
print("Excel created at", out_path)

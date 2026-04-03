# MetCalib 프로젝트 컨텍스트

## 프로젝트 개요
KAERI 기상관측장비 교정 프로그램. 기존 VB6 레거시(MetCalib.exe, 소스코드 없음)를 Python + PyQt6로 완전히 재구축.
기존 DB 마이그레이션 불필요 — 새로 시작. PDF 교정 성적서 양식 새로 설계.

## 상세 계획서
docs/PLAN.md 참조 (DB 스키마, UI 레이아웃, 클래스 설계, 7단계 구현 순서 포함)

## 교정 항목 (5종)
- 풍향: 8점 (0~315°, 45° 간격), ±5°
- 풍속: 8점 (0,1,3,4,5,7,10,15 m/s), ±0.2~1.5
- 온도: 9점 (-10~30°C, 5° 간격), ±0.3°C
- 습도: 7점 (30~90%, 10% 간격), ±3%
- 팬: 1점, ±1

## 센서 장비
- 온도: PT100 (백금저항) → 반드시 Logger 경유 (자체 시리얼 출력 없음)
- 습도: Vaisala HMP 시리즈 → 직접 시리얼 또는 Logger
- 풍향풍속: Vaisala WMT103 → 직접 시리얼 또는 Logger

## 연결 방식 (장비 유형에 종속되지 않음, 세션 구성 시 자유 선택)
- A. 직접 시리얼 (pyserial + 플러그인 파서): HMP 파서, WMT103 파서 2종
- B. CR1000류 데이터로거 경유 (PyCampbellCR1000): PakBus 프로토콜
  - CR1000은 단종. CR1000X, CR300 등 후속 모델 사용. "CR1000류"로 표기할 것.

## 핵심 설계 원칙
- 동시 교정 최대 약 5대
- 장비 수, 연결 방식, 센서 조합이 매번 달라짐 → 세션 구성 화면에서 동적 추가
- "풍향=시리얼, 온습도=Logger" 같은 고정 매핑 없음

## 기술 스택
- Python 3.11+ / PyQt6 / pyqtgraph / pyserial / PyCampbellCR1000
- SQLite / reportlab / openpyxl / PyInstaller

## GitHub/네트워크 환경
- SSH 정상 (git@github.com:mesafalcon0/MetCalib.git)
- HTTPS TLS 인증서 문제 (회사 프록시) → curl -sk 우회, gh CLI는 unset GH_TOKEN 필요
- Git Bash에서 MSYS_NO_PATHCONV=1 필요 (GitHub API URL 경로 변환 방지)

## 작업 방식
- 한국어로 소통
- 구현 전 계획 충분히 검토 → 실제 운용 환경(다중 장비, 가변 구성) 반드시 고려
- 급하게 코딩 시작하지 말 것

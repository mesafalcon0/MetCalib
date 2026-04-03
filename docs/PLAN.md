# MetCalib 기상관측장비 교정 프로그램 재구축 계획

## Context

MetCalib은 KAERI(한국원자력연구원)에서 기상관측장비 교정에 사용하는 VB6 레거시 프로그램이다. 소스코드 없이 컴파일된 EXE만 존재하며, 15년 이상 된 기술 스택(VB6 + Access DB)으로 유지보수가 어렵다. 이를 **Python + PyQt6**로 완전히 재구축하면서, PDF 교정 성적서 자동 생성 등 새로운 기능을 추가한다.

---

## 1. 현재 프로그램 분석 요약

### 기술 스택
- VB6 EXE (368KB) + Access DB (.mdb, 2.4MB)
- ChartRcd.ocx(차트), MSCOMM32.OCX(시리얼), Pic2Png.dll(화면캡처)

### DB 구조 (3개 테이블)
- **Terms** (29행): 장비유형(Wd/Ws/Tm/Hm), 관측소(KAERI/KJRR/ARA), 운용상태
- **Instruments** (29행): 장비 등록정보 (유형, 일련번호, 교정일, 자산번호)
- **Calib** (110행): 교정 측정값 (InstID, S_N, Logger, 날짜, Cal_0~Cal_9)

### 교정 항목 (5종)
| 항목 | 교정점 | 허용오차 |
|------|--------|---------|
| 풍향 | 8점 (0~315°, 45° 간격) | ±5° |
| 풍속 | 8점 (0,1,3,4,5,7,10,15 m/s) | ±0.2~1.5 |
| 온도 | 9점 (-10~30°C, 5° 간격) | ±0.3°C |
| 습도 | 7점 (30~90%, 10% 간격) | ±3% |
| 팬 | 1점 | ±1 |

### UI 워크플로우
1. 시리얼 포트 연결 → 장비 선택
2. 각 교정점: 측정 시작 → 실시간 차트 표시 → 10초/1분 평균±표준편차 계산 → 완료
3. DB 저장 + PNG 캡처 (CalibData_YYYYMM/)

---

## 2. 새 프로그램 설계

### 2.1 기술 스택
| 항목 | 기술 |
|------|------|
| 언어 | Python 3.11+ |
| GUI | PyQt6 |
| 데이터로거 통신 | PyCampbellCR1000 (mesafalcon0/PyCampbellCR1000) |
| 시리얼 기반 | pyserial (PyCampbellCR1000 내부 의존) |
| 차트 | pyqtgraph (실시간 성능 우수) |
| DB | SQLite (sqlite3 표준 라이브러리) |
| PDF | reportlab |
| Excel | openpyxl |
| 패키징 | PyInstaller |

### 2.1.1 장비 연결 방식 (2종 혼합)

실제 운용 환경에서 장비 구성은 매번 달라지며, 두 가지 연결 방식이 혼합된다:

**연결 방식은 장비 유형에 종속되지 않는다.**
풍향/풍속/온도/습도 어떤 센서든 아래 두 방식 중 상황에 맞게 선택:

**A. 직접 시리얼 (pyserial + 플러그인 파서)**
- 센서 → RS-232 → PC 직접 연결
- 습도(HMP시리즈), 풍향풍속(Vaisala WMT103)에서 사용
- 온도(PT100)는 자체 시리얼 출력이 없어 직접 연결 불가 → 반드시 Logger 경유
- 파서 2종 필요: HMP 파서, WMT103 파서
- pyserial로 구현

**B. Campbell CR1000류 데이터로거 경유 (PyCampbellCR1000)**
- 센서 → CR1000류 Logger → PC (시리얼 또는 TCP)
- CR1000은 단종되었으나 CR1000X, CR300 등 후속 모델이 동일 PakBus 프로토콜 사용
- 모든 장비 유형에 사용 가능
- 1 Logger에 1~N개 센서 동시 연결 가능 (가변)
- PyCampbellCR1000 라이브러리 — PakBus 프로토콜 (CR1000류 호환)
- 연결 URL: `serial:COMx:38400` 또는 `tcp:host:port`

**핵심 설계 원칙**:
- 장비 수, 연결 방식, 센서 조합이 매번 달라짐
- 교정 세션 시작 시 **장비마다 연결 방식을 자유롭게 선택**
- "풍향=시리얼, 온습도=Logger" 같은 고정 매핑 없음

### 2.2 프로젝트 구조

```
metcalib-py/
├── metcalib/
│   ├── __init__.py              # 버전, 앱 메타데이터
│   ├── app.py                   # QApplication 진입점
│   ├── config.py                # INI 기반 앱 설정
│   │
│   ├── db/
│   │   ├── schema.py            # SQLite DDL, 마이그레이션
│   │   ├── models.py            # 데이터 모델 (dataclass)
│   │   ├── repository.py        # CRUD 연산
│   │   └── migrate_mdb.py       # Access .mdb → SQLite 마이그레이션 (선택적)
│   │
│   ├── connection/
│   │   ├── base.py              # BaseConnection 추상 클래스
│   │   ├── direct_serial.py     # 직접 시리얼 연결 (HMP, WMT103 등)
│   │   ├── campbell_logger.py   # CR1000류 Logger 연결 (PakBus)
│   │   ├── reader_thread.py     # QThread 기반 데이터 수신 (공통)
│   │   ├── channel.py           # 센서 채널 매핑 (1 Logger → N 센서)
│   │   └── parsers/
│   │       ├── base.py          # BaseParser 추상 클래스
│   │       ├── wmt103.py        # Vaisala WMT103 풍향풍속계 파서
│   │       └── hmp.py           # Vaisala HMP 시리즈 습도계 파서
│   │
│   ├── calibration/
│   │   ├── engine.py            # 10초/1분 평균, 표준편차 계산
│   │   ├── spec.py              # CalibSpec 데이터클래스
│   │   ├── spec_loader.py       # MetCalib.Lst 파서
│   │   ├── session.py           # 교정 세션 (다중 장비 관리)
│   │   ├── sensor_task.py       # 개별 센서 교정 태스크 (상태 머신)
│   │   └── validator.py         # 합격/불합격 판정
│   │
│   ├── ui/
│   │   ├── main_window.py       # QMainWindow (메인 화면)
│   │   ├── widgets/
│   │   │   ├── session_setup.py   # 교정 세션 구성 (장비/연결 동적 추가)
│   │   │   ├── sensor_panel.py    # 개별 센서 패널 (차트+테이블+버튼)
│   │   │   ├── chart_widget.py    # pyqtgraph 실시간 차트
│   │   │   ├── calib_table.py     # 교정점 테이블
│   │   │   └── overview_panel.py  # 전체 세션 요약/진행상황
│   │   ├── dialogs/
│   │   │   ├── instrument_manager.py  # 장비 등록/관리
│   │   │   ├── history_viewer.py      # 교정 이력 조회
│   │   │   ├── settings_dialog.py     # 설정
│   │   │   └── export_dialog.py       # 내보내기
│   │   └── resources/
│   │       └── styles.qss             # Qt 스타일시트
│   │
│   ├── export/
│   │   ├── pdf_certificate.py   # PDF 교정 성적서 생성
│   │   ├── screenshot.py        # PNG 화면 캡처
│   │   ├── excel_export.py      # Excel 내보내기
│   │   └── csv_export.py        # CSV 내보내기
│   │
│   └── i18n/
│       └── ko.py                # 한국어 문자열 상수
│
├── tests/
├── scripts/
│   ├── migrate_from_access.py   # Access→SQLite 마이그레이션 CLI
│   └── build_exe.py             # PyInstaller 빌드
├── pyproject.toml
└── requirements.txt
```

### 2.3 DB 스키마 (SQLite)

```sql
-- 용어/분류 (Terms 테이블 확장)
CREATE TABLE terms (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    category    TEXT NOT NULL,    -- 'device_type', 'site', 'work_status'
    code        TEXT NOT NULL,    -- 'Wd', 'Ws', 'r67m', 'Cali' 등
    label_ko    TEXT NOT NULL,    -- '풍향계', 'KAERI 67m' 등
    sort_order  INTEGER DEFAULT 0,
    UNIQUE(category, code)
);

-- 장비 등록
CREATE TABLE instruments (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    inst_type   TEXT NOT NULL,    -- 'Wd','Ws','Tm','Hm','WdWs','Lg' 등
    idx         INTEGER,
    serial_no   TEXT NOT NULL UNIQUE,
    site        TEXT,
    assets_no   TEXT,
    comment     TEXT,
    created_at  TEXT DEFAULT (datetime('now','localtime')),
    updated_at  TEXT DEFAULT (datetime('now','localtime'))
);

-- 교정 결과 (헤더)
CREATE TABLE calibrations (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    inst_id      INTEGER NOT NULL REFERENCES instruments(id),
    logger_sn    TEXT,
    calib_type   TEXT NOT NULL,   -- 'Wind Direction' 등
    performed_at TEXT NOT NULL,
    operator     TEXT,
    pass_fail    TEXT,            -- 'PASS' / 'FAIL'
    pdf_path     TEXT,
    png_path     TEXT,
    notes        TEXT,
    created_at   TEXT DEFAULT (datetime('now','localtime'))
);

-- 교정점 측정값 (정규화: Cal_0~Cal_9 대신 행 단위)
CREATE TABLE calibration_points (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    calibration_id  INTEGER NOT NULL REFERENCES calibrations(id) ON DELETE CASCADE,
    point_index     INTEGER NOT NULL,
    reference_value REAL NOT NULL,
    tolerance       REAL NOT NULL,
    measured_value  REAL,
    avg_10s         REAL,
    std_10s         REAL,
    avg_1m          REAL,
    std_1m          REAL,
    error           REAL,
    pass_fail       TEXT,
    measured_at     TEXT,
    UNIQUE(calibration_id, point_index)
);

-- 교정 사양 (MetCalib.Lst 대체, UI에서 편집 가능)
CREATE TABLE calibration_specs (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    calib_type      TEXT NOT NULL,
    display_format  TEXT DEFAULT '0.00',
    point_index     INTEGER NOT NULL,
    reference_value REAL NOT NULL,
    tolerance       REAL NOT NULL,
    UNIQUE(calib_type, point_index)
);

-- 원시 측정 로그 (감사 추적용, 선택적)
CREATE TABLE measurement_log (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    calibration_id  INTEGER REFERENCES calibrations(id),
    point_index     INTEGER,
    timestamp       TEXT NOT NULL,
    raw_value       REAL NOT NULL,
    raw_serial_data TEXT
);
```

### 2.4 UI 레이아웃

#### A. 세션 구성 화면 (교정 시작 전)

교정할 장비 목록과 연결 방식을 동적으로 구성한다.

```
+----------------------------------------------------------------------+
| [파일] [장비관리] [내보내기] [설정] [도움말]              메뉴바      |
+----------------------------------------------------------------------+
|                     교정 세션 구성                                    |
+----------------------------------------------------------------------+
| 교정 유형: [풍향▼]  관측소: [KAERI 67m▼]     날짜: [2025-03-28]     |
+----------------------------------------------------------------------+
| # | 센서 S/N        | Logger S/N  | 연결 방식           | 상태      |
|---|-----------------|-------------|---------------------|-----------|
| 1 | W-2255          | (없음)      | 직접시리얼:COM3     | 미연결    |
| 2 | W-2288          | (없음)      | 직접시리얼:COM4     | 미연결    |
| 3 | W-2289          | (없음)      | 직접시리얼:COM5     | 미연결    |
|   |    [+ 장비 추가]    [- 제거]    [연결 테스트]                     |
+----------------------------------------------------------------------+
| [모두 연결]                                [세션 시작]               |
+----------------------------------------------------------------------+
```

#### B. 단일 센서 교정 (1대 교정 시)

```
+----------------------------------------------------------------------+
| [세션구성] [장비관리] [내보내기] [설정]                   메뉴바      |
+----------------------------------------------------------------------+
| 풍향 교정 | W-2255 | COM3 연결됨                          [초기화]  |
+----------------------------------------------------------------------+
|                                                                      |
|   pyqtgraph 실시간 차트 (자동스크롤, 줌/팬 지원)                    |
|                                               현재값: 45.3 ──       |
|                                                                      |
+----------------------------------------------------------------------+
| 기준값 | 측정값 | 10초 평균±표준편차 | 1분 평균±표준편차 | 동작     |
|--------|--------|-------------------|------------------|----------|
|   0    |  -0.1  |   -0.1 ± 0.3     |   0.0 ± 0.1     | [완료]   |
|  45    |  45.3  |   45.0 ± 0.3     |  45.0 ± 0.1     | [완료]   |
|  90    |        |                   |                  |[측정시작]|
| ...    |        |                   |                  |[측정시작]|
+----------------------------------------------------------------------+
| 상태: raw data stream                              | Connected COM3  |
+----------------------------------------------------------------------+
```

#### C. 다중 센서 동시 교정 (챔버 내 여러 센서)

탭 방식으로 각 센서별 교정 화면 + 전체 요약 탭 제공

```
+----------------------------------------------------------------------+
| [세션구성] [장비관리] [내보내기] [설정]                   메뉴바      |
+----------------------------------------------------------------------+
| [요약] [H-1502001/L-1608] [H-60947194/L-2013] [H-1502003/L-2215]   |
+----------------------------------------------------------------------+
|                     << 요약 탭 선택 시 >>                            |
| # | 센서 S/N    | Logger   | 진행 | 현재값 | 합격점 | 상태         |
|---|-------------|----------|------|--------|--------|--------------|
| 1 | H-1502001   | L-1608   | 5/7  | 72.4   |  5/5   | 측정중...    |
| 2 | H-60947194  | L-2013   | 5/7  | 72.1   |  5/5   | 측정중...    |
| 3 | H-1502003   | L-2215   | 4/7  | 60.8   |  4/4   | 대기         |
+----------------------------------------------------------------------+
|  [전체 측정시작]  [전체 완료]           [전체 저장]  [PDF 일괄생성]  |
+----------------------------------------------------------------------+
```

**색상 표시**: 노란색=측정중, 초록색=합격, 빨간색=불합격

### 2.5 핵심 클래스

**연결 계층** (connection/):
- **BaseConnection(ABC)**: 연결 추상 클래스 (`connect()`, `disconnect()`, `read_value()`)
- **DirectSerialConnection**: pyserial 직접 연결 (풍향/풍속용)
- **CR1000Connection**: PyCampbellCR1000 Logger 연결 (온습도용)
- **SensorChannel**: 1 Logger 내 센서 채널 매핑 (Logger → 테이블컬럼 → 센서)
- **ReaderThread(QThread)**: 데이터 수신 스레드, `value_parsed(sensor_id, float, datetime)` 시그널

**교정 계층** (calibration/):
- **CalibSession**: 교정 세션 전체 관리 (다중 센서 태스크 목록 보유)
- **SensorTask**: 개별 센서 교정 상태 머신 (IDLE → MEASURING → POINT_COMPLETE → DONE)
- **CalibEngine**: deque 기반 10초/1분 롤링 버퍼, 평균·표준편차 계산 (센서당 1개 인스턴스)
- **CalibSpec**: 교정 사양 데이터클래스 (교정점, 허용오차, 표시 형식)
- **CalibValidator**: 합격/불합격 판정

**데이터 계층** (db/):
- **CalibRepository**: SQLite CRUD (장비, 교정결과, 이력 조회)

**출력 계층** (export/):
- **PDFCertificateGenerator**: reportlab 기반 PDF 성적서 생성 (개별/일괄)

---

## 3. 구현 순서

### Phase 1: 기반 구축
- 프로젝트 설정 (pyproject.toml, 의존성)
- SQLite DB 스키마 생성
- 데이터 모델 (dataclass)
- MetCalib.Lst 파서 구현
- 기본 MainWindow 골격 + 세션 구성 화면

### Phase 2: 연결 계층 (2종 혼합)
- BaseConnection 추상 클래스
- DirectSerialConnection (풍향/풍속 직접 시리얼)
- CR1000Connection (PyCampbellCR1000, 온습도 Logger 경유)
- SensorChannel (1 Logger → N 센서 채널 매핑)
- ReaderThread (센서별 데이터 수신, sensor_id로 분배)
- 연결 테스트 기능

### Phase 3: 단일 센서 교정 (핵심 워크플로우)
- CalibEngine (10초/1분 통계)
- SensorTask (개별 센서 상태 머신)
- SensorPanel 위젯 (차트 + 테이블 + 버튼 통합)
- pyqtgraph 실시간 차트
- 측정시작/완료/합격판정
- 1대 장비로 전체 파이프라인 검증

### Phase 4: 다중 센서 동시 교정
- CalibSession (다중 SensorTask 관리)
- 탭 기반 다중 SensorPanel
- 요약(Overview) 패널 (전체 진행상황)
- 전체 측정시작/완료 일괄 제어
- 다중 ReaderThread 관리

### Phase 5: 데이터 저장 + 출력
- 교정 결과 DB 저장 (세션 단위 일괄 저장)
- PNG 화면 캡처 (기존 명명규칙 유지)
- PDF 교정 성적서 생성 (개별/일괄)
- Excel/CSV 내보내기

### Phase 6: 장비 관리 + 이력
- 장비 등록/편집 다이얼로그
- 교정 이력 조회 다이얼로그
- 교정 사양 편집 (UI에서 교정점/허용오차 변경)
- 설정 다이얼로그

### Phase 7: 마무리
- PyInstaller EXE 패키징
- 에러 처리/로깅 (시리얼 끊김, DB 오류 복구)
- UI 한국어 검수

---

## 4. 핵심 참조 파일

| 파일 | 용도 |
|------|------|
| `C:/Users/User/Documents/MetCalib/MetCalib.Lst` | 교정 사양 원본 (파싱 대상) |
| `C:/Users/User/Documents/MetCalib/MetCalib.mdb` | 기존 DB (마이그레이션 원본) |
| `C:/Users/User/Documents/MetCalib/CalibData_202411/*.png` | UI 레이아웃 참조 스크린샷 |
| `C:/Users/User/Documents/met/db_common.py` | Access ODBC 연결 패턴 재사용 |
| `C:/Users/User/Documents/met/met_spec.py` | cp949 INI 파싱 패턴 참조 |
| `github.com/mesafalcon0/PyCampbellCR1000` | CR1000류 데이터로거 PakBus 통신 라이브러리 (CR1000X, CR300 등 호환) |

---

## 5. 검증 방법

1. **단위 테스트**: spec_loader, engine(통계 계산), validator(합격판정), repository(CRUD)
2. **파서 테스트**: WMT103, HMP 시리얼 파서의 실제 출력 포맷 파싱 검증
3. **단일 센서 통합 테스트**: 모의 시리얼 데이터 → 파서 → 엔진 → 세션 → DB 저장
4. **다중 센서 통합 테스트**: 복수 모의 연결로 동시 교정 파이프라인 검증
5. **실물 테스트**: 실제 장비 연결하여 전체 교정 사이클 수행

---

## 6. 검토에서 발견된 핵심 설계 결정

### 실제 운용 규모
- **동시 교정 최대 약 5대** (Logger 자체가 많이 연결하지 않음)
- DB에 15대 기록은 같은 날 순차 교정한 누적 결과

### 설계에 반영된 사항
1. **가변 장비 구성**: 장비 수/연결 방식이 매번 달라지므로 세션 구성 화면에서 동적 추가
2. **2종 혼합 연결**: 풍향풍속(직접시리얼) + 온습도(CR1000) 모두 지원
3. **1 Logger → N 센서**: CR1000 채널 매핑으로 하나의 Logger에서 다중 센서 데이터 분리
4. **온습도 복합센서**: 같은 S_N(H-K2340023)으로 Hm과 Tm을 각각 독립 교정
5. **동시+순차 혼합**: 챔버 내 동시 측정(전체 시작/완료) + 개별 센서 독립 제어 모두 가능

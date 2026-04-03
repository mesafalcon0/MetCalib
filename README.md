# MetCalib

기상관측장비 교정 프로그램 (Meteorological Instrument Calibration)

KAERI 기상관측장비의 교정 데이터를 수집, 관리, 분석하는 Windows 데스크톱 프로그램.

## 개요

- **교정 항목**: 풍향, 풍속, 온도, 습도, 팬 (5종)
- **지원 장비**: Vaisala WMT103 (풍향풍속), Vaisala HMP 시리즈 (습도), PT100 (온도)
- **연결 방식**: 직접 시리얼 (RS-232) / Campbell CR1000류 데이터로거 (PakBus)
- **동시 교정**: 최대 약 5대

## 기술 스택

- Python 3.11+ / PyQt6
- pyqtgraph (실시간 차트)
- pyserial / PyCampbellCR1000 (장비 통신)
- SQLite (데이터 저장)
- reportlab (PDF 교정 성적서)

## 상태

계획 수립 중

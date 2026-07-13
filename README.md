# 🚨 진주품질개선팀 고장 현황 대시보드

진주품질개선팀 고장 현황을 실시간으로 확인하는 Streamlit 웹 대시보드입니다.

## 탭 구성

| 탭 | 내용 |
|---|---|
| 🔔 금일 알람 | 오늘 날짜 기준 신규 알람 (5G_LTE OOS_진주 시트) |
| 🕐 13시 알람 | 13시 알람 공유 시트 |
| 📋 전체 RAW | 전체 고장 데이터 검색/필터 |
| 📊 대시보드 | 복구/미복구 현황, 고장구분 차트 |
| 📡 중계기·MIBOS | 중계기 및 MIBOS·gREMS 알람 |
| 🔔 기타 알람 | RMS A망·DACS·통합RCU 미복구 |
| 📶 VSWR 알람 | 5G/LTE VSWR 알람 + PORT현황 매칭 |

## 사용 방법

1. Streamlit 앱에 접속
2. 고장 RAW DATA 엑셀 파일(`.xlsx`) 업로드
3. 각 탭에서 현황 확인

## 파일 구조 (업로드할 엑셀 시트 목록)

- `5G_LTE OOS_진주` — 메인 고장 데이터
- `13시 알람 공유` — 13시 알람
- `중계기 및 MIBOS 알람` — 중계기 현황
- `gREMS` — gREMS 알람
- `RMS_A망 미복구` / `RMS_DACS 미복구` / `RMS_통합RCU미복구`
- `5G Raw` / `LTE Raw` / `3G MOD ` — TT 알람 원본
- `AAU PORT_0518` / `RRU PORT_0518` / `3G MOD_1027` — PORT 현황

## 로컬 실행

```bash
pip install -r requirements.txt
streamlit run app.py
```

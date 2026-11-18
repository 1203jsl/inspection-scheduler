# 검수일정 관리 — 엑셀 매크로판

회사 규정상 별도 프로그램(Electron) 대신 엑셀 매크로로 쓰는 버전입니다.

- 배포 파일: `dist/검수일정관리_엑셀.zip` (검수일정관리.xlsx + modInspection.bas + clsAppEvents.cls + 설치·사용 안내.txt)
- 설치·사용 방법: `help.txt` (엑셀 안의 [사용안내] 시트와 같은 내용)

## 만들기
```bash
./build.sh
```
- `build_workbook.py` : 시트·서식·드롭다운·설정·메일 양식·화면 문구(숨김 `_text` 시트)를 만든 .xlsx 생성 (openpyxl)
- `vba/part*.bas` : 매크로 코드. `build.sh`가 `vba/modInspection.bas` 하나로 합칩니다. 코드에는 한글이 없고(코드 페이지 문제 방지) 화면 문구는 모두 `texts.py` → `_text` 시트에서 읽습니다.
- `vba/clsAppEvents.cls` : 엑셀 이벤트(셀 변경·더블클릭·저장 후)를 모듈로 넘기는 클래스

## 시험
```bash
python3 test/name_check.py      # 지역 변수가 함수 이름을 가리는지 (VBA는 대소문자 구분 없음)
python3 test/lo_check.py        # LibreOffice로 함수별 문법 검사 + test/tests.bas 계산 시험
```
LibreOffice로는 엑셀 개체(시트·셀·도형) 동작까지는 시험할 수 없어 실제 엑셀에서 확인이 필요합니다.

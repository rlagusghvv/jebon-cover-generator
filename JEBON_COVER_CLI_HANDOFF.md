# JEBON Cover Generator - CLI Handoff

Last updated: 2026-03-04 (KST)

## 0) 키워드 트리거 (`제본`)
- 사용자가 `제본`이라고 말하면 이 프로젝트를 의미함.
- 다음 CLI는 즉시 아래 파일부터 읽고 이어서 작업할 것:
  1. `/Users/kimhyeonho/JEBON_COVER_CLI_HANDOFF.md`
  2. `/Users/kimhyeonho/cover_generator.py`
  3. `/Users/kimhyeonho/jebon-cover-generator/README.md`
- 기본 목표: `권/날짜/지급번호` 기반 대량 표지 PDF 생성 워크플로 유지/개선.

## 1) 현재 산출물 위치
- 메인 소스: `/Users/kimhyeonho/cover_generator.py`
- 실행 앱(최신): `/Users/kimhyeonho/Desktop/JEBONCoverGenerator.app`
- 빌드 산출물: `/Users/kimhyeonho/dist/JEBONCoverGenerator.app`
- PyInstaller spec: `/Users/kimhyeonho/JEBONCoverGenerator.spec`
- 테스트 출력 폴더 예시: `/Users/kimhyeonho/Desktop/JEBON_OUTPUT_TEST`

## 2) 프로젝트 목표
- Excel(`.xlsx/.xlsm`) 데이터로 제본 표지 PDF를 대량 생성.
- macOS에서 한글 출력 지원.
- 사용자 친화 GUI + 대량 처리(최대 100권 이상) + 진행 상태 가시화.

## 3) 구현된 주요 기능 (완료)
1. GUI 기반 처리 (`tkinter`)
- Excel 파일 선택
- 출력 폴더 선택
- 진행률/상태/로그 실시간 표시
- 권별 처리 상태 테이블 표시
- 작업 중단 버튼 지원

2. 입력 방식 3종
- Excel 읽기 (`RAW` 시트 A:C = 권/날짜/지급번호)
- 권 범위 자동생성 (`1-100`, `1~100`, `1,3,5-8`)
- 클립보드 일괄입력 (권/날짜/지급번호 3열 복붙)

3. 빠른 일괄 입력
- 작업일/지급명령번호 앱에서 직접 입력 가능
- 생성 시 일괄 덮어쓰기 옵션

4. PDF 생성
- A4 세로
- 상단: `제 본 표 지`
- 중단: `제 [권] 권`
- 하단: `지급명령번호: ...`, `작업일: ...`
- 파일명: `[작업일]_cover_[권].pdf`

5. 폰트 안정화
- 우선 시도: `/System/Library/Fonts/Supplemental/AppleGothic.ttf`
- 실패 시 대체: `/System/Library/Fonts/AppleSDGothicNeo.ttc`
- 이유: 일부 환경에서 AppleGothic이 `fpdf2`에서 `KeyError('OS/2')` 발생 가능

6. Windows 호환 분기 추가
- Windows 폰트 후보 자동 탐색: `C:\\Windows\\Fonts\\malgun.ttf` 등
- 출력 폴더 열기 OS 분기:
  - Windows: `os.startfile`
  - macOS: `open`
  - Linux: `xdg-open`

## 4) 사용자 요구 반영 포인트
- “권마다 날짜/번호 다름” 대응 완료:
  - Excel `RAW` A:C 그대로 읽기
  - 또는 Excel에서 3열 복사 후 `클립보드 일괄입력`
- “100권 이상 대량 처리” 대응 완료:
  - 권 범위 자동생성 + 일괄 생성

## 5) 현재 UI 사용 절차 (권장)
### 케이스 A: 권별 값이 모두 다름
1. 원본 Excel에서 `권/날짜/지급번호` 3열 여러 행 복사
2. 앱에서 `클립보드 일괄입력`
3. 테이블 확인 후 `PDF 생성 시작`

### 케이스 B: 권만 다르고 날짜/번호는 동일
1. `작업일`, `지급명령번호` 입력
2. `권 범위`에 `1-100` 입력
3. `권 자동생성` 클릭
4. `PDF 생성 시작`

### 케이스 C: Excel 데이터 그대로 사용
1. Excel 파일 지정
2. `데이터 확인`
3. 필요 시 일괄 덮어쓰기
4. `PDF 생성 시작`

## 6) 의존성 / 실행 / 빌드
### 개발 실행
```bash
python3 /Users/kimhyeonho/cover_generator.py
```

### 의존성 설치
```bash
pip3 install pandas openpyxl fpdf2 pyinstaller
```

### 앱 빌드
```bash
python3 -m PyInstaller --noconfirm --clean --windowed --name JEBONCoverGenerator /Users/kimhyeonho/cover_generator.py
```

### 데스크톱 앱 교체/실행
```bash
rm -rf /Users/kimhyeonho/Desktop/JEBONCoverGenerator.app
cp -R /Users/kimhyeonho/dist/JEBONCoverGenerator.app /Users/kimhyeonho/Desktop/
open /Users/kimhyeonho/Desktop/JEBONCoverGenerator.app
```

## 7) 검증 이력 (완료)
- 문법 체크: `python3 -m py_compile /Users/kimhyeonho/cover_generator.py`
- 함수 단위 생성 테스트: 실제 PDF 1건 생성 확인
- `.app` 재빌드 및 실행 프로세스 확인 완료

## 8) 알려진 제약 / 다음 작업 후보
1. PDF 표지 좌표는 “요구 텍스트” 기준으로 배치되어 있음
- 기존 Excel `cover` 시트의 시각 배치와 1:1 완전 동일 매칭은 아직 아님
- 필요 시 좌표/폰트/문구를 `generate_cover_pdf()`에서 미세조정 가능

2. 파일명 충돌 정책
- 동일 `작업일 + 권`이면 덮어쓰기됨
- 필요 시 자동 `(_1, _2)` suffix 로직 추가 가능

3. 입력 검증 강화 가능
- 날짜 형식 정규화
- 지급번호 형식 검증
- 빈 값 허용 정책 UI 옵션화

## 9) 참고 데이터 경로 (사용자가 제공한 매크로 파일)
- `/Users/kimhyeonho/Desktop/지출증빙서 표지 자동화건_2026.01.14/지출서류정리목록(자동생성기)_2026-01-17_김현호 수정.xlsm`

## 10) 다른 CLI를 위한 즉시 시작 체크리스트
1. `cat /Users/kimhyeonho/JEBON_COVER_CLI_HANDOFF.md`
2. `python3 /Users/kimhyeonho/cover_generator.py` 실행
3. 사용자 입력 방식 선택 (Excel / 클립보드 / 권범위)
4. 샘플 1~2건 생성 확인
5. 필요 시 `.app` 재빌드

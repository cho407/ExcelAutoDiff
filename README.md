# Excel AutoDiff

`xlsx`, `xlsm` 파일(또는 폴더)을 비교해서 어떤 셀/시트가 바뀌었는지 리포트하는 도구입니다.

## 지원 범위

- 파일 vs 파일 비교 (`.xlsx`, `.xlsm`)
- 폴더 vs 폴더 비교 (하위 폴더 재귀 탐색)
- 숨김 폴더 내부 파일 포함 탐색
- Excel 임시 잠금 파일(`~$...`) 자동 제외
- 권한/잠금/손상 등 읽기 실패 파일은 중단하지 않고 오류로 리포트
- 시트 추가/삭제, 시트 상태(visible/hidden) 변경 감지
- 셀 추가/삭제/수정 감지
- 리포트는 `수정된 항목(modified)`만 출력
- 2번째 파일 보호/숨김 해제 후 1번째 파일과 다른 셀을 색상 표시한 결과 파일 생성
- 변경점 있는 시트를 탭 색상으로 표시 (노랑: 변경, 초록: 신규)
- 결과 파일의 모든 기존 셀 스타일을 흰 배경/검은 글씨로 통일 후 변경점만 색상 표시
- `_DIFF_SUMMARY` 시트에 시트/셀/이전값/변경값 상세 diff 표 제공
- `_DIFF_SUMMARY` 상세 diff는 2만 건 제한 없이 전체 표시
  - 엑셀 한 시트 한계를 넘으면 `_DIFF_SUMMARY_2`, `_DIFF_SUMMARY_3`로 자동 분할
- 숫자 비교 시 `1`과 `1.0`은 같은 값으로 처리
- 문자열 비교 시 공백/줄바꿈/NBSP/zero-width 문자 차이는 무시
- 수식 비교 시 시트명 따옴표/공백/대소문자/@ 참조 차이는 무시
- 동일한 드롭다운(데이터 유효성 list) 셀의 선택값 변경은 diff에서 제외

## 빠른 시작

권장 환경:
- macOS
- Python 3.9 이상
- 첫 실행 시 패키지 설치를 위한 인터넷 연결

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 사용법

### GUI 실행 (PySide6, 파일 2개 비교 전용)

```bash
python excel_diff_gui.py
```

맥에서는 `run_gui.command`를 더블클릭해도 실행됩니다.

터미널 실행(클론 위치 무관):

```bash
cd <클론한-폴더>/ExcelAutoDiff
./run_gui.command
```

`no such file or directory: Users/...` 오류가 나오면 경로 맨 앞 `/`가 빠졌거나,
현재 폴더가 레포 루트가 아닐 수 있습니다. 레포 루트에서 `./run_gui.command`로 실행하세요.

실행 로그는 `.logs/run_gui-YYYYMMDD-HHMMSS.log`에 저장됩니다.
기본 실행은 포그라운드 모드(더블클릭 안정성 우선)입니다.
필요 시 아래처럼 분리(detach) 실행도 가능합니다.

```bash
RUN_GUI_DETACH=1 ./run_gui.command
```

처음 실행 시 `requirements.txt` 설치가 진행됩니다.
인터넷 연결 문제로 설치에 실패하면 터미널에 실패 원인 마지막 로그가 출력되며,
상세 내용은 `.logs/run_gui-*.log`에서 확인할 수 있습니다.

GUI 흐름:
1. 기준 파일 / 비교 파일을 선택
2. 결과 파일 경로를 확인(기본 자동 생성)
3. `비교 실행` 클릭
4. 필요 시 하단 콘솔에서 `run`, `clear`, `reset`, `help` 명령 사용

## 배포

### macOS 로컬 배포 파일 생성

아래 스크립트로 `.app`, `.zip`, `.dmg`를 한 번에 생성합니다.

```bash
./scripts/build_macos.sh v1.0.0
```

생성 경로:
- `dist/ExcelAutoDiff.app`
- `dist/ExcelAutoDiff-v1.0.0-macOS.zip`
- `dist/ExcelAutoDiff-v1.0.0-macOS.dmg`

선택 사항(서명/노타라이즈):
- 환경 변수로 아래 값을 주면 빌드 스크립트가 서명/노타라이즈를 수행합니다.
  - `APPLE_SIGN_IDENTITY`
  - `APPLE_NOTARY_KEY_ID`
  - `APPLE_NOTARY_ISSUER_ID`
  - `APPLE_NOTARY_KEY_P8`

### GitHub Releases 배포

이미 포함된 워크플로우:
- `.github/workflows/release-macos.yml`

배포 방법:
1. 태그 생성/푸시 (`v1.0.0` 형식 권장)
2. GitHub Actions가 macOS 빌드 수행
3. Release Assets에 `.dmg`, `.zip` 자동 업로드

주의:
- `v*` 태그 릴리즈는 GitHub Actions에서 서명 + 노타라이즈를 수행하도록 구성되어 있습니다.
- 아래 GitHub Secrets가 없으면 태그 릴리즈 작업은 실패하도록 설정되어 있습니다.
  - `APPLE_SIGN_IDENTITY` (예: `Developer ID Application: <Name> (<TEAMID>)`)
  - `APPLE_NOTARY_KEY_ID`
  - `APPLE_NOTARY_ISSUER_ID`
  - `APPLE_NOTARY_KEY_P8` (App Store Connect API key `.p8` 전체 내용)
- 외부 배포(일반 사용자 대상) 시 Apple Developer 서명/노타라이즈를 반드시 유지하세요.

## 보안/운영 주의사항

- `report.md`, `.logs/run_gui-*.log`에는 로컬 파일 경로/셀 값 일부가 포함될 수 있으므로 외부 공유 전에 마스킹하세요.
- 외부에서 받은 신뢰되지 않은 엑셀 파일은 사내/개인 중요 데이터와 분리된 환경에서 검증 후 사용을 권장합니다.
- 배포 산출물(`.app`, `.zip`, `.dmg`)은 기본적으로 코드 서명/노타라이즈가 없으므로, 외부 공개 배포 전 서명 절차를 권장합니다.

### 1) 파일 2개 직접 비교

```bash
python excel_diff.py "/path/old_quote.xlsx" "/path/new_quote.xlsm" -o report.md
```

### 2) 폴더 2개 비교

```bash
python excel_diff.py "/path/old_folder" "/path/new_folder" -o report.md
```

### 3) JSON 리포트로 저장

```bash
python excel_diff.py "/path/old_folder" "/path/new_folder" --format json -o report.json
```

### 4) 포맷(number format) 차이도 비교

```bash
python excel_diff.py "/path/old_folder" "/path/new_folder" --include-format -o report.md
```

### 5) 2번째 파일에 차이 색상 표시 파일 생성 (추천)

```bash
python excel_diff.py "/path/first.xlsx" "/path/second.xlsm" \
  --highlight-output "/path/second_diff_highlight.xlsm"
```

- 기본 동작:
  - 2번째 파일에 보호 해제/숨김 해제를 적용
  - 2번째 파일 전체를 흰 배경/검은 글씨로 정리
  - 1번째 파일과 다른 셀을 색상으로 표시
  - `_DIFF_SUMMARY` 시트를 추가해 색상 의미와 개수를 제공
- 색상 규칙:
  - 노란색: 수정 셀
  - 초록색: 2번째 파일에만 있는 셀(추가)
  - 주황색: 1번째 파일에만 있는 셀(삭제)
  - 노란 탭: 변경된 공통 시트
  - 초록 탭: 2번째 파일 신규 시트

## 옵션

- `--output, -o`: 리포트 저장 경로 (`.md` 또는 `.json`)
- `--format {md,json}`: 리포트 출력 포맷 (기본: `md`)
- `--include-format`: 셀 값 외 `number format` 차이도 변경으로 처리
- `--max-cell-details`: 시트별 상세 변경 표시 최대 개수 (기본: `2000`)
- `--workers`: 폴더 비교 시 병렬 워커 수 (기본: 자동, 최대 8)
- `--highlight-output`: 차이 색상 표시된 결과 엑셀 파일 경로 (`xlsx/xlsm`)
- `--no-unprotect-second`: 2번째 파일 보호/숨김 해제 단계를 건너뜀

## 결과 예시

- 요약에는 `수정 파일 수`, `오류 수`, `스캔 오류 수`만 표시
- 파일 상세에는 `수정 셀`과 `시트 상태 변경`만 표시
- 추가/삭제/동일 항목은 리포트에서 제외
- 색상 파일 생성 모드에서는 엑셀 파일 자체에 차이 위치가 표시됨

# 그룹웨어 문서 PDF 다운로드 자동화

사내 그룹웨어(전자결재 시스템)에서 특정 양식(출장신청서, 출장보고서 등)의 문서를
목록에서 한 건씩 자동으로 열어 PDF로 저장해주는 Selenium 기반 자동화 도구입니다.

> ⚠️ 이 코드는 특정 그룹웨어(더존 Bizbox 계열)의 화면 구조에 맞춰 작성되었습니다.
> 다른 그룹웨어를 쓰신다면 화면 구조가 달라 selector(요소 선택자) 수정이 필요할 수 있습니다.

## 포함된 스크립트

| 파일 | 용도 |
|---|---|
| `scripts/download_documents.py` | **(권장)** 로그인부터 메뉴 이동(전자결재 > 문서함 > 기록물등록대장(모든부서)), 양식명/등록일자 검색 조건 설정, PDF 다운로드까지 전부 자동화한 통합 스크립트. 양식명을 콘솔에서 입력받으므로 출장신청서/출장보고서 등 어떤 양식이든 이 스크립트 하나로 처리 가능 |
| `scripts/retry_missing_documents.py` | `compare_missing_documents.py`가 찾아낸 미확보 문서를 문서번호로 하나씩 재검색해서 PDF로 다운로드. 검색 결과가 없는 건은 실패목록.csv에 별도 기록 |
| `scripts/compare_missing_documents.py` | 내가 가진 전체 대상 목록(xlsx)과, 다운로드 스크립트가 만든 완료목록(csv, 여러 개 가능)을 대조해서 아직 못 받은 문서를 뽑아냄 |
| `scripts/download_travel_application.py` | *(레거시)* 출장신청서 전용 개별 스크립트. `download_documents.py`로 대체됨 |
| `scripts/download_travel_report.py` | *(레거시)* 출장보고서 전용 개별 스크립트. `download_documents.py`로 대체됨 |

자세한 사용법은 [`docs/USAGE.md`](docs/USAGE.md) 를 참고해주세요. (프로그래밍을 모르셔도 따라할 수 있도록 작성했습니다.)

변경 이력은 맨 아래 [변경 이력](#변경-이력) 섹션을 참고해주세요.

## 요구 사항

- Windows 10/11 (Mac에서도 동작하나, 아래 안내는 Windows 기준입니다)
- Python 3.9 이상
- Google Chrome 브라우저

## 빠른 시작

```bash
pip install selenium pandas openpyxl
```

각 스크립트 상단의 `설정값` 부분(특히 `GROUPWARE_URL`)을 본인의 그룹웨어 주소로 수정한 뒤 실행하세요.

```bash
python scripts/download_documents.py
```

## 커스터마이징이 필요한 설정값

### `download_documents.py` (권장, 통합 스크립트)

그룹웨어 주소·아이디·비밀번호·양식명·등록일자 범위는 **모두 실행 시 콘솔에서 입력**받으므로 코드 수정이 필요 없습니다. 아래는 그룹웨어 화면 구조가 달라졌을 때만 손댈 selector들입니다.

| 변수명 | 기본값 | 설명 |
|---|---|---|
| `LOGIN_ID_SELECTOR` / `LOGIN_PW_SELECTOR` / `LOGIN_BUTTON_SELECTOR` | `input#userId` / `input#userPw` / `div.log_btn` | 로그인 폼 selector |
| `FORM_NAME_INPUT_SELECTOR` | `input#tiname` | 상세검색 > 양식명 입력창 |
| `DATE_START_HIDDEN_ID` / `DATE_END_HIDDEN_ID` | `c_startDate` / `c_endDate` | 상세검색 > 등록일자 시작/종료 hidden input id |
| `SEARCH_DETAIL_TOGGLE_SELECTOR` | `span.btn_Detail` | 상세검색 패널을 펼치는 토글 버튼 |
| `SEARCH_BUTTON_SELECTOR` | `None` | 별도 검색 버튼이 있으면 지정. `None`이면 Enter 키로 검색 시도 |
| `WAIT_SECONDS` | `15` | 요소 대기 최대 시간(초) |
| `IFRAME_SELECTOR` | `iframe[name='_content']` | 목록/검색폼이 들어있는 iframe |
| `get_rows()` 내부 selector | `div.grid-content table tbody tr` | 문서 목록 표의 각 행 |
| `row_info()` 내부 인덱스 | `cells[2]`, `cells[4]`, `cells[5]` | 등록번호/양식명/제목 컬럼 순서 |
| PDF 저장 버튼 selector | `input[value='PDF저장']` | 팝업 안 PDF 저장 버튼 |
| 메뉴 이동 대상 텍스트 | `"전자결재"` → `"문서함"` → `"[기록물등록대장(모든부서)]"` | 로그인 후 목록 화면까지 자동으로 클릭하는 메뉴 경로 (`navigate_to_document_list()`) |

실행 중 **Enter를 누르면 안전하게 중단**됩니다 (처리 중인 문서 하나는 끝까지 마친 뒤 종료).

### `retry_missing_documents.py`

| 변수명 | 기본값 | 설명 |
|---|---|---|
| `MISSING_XLSX_PATH` | `미확보목록.xlsx` (스크립트와 같은 폴더 기준) | `compare_missing_documents.py`가 만든 결과 파일 |
| `MISSING_SHEET_NAME` | `"미확보"` | 읽어올 시트 이름 |
| `DOC_NUMBER_COLUMN` | `"문서번호"` | 검색에 사용할 문서번호 컬럼명 |
| `DOC_NUM_INPUT_SELECTOR` | `input#docNum` | 상세검색 > 문서번호 입력창 |

### `download_travel_application.py`, `download_travel_report.py` (레거시)

> 이 두 스크립트는 `download_documents.py`로 기능이 통합되어 더 이상 업데이트되지 않습니다. 참고용으로만 남겨둡니다.

| 변수명 | 기본값 | 설명 |
|---|---|---|
| `GROUPWARE_URL` | `https://your-groupware-url.example.com` | **필수 수정.** 본인 그룹웨어 로그인 페이지 주소 |
| `TARGET_FORM_NAME` | `"출장신청서"` / `"출장보고서"` | 목록에서 필터링할 양식명 |
| `WAIT_SECONDS` | `15` | 요소 대기 최대 시간(초) |

`DOWNLOAD_DIR`, `LOG_CSV`는 `TARGET_FORM_NAME` 기준으로 자동 결정되므로 직접 수정할 필요는 없습니다.

### `compare_missing_documents.py`

| 변수명 | 기본값 | 설명 |
|---|---|---|
| `XLSX_PATH` | `전체목록.xlsx` (스크립트와 같은 폴더 기준) | **필수 수정.** 본인이 가진 전체 대상 목록 엑셀 파일 경로 |
| `XLSX_TITLE_COLUMN` | `"신청내역"` | 그 엑셀에서 문서제목이 들어있는 컬럼명 |
| `CSV_PATHS` | `["완료목록.csv"]` | 다운로드 스크립트가 만든 완료목록.csv 경로들 (여러 개 나열 가능) |
| `CSV_AUTO_SEARCH_ROOT` | `None` | 지정하면 이 폴더 하위의 모든 `완료목록.csv`를 자동으로 찾아 합침 (`CSV_PATHS`를 빈 리스트로 둬야 적용됨) |
| `OUTPUT_PATH` | `미확보목록.xlsx` | 결과 파일 저장 경로 |
| `SIMILARITY_THRESHOLD` | `0.85` | 이 값 이상 유사하면 "유사매칭"으로 분류 (0~1, 1에 가까울수록 엄격) |

selector 값을 어떻게 확인하고 수정하는지는 `docs/USAGE.md`의 "3-2. 화면 구조(selector) 확인" 항목에 개발자도구 사용법과 함께 자세히 안내되어 있습니다.

## 주의사항 / 면책

- 이 코드는 **본인 소속 기관의 정당한 업무 목적**(본인이 열람 권한을 가진 문서를 다운로드)으로만 사용해야 합니다.
- 그룹웨어 로그인 정보(아이디/비밀번호)는 코드에 저장하지 않으며, 실행 시 브라우저에서 직접 로그인하는 방식입니다.
- 반복적인 자동 요청이 서버에 부담을 줄 수 있으니, 짧은 기간에 대량으로 돌리기보다는 날짜 범위를 나눠서 실행하는 것을 권장합니다.
- 그룹웨어 화면 구조는 회사/버전마다 다르므로, 본인의 그룹웨어에 맞게 selector를 직접 확인하고 수정해야 합니다. (`docs/USAGE.md`의 "다른 그룹웨어에 맞게 수정하는 법" 참고)
- 이 저장소는 특정 회사 내부 시스템의 실제 주소를 포함하지 않습니다. 사용 전 본인 환경의 주소로 반드시 채워 넣어야 동작합니다.

## 변경 이력

### 2026-08-10
- **`download_documents.py` 신설(통합 스크립트)** — 로그인, 메뉴 이동(전자결재 > 문서함 > [기록물등록대장(모든부서)]), 검색조건(양식명/등록일자) 설정, PDF 다운로드까지 하나로 통합. 그룹웨어 주소/아이디/비밀번호/양식명/등록일자를 모두 콘솔 입력으로 처리
- 로그인 자동화 추가 — 아이디/비밀번호를 코드에 저장하지 않고 실행 시 콘솔에서 입력(비밀번호는 화면에 표시 안 됨). 이차인증(OTP 등) 감지 시 사람이 마무리하도록 대기
- 상세검색 패널이 기본적으로 접혀있는 문제 대응 — 토글 클릭이 안 먹힐 경우 JS로 강제 표시하는 안전장치 추가
- 실행 중 **Enter로 안전하게 중단**하는 기능 추가 (처리 중인 문서 하나는 끝까지 마친 뒤 종료)
- `retry_missing_documents.py` 신설 — `compare_missing_documents.py`가 찾아낸 미확보 문서를 문서번호로 재검색해서 다운로드, 실패 건은 실패목록.csv에 기록
  - 문서번호가 엑셀에서 float로 읽혀 `.0`이 붙는 문제 수정 (`normalize_doc_no()`)
- `compare_missing_documents.py`에 완료목록.csv 여러 개를 합쳐서 대조하는 기능 추가 (`CSV_PATHS`, `CSV_AUTO_SEARCH_ROOT`)
- 페이지네이션 로직을 화살표 버튼(`span.nex`) 방식에서 **숫자 페이지 버튼 방식**으로 변경 — 화살표 버튼의 `disabled` 판단이 부정확해 중간 페이지에서 멈추던 문제 해결
- 프로젝트 명칭에서 "매크로"라는 표현을 "자동화"로 전체 변경 (부정적 어감 방지)

### 2026-08-06 ~ 2026-08-07
- 최초 배포 — `download_travel_application.py`(출장신청서), `download_travel_report.py`(출장보고서), `compare_missing_documents.py`(미확보 대조) 3종 스크립트
- 목록이 iframe(`_content`) 안에 있는 구조 대응 (`switch_to_content_frame()`)
- PDF저장 버튼 클릭 시 오버레이에 가로채이는 문제에 JS 강제 클릭으로 대응

## 라이선스

MIT License — 자유롭게 가져다 쓰시되, 무보증(as-is)입니다.

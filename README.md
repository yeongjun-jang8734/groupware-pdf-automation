# sanai_three
bizbox alpha 그룹웨어 문서 자동 다운로드 매크로
그룹웨어 문서 PDF 자동 다운로드 매크로
사내 그룹웨어(전자결재 시스템)에서 특정 양식(출장신청서, 출장보고서 등)의 문서를
목록에서 한 건씩 자동으로 열어 PDF로 저장해주는 Selenium 기반 자동화 도구입니다.
> ⚠️ 이 코드는 특정 그룹웨어(더존 Bizbox 계열)의 화면 구조에 맞춰 작성되었습니다.
> 다른 그룹웨어를 쓰신다면 화면 구조가 달라 selector(요소 선택자) 수정이 필요할 수 있습니다.
포함된 스크립트 (3종)
파일	용도
`scripts/download_travel_application.py`	목록에서 "출장신청서" 양식 문서를 찾아 PDF로 자동 다운로드
`scripts/download_travel_report.py`	목록에서 "출장보고서" 양식 문서를 찾아 PDF로 자동 다운로드
`scripts/compare_missing_documents.py`	내가 가진 전체 대상 목록(xlsx)과, 위 매크로가 만든 완료목록(csv)을 대조해서 아직 못 받은 문서를 뽑아냄
자세한 사용법은 `docs/USAGE.md` 를 참고해주세요. (프로그래밍을 모르셔도 따라할 수 있도록 작성했습니다.)
요구 사항
Windows 10/11 (Mac에서도 동작하나, 아래 안내는 Windows 기준입니다)
Python 3.9 이상
Google Chrome 브라우저
빠른 시작
```bash
pip install selenium pandas openpyxl
```
각 스크립트 상단의 `설정값` 부분(특히 `GROUPWARE_URL`)을 본인의 그룹웨어 주소로 수정한 뒤 실행하세요.
```bash
python scripts/download_travel_application.py
```
주의사항 / 면책
이 코드는 본인 소속 기관의 정당한 업무 목적(본인이 열람 권한을 가진 문서를 다운로드)으로만 사용해야 합니다.
그룹웨어 로그인 정보(아이디/비밀번호)는 코드에 저장하지 않으며, 실행 시 브라우저에서 직접 로그인하는 방식입니다.
반복적인 자동 요청이 서버에 부담을 줄 수 있으니, 짧은 기간에 대량으로 돌리기보다는 날짜 범위를 나눠서 실행하는 것을 권장합니다.
그룹웨어 화면 구조는 회사/버전마다 다르므로, 본인의 그룹웨어에 맞게 selector를 직접 확인하고 수정해야 합니다. (`docs/USAGE.md`의 "다른 그룹웨어에 맞게 수정하는 법" 참고)
이 저장소는 특정 회사 내부 시스템의 실제 주소를 포함하지 않습니다. 사용 전 본인 환경의 주소로 반드시 채워 넣어야 동작합니다.
라이선스
MIT License — 자유롭게 가져다 쓰시되, 무보증(as-is)입니다.

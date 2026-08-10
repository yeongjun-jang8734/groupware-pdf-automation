# -*- coding: utf-8 -*-
"""
미확보 문서 재검색 및 PDF 자동 다운로드 (Selenium)

compare_missing_documents.py가 만든 '미확보목록.xlsx'의 '미확보' 시트에서
'문서번호' 값을 하나씩 꺼내 목록 화면 상세검색(문서번호)으로 검색하고,
검색된 문서를 열어 PDF로 저장합니다.

동작 흐름:
1. 그룹웨어 로그인 (콘솔에서 아이디/비밀번호 입력, 코드에 저장하지 않음)
2. 로그인 후, 사람이 직접 등록일자 범위를 "충분히 넓게" 설정해 목록 화면까지 진입 (중요: 아래 안내 참고)
3. 미확보목록.xlsx의 문서번호를 하나씩 상세검색(문서번호) 입력창에 넣어 검색
4. 검색 결과가 있으면: 문서 클릭 -> 팝업 -> PDF 저장 -> 완료목록.csv에 기록
5. 검색 결과가 없으면: 실패목록.csv에 사유와 함께 기록하고 다음 건으로 진행
6. 모든 건 처리 후 성공/실패 건수 요약 출력

★★★ 주의 ★★★
문서번호로 검색해도 화면에 걸려있는 등록일자 범위를 벗어난 문서는 안 나올 수 있습니다.
실행 전 목록 화면에서 등록일자 범위를 최대한 넓게(전체기간 등) 설정해두세요.
"""

import csv
import getpass
import os
import time

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
)

# ------------------- 설정값 -------------------
GROUPWARE_URL = "https://your-groupware-url.example.com"  # ★ 사용자의 그룹웨어 주소로 반드시 교체하세요

AUTO_LOGIN = True
LOGIN_ID_SELECTOR = (By.CSS_SELECTOR, "input#userId")
LOGIN_PW_SELECTOR = (By.CSS_SELECTOR, "input#userPw")
LOGIN_BUTTON_SELECTOR = (By.CSS_SELECTOR, "div.log_btn")

IFRAME_SELECTOR = (By.CSS_SELECTOR, "iframe[name='_content']")

# 상세검색 > 문서번호 입력창
DOC_NUM_INPUT_SELECTOR = (By.CSS_SELECTOR, "input#docNum")
# 검색 버튼이 따로 있다면 selector를 채워주세요 (예: (By.CSS_SELECTOR, "button.btnSearch"))
# None으로 두면 문서번호 입력 후 Enter 키로 검색을 시도합니다.
SEARCH_BUTTON_SELECTOR = None

WAIT_SECONDS = 15

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
MISSING_XLSX_PATH = os.path.join(SCRIPT_DIR, "미확보목록.xlsx")  # ★ compare_missing_documents.py 결과 파일
MISSING_SHEET_NAME = "미확보"
DOC_NUMBER_COLUMN = "문서번호"
TITLE_COLUMN_FOR_REPORT = "신청내역"  # 실패 목록에 참고용으로 같이 남길 제목 컬럼 (없으면 빈 값 처리)

DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads", "미확보_재다운로드_PDF")
LOG_CSV = os.path.join(DOWNLOAD_DIR, "완료목록.csv")
FAILED_CSV = os.path.join(DOWNLOAD_DIR, "실패목록.csv")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)


def get_driver():
    options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
    }
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    return driver


def load_done_list():
    done = set()
    if os.path.exists(LOG_CSV):
        with open(LOG_CSV, newline="", encoding="utf-8-sig") as f:
            for row in csv.reader(f):
                if row:
                    done.add(row[0])
    return done


def append_done(doc_no, title):
    with open(LOG_CSV, "a", newline="", encoding="utf-8-sig") as f:
        csv.writer(f).writerow([doc_no, title, time.strftime("%Y-%m-%d %H:%M:%S")])


def append_failed(doc_no, ref_title, reason):
    is_new = not os.path.exists(FAILED_CSV)
    with open(FAILED_CSV, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        if is_new:
            writer.writerow(["문서번호", "참고제목(원본xlsx)", "실패사유", "처리시각"])
        writer.writerow([doc_no, ref_title, reason, time.strftime("%Y-%m-%d %H:%M:%S")])


def switch_to_content_frame(driver):
    """목록/검색폼이 들어있는 iframe(_content) 안으로 전환한다."""
    driver.switch_to.default_content()
    WebDriverWait(driver, WAIT_SECONDS).until(
        EC.frame_to_be_available_and_switch_to_it(IFRAME_SELECTOR)
    )


def get_rows(driver):
    return driver.find_elements(By.CSS_SELECTOR, "div.grid-content table tbody tr")


def row_info(row):
    cells = row.find_elements(By.TAG_NAME, "td")
    return {
        "doc_no": cells[2].text.strip(),
        "form_name": cells[4].text.strip(),
        "title": cells[5].text.strip(),
    }


def open_popup_and_save_pdf(driver, row):
    main_handle = driver.current_window_handle
    before_handles = set(driver.window_handles)

    title_span = row.find_element(
        By.CSS_SELECTOR, "td:nth-child(6) span[onclick*='titleOnClickPudd']"
    )
    title_span.click()

    WebDriverWait(driver, WAIT_SECONDS).until(
        lambda d: len(d.window_handles) > len(before_handles)
    )
    new_handle = (set(driver.window_handles) - before_handles).pop()
    driver.switch_to.window(new_handle)

    try:
        WebDriverWait(driver, WAIT_SECONDS).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
    except TimeoutException:
        pass

    try:
        pdf_btn = WebDriverWait(driver, WAIT_SECONDS).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[value='PDF저장']"))
        )
        try:
            WebDriverWait(driver, 5).until(
                EC.invisibility_of_element_located(
                    (By.CSS_SELECTOR, "div[style*='z-index: 2002']")
                )
            )
        except TimeoutException:
            pass

        try:
            pdf_btn.click()
        except Exception:
            driver.execute_script("arguments[0].click();", pdf_btn)

        time.sleep(3)
    finally:
        driver.close()
        driver.switch_to.window(main_handle)
        switch_to_content_frame(driver)


def search_by_doc_number(driver, doc_number):
    """문서번호 입력창에 값을 넣고 검색을 실행한다."""
    print(f"[디버그] 검색창에 입력할 값: {repr(doc_number)}")

    doc_input = WebDriverWait(driver, WAIT_SECONDS).until(
        EC.presence_of_element_located(DOC_NUM_INPUT_SELECTOR)
    )
    doc_input.clear()
    doc_input.send_keys(str(doc_number))

    old_rows = get_rows(driver)
    reference_row = old_rows[0] if old_rows else None

    if SEARCH_BUTTON_SELECTOR:
        driver.find_element(*SEARCH_BUTTON_SELECTOR).click()
    else:
        doc_input.send_keys(Keys.RETURN)

    if reference_row is not None:
        try:
            WebDriverWait(driver, WAIT_SECONDS).until(EC.staleness_of(reference_row))
        except TimeoutException:
            pass

    try:
        WebDriverWait(driver, WAIT_SECONDS).until(
            EC.invisibility_of_element_located(
                (By.CSS_SELECTOR, "div.PUDD-UI-loading")
            )
        )
    except TimeoutException:
        pass

    time.sleep(1)  # 그리드 갱신 여유 시간


def normalize_doc_no(raw):
    """엑셀에서 숫자로 인식된 문서번호가 4074054.0처럼 읽히는 문제를 방지한다."""
    if isinstance(raw, float):
        if raw.is_integer():
            return str(int(raw))
        return str(raw).strip()
    text = str(raw).strip()
    # 혹시 문자열인데도 '4074054.0' 형태로 들어있는 경우까지 방어
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text


def load_missing_targets():
    """미확보목록.xlsx의 '미확보' 시트에서 (문서번호, 참고제목) 리스트를 읽어온다."""
    df = pd.read_excel(MISSING_XLSX_PATH, sheet_name=MISSING_SHEET_NAME)
    if DOC_NUMBER_COLUMN not in df.columns:
        raise ValueError(
            f"'{DOC_NUMBER_COLUMN}' 컬럼을 찾을 수 없습니다. 실제 컬럼명: {list(df.columns)}"
        )

    print(f"[디버그] '{DOC_NUMBER_COLUMN}' 컬럼의 실제 데이터 타입: {df[DOC_NUMBER_COLUMN].dtype}")

    targets = []
    for _, row in df.iterrows():
        raw = row.get(DOC_NUMBER_COLUMN)
        if pd.isna(raw) or str(raw).strip() == "":
            continue
        doc_no = normalize_doc_no(raw)
        ref_title = row.get(TITLE_COLUMN_FOR_REPORT, "")
        ref_title = "" if pd.isna(ref_title) else str(ref_title)
        targets.append((doc_no, ref_title))

    if targets:
        print(f"[디버그] 첫 번째 문서번호 예시: {repr(targets[0][0])}")

    return targets


def auto_login(driver, user_id, user_pw):
    wait = WebDriverWait(driver, WAIT_SECONDS)
    id_input = wait.until(EC.presence_of_element_located(LOGIN_ID_SELECTOR))
    pw_input = driver.find_element(*LOGIN_PW_SELECTOR)

    id_input.clear()
    id_input.send_keys(user_id)
    pw_input.clear()
    pw_input.send_keys(user_pw)

    driver.find_element(*LOGIN_BUTTON_SELECTOR).click()

    try:
        WebDriverWait(driver, WAIT_SECONDS).until(EC.staleness_of(id_input))
        print("[안내] 로그인에 성공한 것으로 보입니다.")
    except TimeoutException:
        print(
            "[안내] 로그인 처리가 자동으로 확인되지 않았습니다. "
            "이차인증(OTP 등) 화면이 떴을 수 있으니, 브라우저에서 직접 완료해주세요."
        )
        input("로그인/추가인증을 마치신 뒤 Enter를 누르세요: ")


def main():
    global GROUPWARE_URL

    entered_url = input(
        f"그룹웨어 주소를 입력하세요 (그냥 Enter 시 기본값 사용: {GROUPWARE_URL}): "
    ).strip()
    if entered_url:
        GROUPWARE_URL = entered_url

    if AUTO_LOGIN:
        user_id = input("그룹웨어 아이디를 입력하세요: ").strip()
        user_pw = getpass.getpass(
            "그룹웨어 비밀번호를 입력하세요 (입력한 문자는 화면에 표시되지 않습니다): "
        )

    print(f"[안내] 대상 파일: {MISSING_XLSX_PATH}")
    targets = load_missing_targets()
    print(f"[안내] 재검색할 문서 수: {len(targets)}건")

    driver = get_driver()
    done = load_done_list()

    driver.get(GROUPWARE_URL)

    if AUTO_LOGIN:
        auto_login(driver, user_id, user_pw)

    input(
        "목록 화면까지 진입해주세요. 등록일자 범위를 최대한 넓게(전체기간 등) 설정해야 "
        "문서번호 검색이 누락 없이 동작합니다. 준비되면 Enter를 누르세요: "
    )

    switch_to_content_frame(driver)

    success_count = 0
    fail_count = 0
    skip_count = 0

    for doc_no, ref_title in targets:
        switch_to_content_frame(driver)  # 팝업 처리 후 프레임 컨텍스트가 바뀌므로 매번 재진입

        try:
            search_by_doc_number(driver, doc_no)
        except Exception as e:
            print(f"[실패] {doc_no}: 검색 중 오류 - {e}")
            append_failed(doc_no, ref_title, f"검색 오류: {e}")
            fail_count += 1
            continue

        rows = get_rows(driver)
        if not rows:
            print(f"[실패] {doc_no}: 검색 결과 없음")
            append_failed(doc_no, ref_title, "검색 결과 없음")
            fail_count += 1
            continue

        if len(rows) > 1:
            print(f"[경고] {doc_no}: 검색 결과 {len(rows)}건 중 첫 번째 건을 처리합니다.")

        row = rows[0]
        try:
            info = row_info(row)
        except Exception:
            info = {"doc_no": doc_no, "title": ref_title}

        if info["doc_no"] in done:
            print(f"[건너뜀] {doc_no}: 이미 처리된 문서입니다.")
            skip_count += 1
            continue

        try:
            open_popup_and_save_pdf(driver, row)
            append_done(info["doc_no"], info["title"])
            done.add(info["doc_no"])
            print(f"[완료] {doc_no} - {info['title']}")
            success_count += 1
        except Exception as e:
            print(f"[실패] {doc_no}: PDF 저장 중 오류 - {e}")
            append_failed(doc_no, ref_title, f"PDF 저장 오류: {e}")
            fail_count += 1

    driver.quit()

    print("\n=== 처리 결과 ===")
    print(f"성공: {success_count}건 / 이미처리(건너뜀): {skip_count}건 / 실패: {fail_count}건 / 전체 대상: {len(targets)}건")
    print(f"완료목록: {LOG_CSV}")
    if fail_count > 0:
        print(f"실패목록: {FAILED_CSV}")


if __name__ == "__main__":
    main()

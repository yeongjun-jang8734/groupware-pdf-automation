# -*- coding: utf-8 -*-
"""
출장신청서 문서 PDF 자동 다운로드 자동화 (Selenium)

사용 전 준비:
1. pip install selenium
2. Chrome 브라우저가 PC에 설치되어 있어야 합니다.
   (selenium 4.6 이상은 크롬드라이버를 자동으로 관리해줍니다. 별도 설치 불필요)

동작 흐름:
1. 그룹웨어 로그인 페이지를 열고, 사용자가 직접 로그인 + 검색 조건 입력
2. 콘솔에서 Enter를 누르면 자동 진행 시작
3. 목록 페이지의 각 행을 순회하며 "출장신청서" 양식만 필터링
4. 행 클릭 -> 팝업 열림 -> PDF 저장 버튼 클릭 -> 다운로드
5. 팝업 닫고 다음 행으로, 페이지 끝나면 AJAX로 다음 페이지 이동
6. 처리한 문서번호는 CSV에 기록해서 재실행 시 중복 방지
"""

import csv
import getpass
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementNotInteractableException,
    ElementClickInterceptedException,
)

# ------------------- 설정값 -------------------
GROUPWARE_URL = "https://your-groupware-url.example.com"  # ★ 사용자의 그룹웨어 주소로 반드시 교체하세요

# 로그인 자동화 여부. False로 두면 이전처럼 브라우저에서 직접 로그인하고 Enter를 눌러 진행합니다.
AUTO_LOGIN = True

# 로그인 폼 selector (아이디/비밀번호를 코드에 저장하지 않고, 실행 시 콘솔에서 입력받아 채워 넣습니다)
LOGIN_ID_SELECTOR = (By.CSS_SELECTOR, "input#userId")
LOGIN_PW_SELECTOR = (By.CSS_SELECTOR, "input#userPw")
LOGIN_BUTTON_SELECTOR = (By.CSS_SELECTOR, "div.log_btn")

TARGET_FORM_NAME = "출장보고서"  # 필터링할 양식명
WAIT_SECONDS = 15  # 요소 대기 최대 시간

DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads", f"{TARGET_FORM_NAME}_PDF")
LOG_CSV = os.path.join(DOWNLOAD_DIR, "완료목록.csv")
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
    if LOG_CSV and os.path.exists(LOG_CSV):
        with open(LOG_CSV, newline="", encoding="utf-8-sig") as f:
            for row in csv.reader(f):
                if row:
                    done.add(row[0])
    return done


def append_done(doc_no, title):
    with open(LOG_CSV, "a", newline="", encoding="utf-8-sig") as f:
        csv.writer(f).writerow([doc_no, title, time.strftime("%Y-%m-%d %H:%M:%S")])


def wait_click(driver, by, selector, timeout=WAIT_SECONDS):
    el = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((by, selector))
    )
    el.click()
    return el


IFRAME_SELECTOR = (By.CSS_SELECTOR, "iframe[name='_content']")


def switch_to_content_frame(driver):
    """목록이 들어있는 iframe(_content) 안으로 전환한다."""
    driver.switch_to.default_content()
    WebDriverWait(driver, WAIT_SECONDS).until(
        EC.frame_to_be_available_and_switch_to_it(IFRAME_SELECTOR)
    )


def get_rows(driver):
    return driver.find_elements(By.CSS_SELECTOR, "div.grid-content table tbody tr")


def row_info(row):
    cells = row.find_elements(By.TAG_NAME, "td")
    return {
        "doc_no": cells[2].text.strip(),      # 등록번호 (예: 경영기획부-24)
        "form_name": cells[4].text.strip(),   # 양식명 (예: 출장신청서)
        "title": cells[5].text.strip(),       # 제목
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

    # 팝업 문서 로딩이 끝날 때까지 대기
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
        # 로딩/딤 처리 오버레이가 사라질 때까지 잠시 대기
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
            # 여전히 다른 요소가 클릭을 가로챈다면 JS로 강제 클릭
            driver.execute_script("arguments[0].click();", pdf_btn)

        # 다운로드 완료 대기 (간단 버전: 고정 대기)
        # 더 견고하게 하려면 DOWNLOAD_DIR을 폴링해서 .crdownload가 사라질 때까지 대기하는 함수로 교체 권장
        time.sleep(3)
    finally:
        driver.close()
        driver.switch_to.window(main_handle)
        # 창 전환 시 iframe 컨텍스트가 초기화되므로 다시 진입
        switch_to_content_frame(driver)


def get_current_page_number(driver):
    """현재 활성화된 페이지 번호(<li class="on">)를 읽어온다."""
    on_li = driver.find_element(By.CSS_SELECTOR, "div.paging ol li.on")
    text = on_li.text.strip()
    return int(text)


def go_next_page(driver):
    try:
        current = get_current_page_number(driver)
    except (NoSuchElementException, ValueError) as e:
        print(f"[경고] 현재 페이지 번호를 읽지 못했습니다: {e}")
        return False

    target = current + 1
    print(f"[디버그] 현재 페이지: {current} -> 목표 페이지: {target}")

    try:
        target_link = driver.find_element(
            By.XPATH,
            f"//div[contains(@class,'paging')]//ol/li/a[normalize-space(text())='{target}']",
        )
    except NoSuchElementException:
        # 목표 번호가 화면에 안 보이면(페이지 그룹이 넘어간 경우) '다음' 화살표로 창을 한 번 밀고 재시도
        # 단, 진짜 마지막 페이지라면 화살표 자체가 비활성 상태라 클릭이 안 될 수 있음 (정상 상황)
        try:
            driver.find_element(By.CSS_SELECTOR, "div.paging span.nex a").click()
            time.sleep(1)
            target_link = driver.find_element(
                By.XPATH,
                f"//div[contains(@class,'paging')]//ol/li/a[normalize-space(text())='{target}']",
            )
        except (NoSuchElementException, ElementNotInteractableException, ElementClickInterceptedException):
            print(f"[안내] {target}페이지가 없습니다. 마지막 페이지로 판단하고 종료합니다.")
            return False

    old_rows = get_rows(driver)
    reference_row = old_rows[0] if old_rows else None

    target_link.click()

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

    # 실제로 목표 페이지로 이동했는지 확인
    try:
        new_current = get_current_page_number(driver)
        print(f"[디버그] 이동 후 현재 페이지: {new_current}")
        if new_current != target:
            print(f"[경고] 목표 페이지({target})와 실제 페이지({new_current})가 다릅니다.")
    except (NoSuchElementException, ValueError):
        pass

    return True


def auto_login(driver):
    """로그인 폼에 아이디/비밀번호를 채워 넣고 로그인 버튼을 클릭한다.
    아이디/비밀번호는 코드나 파일에 저장하지 않고, 실행할 때마다 콘솔에서 직접 입력받는다.
    이차인증(OTP, 기기등록 등)이 있는 조직이라면 로그인 버튼 클릭 후 추가 인증 화면이
    뜰 수 있으므로, 로그인 성공 여부를 확인해서 실패 시 사람이 마무리하도록 안내한다.
    """
    wait = WebDriverWait(driver, WAIT_SECONDS)

    id_input = wait.until(EC.presence_of_element_located(LOGIN_ID_SELECTOR))
    pw_input = driver.find_element(*LOGIN_PW_SELECTOR)

    user_id = input("그룹웨어 아이디를 입력하세요: ").strip()
    user_pw = getpass.getpass("그룹웨어 비밀번호를 입력하세요 (입력한 문자는 화면에 표시되지 않습니다): ")

    id_input.clear()
    id_input.send_keys(user_id)
    pw_input.clear()
    pw_input.send_keys(user_pw)

    driver.find_element(*LOGIN_BUTTON_SELECTOR).click()

    # 로그인 성공 시 로그인 폼이 있던 페이지 자체가 바뀌므로, 기존 아이디 입력창이 사라질 때까지 대기
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
    driver = get_driver()
    done = load_done_list()

    driver.get(GROUPWARE_URL)

    if AUTO_LOGIN:
        auto_login(driver)
        input(
            "등록일자/부서 등 검색 조건까지 입력해 검색 버튼을 누른 뒤 "
            "여기로 돌아와 Enter를 누르세요: "
        )
    else:
        input(
            "브라우저에서 로그인을 완료하고, 등록일자/부서 등 검색 조건까지 입력해 "
            "검색 버튼을 누른 뒤 여기로 돌아와 Enter를 누르세요: "
        )

    print(f"[안내] 다운로드 폴더: {DOWNLOAD_DIR}")
    switch_to_content_frame(driver)

    page = 1
    while True:
        print(f"--- {page} 페이지 처리 중 ---")
        rows = get_rows(driver)
        print(f"[디버그] 찾은 행 개수: {len(rows)}")
        skipped_form_mismatch = 0
        if len(rows) == 0:
            # 원인 파악용: 현재 페이지에 grid-content가 있는지, table이 몇 개인지 확인
            grids = driver.find_elements(By.CSS_SELECTOR, "div.grid-content")
            tables = driver.find_elements(By.TAG_NAME, "table")
            print(f"[디버그] div.grid-content 개수: {len(grids)}, table 태그 개수: {len(tables)}")
            for i, t in enumerate(tables):
                print(f"[디버그] table[{i}] class={t.get_attribute('class')!r} id={t.get_attribute('id')!r}")
        for row in rows:
            try:
                info = row_info(row)
            except Exception:
                continue  # 헤더 행 등 td가 부족한 행은 건너뜀

            if info["form_name"] != TARGET_FORM_NAME:
                skipped_form_mismatch += 1
                if skipped_form_mismatch <= 3:
                    print(f"[디버그] 양식명 불일치로 건너뜀: '{info['form_name']}' (기대값: '{TARGET_FORM_NAME}')")
                continue
            if info["doc_no"] in done:
                print(f"이미 처리됨, 건너뜀: {info['doc_no']}")
                continue
            try:
                open_popup_and_save_pdf(driver, row)
                append_done(info["doc_no"], info["title"])
                done.add(info["doc_no"])
                print(f"완료: {info['doc_no']} - {info['title']}")
            except TimeoutException:
                print(f"실패(타임아웃): {info['doc_no']} - {info['title']}")
            except Exception as e:
                print(f"실패: {info['doc_no']} - {e}")
            time.sleep(1)  # 서버 부하 방지용 딜레이

        if not go_next_page(driver):
            print("마지막 페이지입니다. 종료합니다.")
            break
        page += 1

    driver.quit()
    print(f"완료된 목록은 {LOG_CSV} 에서 확인할 수 있습니다.")


if __name__ == "__main__":
    main()

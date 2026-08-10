# -*- coding: utf-8 -*-
"""
그룹웨어 문서 PDF 자동 다운로드 (통합판, Selenium)

download_travel_application.py / download_travel_report.py를 하나로 합친 버전입니다.
양식명(출장신청서/출장보고서 등)과 등록일자 범위를 콘솔에서 입력받아
검색 조건까지 자동으로 설정한 뒤 목록을 순회하며 PDF를 다운로드합니다.

동작 흐름:
1. 콘솔에서 그룹웨어 주소 입력 (직접 타이핑)
2. 콘솔에서 아이디/비밀번호 입력
   - 둘 다 입력하면: 자동 로그인
   - 비워두고 Enter: 브라우저에서 직접 로그인하도록 대기
3. 콘솔에서 양식명(필수), 등록일자 시작/종료(선택) 입력
4. 로그인 후 목록 화면에서 위 조건을 자동으로 채우고 검색 실행
   (다른 조건을 더 조정하고 싶다면 이 시점에 직접 조정 후 Enter)
5. 목록의 각 행을 순회하며 양식명이 일치하는 문서만 필터링
6. 행 클릭 -> 팝업 열림 -> PDF 저장 버튼 클릭 -> 다운로드
7. 팝업 닫고 다음 행으로, 페이지 끝나면 다음 페이지(숫자 버튼) 이동
8. 처리한 문서번호는 CSV에 기록해서 재실행 시 중복 방지
"""

import csv
import getpass
import os
import threading
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementNotInteractableException,
    ElementClickInterceptedException,
)

# ------------------- 설정값 -------------------
GROUPWARE_URL = ""  # 콘솔에서 입력받으므로 기본값은 비워둠

# 로그인 폼 selector (아이디/비밀번호를 코드에 저장하지 않고, 실행 시 콘솔에서 입력받아 채워 넣습니다)
LOGIN_ID_SELECTOR = (By.CSS_SELECTOR, "input#userId")
LOGIN_PW_SELECTOR = (By.CSS_SELECTOR, "input#userPw")
LOGIN_BUTTON_SELECTOR = (By.CSS_SELECTOR, "div.log_btn")

# 상세검색 > 양식명 입력창
FORM_NAME_INPUT_SELECTOR = (By.CSS_SELECTOR, "input#tiname")

# 상세검색 > 등록일자 시작/종료 (실제 폼에 제출되는 hidden input의 id)
DATE_START_HIDDEN_ID = "c_startDate"
DATE_END_HIDDEN_ID = "c_endDate"

# 검색 버튼이 따로 있다면 selector를 채워주세요 (예: (By.CSS_SELECTOR, "button.btnSearch"))
# None으로 두면 양식명 입력 후 Enter 키로 검색을 시도합니다.
SEARCH_BUTTON_SELECTOR = None

WAIT_SECONDS = 15  # 요소 대기 최대 시간

# 다운로드 폴더/로그 경로는 콘솔에서 양식명을 입력받은 뒤 main()에서 결정됩니다.
TARGET_FORM_NAME = None
DOWNLOAD_DIR = None
LOG_CSV = None


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


def set_date_field(driver, hidden_id, date_obj):
    """등록일자 hidden input과, 화면에 보이는 readonly 표시용 input을 함께 갱신한다."""
    weekdays = ["월", "화", "수", "목", "금", "토", "일"]
    date_str = date_obj.strftime("%Y-%m-%d")
    display_str = f"{date_str}({weekdays[date_obj.weekday()]})"

    try:
        hidden_el = driver.find_element(By.CSS_SELECTOR, f"input#{hidden_id}")
    except NoSuchElementException:
        print(f"[경고] '{hidden_id}' 날짜 필드를 찾지 못했습니다. selector를 확인해주세요.")
        return

    driver.execute_script("arguments[0].value = arguments[1];", hidden_el, date_str)
    driver.execute_script(
        "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));", hidden_el
    )

    # 같은 영역의 화면 표시용(readonly) input도 함께 갱신 시도
    try:
        display_el = driver.find_element(
            By.XPATH, f"//input[@id='{hidden_id}']/following-sibling::input[1]"
        )
        driver.execute_script("arguments[0].value = arguments[1];", display_el, display_str)
    except NoSuchElementException:
        pass


SEARCH_DETAIL_TOGGLE_SELECTOR = (By.CSS_SELECTOR, "span.btn_Detail")


def ensure_search_detail_open(driver):
    """상세검색(양식명/등록일자 등이 있는) 패널이 접혀있으면 펼친다."""
    panels = driver.find_elements(By.CSS_SELECTOR, "div.SearchDetail")
    print(f"[디버그] div.SearchDetail 개수: {len(panels)}")
    for i, p in enumerate(panels):
        print(f"[디버그] panel[{i}] displayed={p.is_displayed()} style={p.get_attribute('style')!r} class={p.get_attribute('class')!r}")

    if panels and panels[0].is_displayed():
        print("[디버그] 이미 열려있는 것으로 판단, 토글 클릭 생략")
        return

    toggles = driver.find_elements(*SEARCH_DETAIL_TOGGLE_SELECTOR)
    print(f"[디버그] span.btn_Detail 개수: {len(toggles)}")
    for i, t in enumerate(toggles):
        print(f"[디버그] toggle[{i}] displayed={t.is_displayed()} class={t.get_attribute('class')!r}")

    if not toggles:
        print("[경고] 상세검색 토글(span.btn_Detail)을 찾지 못했습니다. selector 확인이 필요합니다.")
        return

    print("[안내] 상세검색 패널이 접혀있어 펼칩니다.")
    toggle = toggles[0]
    try:
        toggle.click()
    except (ElementNotInteractableException, ElementClickInterceptedException):
        driver.execute_script("arguments[0].click();", toggle)
    time.sleep(1)

    panels_after = driver.find_elements(By.CSS_SELECTOR, "div.SearchDetail")
    for i, p in enumerate(panels_after):
        print(f"[디버그] 클릭 후 panel[{i}] displayed={p.is_displayed()} style={p.get_attribute('style')!r} class={p.get_attribute('class')!r}")

    # 클릭으로 안 열렸다면, 패널 스타일을 JS로 직접 강제 표시
    if panels_after and not panels_after[0].is_displayed():
        print("[안내] 클릭으로 안 열려 JS로 패널을 강제 표시합니다.")
        driver.execute_script("arguments[0].style.display = 'block';", panels_after[0])
        time.sleep(0.5)
        print(f"[디버그] 강제 표시 후 displayed={panels_after[0].is_displayed()}")


def apply_search_conditions(driver, form_name, date_start, date_end):
    """양식명/등록일자 조건을 검색폼에 채워 넣고 검색을 실행한다."""
    ensure_search_detail_open(driver)

    form_input = WebDriverWait(driver, WAIT_SECONDS).until(
        EC.visibility_of_element_located(FORM_NAME_INPUT_SELECTOR)
    )
    try:
        form_input.clear()
        form_input.send_keys(form_name)
    except ElementNotInteractableException:
        print("[경고] 양식명 입력창이 여전히 안 보여 JS로 강제 입력합니다.")
        driver.execute_script("arguments[0].value = arguments[1];", form_input, form_name)

    if date_start:
        set_date_field(driver, DATE_START_HIDDEN_ID, date_start)
    if date_end:
        set_date_field(driver, DATE_END_HIDDEN_ID, date_end)

    old_rows = get_rows(driver)
    reference_row = old_rows[0] if old_rows else None

    if SEARCH_BUTTON_SELECTOR:
        driver.find_element(*SEARCH_BUTTON_SELECTOR).click()
    else:
        try:
            form_input.send_keys(Keys.RETURN)
        except ElementNotInteractableException:
            print("[경고] Enter 키 전송이 안 되어 폼을 JS로 직접 제출합니다.")
            driver.execute_script(
                "arguments[0].form && arguments[0].form.submit && arguments[0].form.submit();",
                form_input,
            )

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


def parse_date_input(prompt_text):
    """콘솔에서 YYYY-MM-DD 형식 날짜를 입력받는다. 빈 입력이면 None 반환(조건 미설정)."""
    while True:
        text = input(prompt_text).strip()
        if not text:
            return None
        try:
            return datetime.strptime(text, "%Y-%m-%d")
        except ValueError:
            print("[안내] 날짜 형식이 올바르지 않습니다. 예: 2026-01-01")


def click_menu_by_text(driver, text, timeout=WAIT_SECONDS):
    """화면에 보이는 텍스트로 메뉴/트리 항목을 찾아 클릭한다.
    id는 계정/권한에 따라 달라질 수 있어 텍스트 기준으로 찾는 것이 더 안전하다.
    화면에 안 보여 일반 클릭이 안 되면 JS로 강제 클릭한다.
    """
    xpath = f"//*[self::div or self::a or self::span][normalize-space(text())='{text}']"
    el = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.XPATH, xpath))
    )
    try:
        el.click()
    except (ElementNotInteractableException, ElementClickInterceptedException):
        driver.execute_script("arguments[0].click();", el)


def navigate_to_document_list(driver):
    """로그인 직후 화면에서 전자결재 > 문서함 > 기록물등록대장까지 자동으로 이동한다."""
    driver.switch_to.default_content()

    print("[안내] '전자결재' 메뉴로 이동합니다.")
    click_menu_by_text(driver, "전자결재")
    time.sleep(2)  # 좌측 트리 로딩 대기

    print("[안내] '문서함'을 엽니다.")
    try:
        click_menu_by_text(driver, "문서함")
        time.sleep(1)
    except TimeoutException:
        print("[안내] '문서함' 항목을 찾지 못했습니다. 이미 펼쳐진 상태일 수 있어 계속 진행합니다.")

    print("[안내] '기록물등록대장(모든부서)'를 클릭합니다.")
    click_menu_by_text(driver, "[기록물등록대장(모든부서)]")
    time.sleep(2)  # 우측 문서 목록(iframe) 로딩 대기


STOP_EVENT = threading.Event()


def start_stop_listener():
    """백그라운드에서 Enter 입력을 감시하다가, 눌리면 STOP_EVENT를 세운다.
    현재 처리 중인 문서 하나는 끝까지 마친 뒤, 다음 문서로 넘어가기 전에 안전하게 멈춘다.
    """
    def _listener():
        try:
            input()
        except Exception:
            return
        STOP_EVENT.set()
        print("\n[안내] 중단 요청을 받았습니다. 현재 처리 중인 문서까지 마치고 안전하게 종료합니다...")

    t = threading.Thread(target=_listener, daemon=True)
    t.start()


def auto_login(driver, user_id, user_pw):
    """로그인 폼에 아이디/비밀번호를 채워 넣고 로그인 버튼을 클릭한다.
    아이디/비밀번호는 코드나 파일에 저장하지 않고, main()에서 콘솔로 입력받은 값을 그대로 전달받는다.
    이차인증(OTP, 기기등록 등)이 있는 조직이라면 로그인 버튼 클릭 후 추가 인증 화면이
    뜰 수 있으므로, 로그인 성공 여부를 확인해서 실패 시 사람이 마무리하도록 안내한다.
    """
    wait = WebDriverWait(driver, WAIT_SECONDS)

    id_input = wait.until(EC.presence_of_element_located(LOGIN_ID_SELECTOR))
    pw_input = driver.find_element(*LOGIN_PW_SELECTOR)

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
    global GROUPWARE_URL, TARGET_FORM_NAME, DOWNLOAD_DIR, LOG_CSV

    # 1) 그룹웨어 주소 입력
    GROUPWARE_URL = input("그룹웨어 주소를 입력하세요: ").strip()
    while not GROUPWARE_URL:
        GROUPWARE_URL = input("주소가 비어있습니다. 다시 입력해주세요: ").strip()

    # 2) 로그인 정보 입력 (둘 다 비워두면 브라우저에서 직접 로그인)
    user_id = input("그룹웨어 아이디를 입력하세요 (직접 로그인하려면 비워두고 Enter): ").strip()
    user_pw = ""
    if user_id:
        user_pw = getpass.getpass(
            "그룹웨어 비밀번호를 입력하세요 (입력한 문자는 화면에 표시되지 않습니다): "
        )

    # 3) 검색 조건 입력
    TARGET_FORM_NAME = input("양식명을 입력하세요 (예: 출장신청서, 출장보고서): ").strip()
    while not TARGET_FORM_NAME:
        TARGET_FORM_NAME = input("양식명은 필수입니다. 다시 입력해주세요: ").strip()

    print("등록일자 범위를 입력하세요. 형식: YYYY-MM-DD (조건 없이 진행하려면 비워두고 Enter)")
    date_start = parse_date_input("  시작일: ")
    date_end = parse_date_input("  종료일: ")

    # 4) 양식명 기준으로 다운로드 폴더/로그 경로 결정
    DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads", f"{TARGET_FORM_NAME}_PDF")
    LOG_CSV = os.path.join(DOWNLOAD_DIR, "완료목록.csv")
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    driver = get_driver()
    done = load_done_list()

    driver.get(GROUPWARE_URL)

    if user_id and user_pw:
        auto_login(driver, user_id, user_pw)
    else:
        input("브라우저에서 직접 로그인해주세요. 완료 후 Enter를 누르세요: ")

    navigate_to_document_list(driver)

    switch_to_content_frame(driver)

    print(f"[안내] 검색 조건 적용 중... 양식명='{TARGET_FORM_NAME}', 시작일={date_start}, 종료일={date_end}")
    apply_search_conditions(driver, TARGET_FORM_NAME, date_start, date_end)

    input(
        "검색 조건을 자동으로 설정하고 검색했습니다. 화면을 확인해서 "
        "부서 등 다른 조건을 추가로 조정하고 싶다면 지금 하시고, "
        "준비되면 여기로 돌아와 Enter를 누르세요: "
    )

    print(f"[안내] 다운로드 폴더: {DOWNLOAD_DIR}")
    switch_to_content_frame(driver)

    print("[안내] 지금부터 자동으로 문서를 처리합니다. 중단하려면 아무 때나 Enter를 누르세요.")
    start_stop_listener()

    page = 1
    stopped = False
    while True:
        if STOP_EVENT.is_set():
            stopped = True
            break

        print(f"--- {page} 페이지 처리 중 ---")
        rows = get_rows(driver)
        print(f"[디버그] 찾은 행 개수: {len(rows)}")
        skipped_form_mismatch = 0
        if len(rows) == 0:
            grids = driver.find_elements(By.CSS_SELECTOR, "div.grid-content")
            tables = driver.find_elements(By.TAG_NAME, "table")
            print(f"[디버그] div.grid-content 개수: {len(grids)}, table 태그 개수: {len(tables)}")
            for i, t in enumerate(tables):
                print(f"[디버그] table[{i}] class={t.get_attribute('class')!r} id={t.get_attribute('id')!r}")
        for row in rows:
            if STOP_EVENT.is_set():
                stopped = True
                break

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

        if stopped:
            break

        if not go_next_page(driver):
            print("마지막 페이지입니다. 종료합니다.")
            break
        page += 1

    driver.quit()

    if stopped:
        print("[안내] 사용자 요청으로 중단되었습니다. 지금까지 처리된 내용은 완료목록.csv에 저장되어 있습니다.")
    print(f"완료된 목록은 {LOG_CSV} 에서 확인할 수 있습니다.")


if __name__ == "__main__":
    main()

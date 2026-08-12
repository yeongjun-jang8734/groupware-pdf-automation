# -*- coding: utf-8 -*-
"""
그룹웨어 문서 PDF 자동 다운로드 (연도별 반복판, Selenium)

콘솔에서 등록일자(기간)만 입력받고, 양식명을 비롯한 나머지 상세검색 조건은
사람이 그룹웨어 화면에서 직접 설정합니다. 검색 결과로 나온 행은 필터링 없이
전부 다운로드하며, 한 기간(보통 연도 단위) 처리가 끝나면 파일에 다운로드 순서대로
일련번호를 붙이고, 이어서 다음 기간을 설정해 반복할 수 있습니다.

동작 흐름:
1. 콘솔에서 그룹웨어 주소, 아이디/비밀번호(선택) 입력
2. 로그인 후 전자결재 > 문서함 > 기록물등록대장(모든부서)까지 자동 이동
3. [반복] 콘솔에서 등록일자 시작/종료 입력
   -> 등록일자만 자동으로 채움 (상세검색 패널을 억지로 열지 않음, JS로 hidden 값만 설정)
   -> 양식명 등 나머지 조건은 브라우저에서 직접 설정 + 검색 버튼도 직접 클릭
   -> 준비되면 Enter -> 자동으로 목록을 순회하며 PDF 다운로드
   -> 다운로드 폴더는 등록일자 연도를 이름에 포함해서 자동 생성 (예: 2024_PDF)
   -> 이 기간의 마지막 페이지까지 끝나면, 폴더 내 파일에 다운로드 순서대로
      "1. 원래제목.pdf" 형태로 일련번호를 붙임
   -> 이어서 다른 기간을 처리할지 물어봄 (y: 계속 / n: 종료)
4. 실행 중 Enter를 누르면 안전하게 중단 가능 (처리 중인 문서 하나는 끝까지 마침)
"""

import csv
import difflib
import getpass
import os
import re
import time
from datetime import datetime

try:
    import msvcrt  # Windows 전용
except ImportError:
    import select
    import sys as _sys

    def _key_pressed():
        return select.select([_sys.stdin], [], [], 0)[0] != []

    def _read_key():
        return _sys.stdin.read(1).encode()
else:
    def _key_pressed():
        return msvcrt.kbhit()

    def _read_key():
        return msvcrt.getch()

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
GROUPWARE_URL = "https://gw.volunteer.seoul.kr/gw/bizbox.do"  # 콘솔에서 비워두면 이 기본값 사용

# 로그인 폼 selector (아이디/비밀번호를 코드에 저장하지 않고, 실행 시 콘솔에서 입력받아 채워 넣습니다)
LOGIN_ID_SELECTOR = (By.CSS_SELECTOR, "input#userId")
LOGIN_PW_SELECTOR = (By.CSS_SELECTOR, "input#userPw")
LOGIN_BUTTON_SELECTOR = (By.CSS_SELECTOR, "div.log_btn")

# 상세검색 > 등록일자 시작/종료 (실제 폼에 제출되는 hidden input의 id)
DATE_START_HIDDEN_ID = "c_startDate"
DATE_END_HIDDEN_ID = "c_endDate"

WAIT_SECONDS = 15  # 요소 대기 최대 시간

# PDF 저장 버튼 클릭 후, 다운로드 완료를 확인 못하면 버튼을 다시 눌러 재시도하는 횟수/시간
PDF_SAVE_MAX_RETRIES = 10
PDF_SAVE_TIMEOUT_PER_TRY = 30  # 한 번 시도당 최대 대기 시간(초)

# 다운로드 폴더/로그 경로는 기간(연도)마다 main()의 반복 루프 안에서 결정됩니다.
DOWNLOAD_DIR = None
LOG_CSV = None


def get_driver():
    """브라우저를 기본 다운로드 경로로 띄운다. 실제 연도별 경로는 이후
    execute_cdp_cmd로 매번 재설정한다 (브라우저를 새로 띄우지 않고 경로만 바꿈)."""
    default_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": default_dir,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
    }
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    return driver


def set_download_dir(driver, folder):
    """이미 열려있는 브라우저의 다운로드 경로를 CDP 명령으로 재설정한다."""
    os.makedirs(folder, exist_ok=True)
    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": folder},
    )


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


def wait_for_new_download(before_files, timeout=30, stable_checks=2, poll_interval=0.5):
    """DOWNLOAD_DIR에 새 파일이 나타나 크기가 안정될 때까지(다운로드가 끝날 때까지) 기다린다."""
    deadline = time.time() + timeout
    stable_count = 0
    last_size = -1

    while time.time() < deadline:
        try:
            current_files = set(os.listdir(DOWNLOAD_DIR))
        except FileNotFoundError:
            current_files = set()

        new_files = current_files - before_files
        in_progress = [f for f in new_files if f.endswith(".crdownload") or f.endswith(".tmp")]
        completed_candidates = [f for f in new_files if not (f.endswith(".crdownload") or f.endswith(".tmp"))]

        if in_progress:
            stable_count = 0
            time.sleep(poll_interval)
            continue

        if completed_candidates:
            target_file = max(
                completed_candidates,
                key=lambda f: os.path.getmtime(os.path.join(DOWNLOAD_DIR, f)),
            )
            try:
                size = os.path.getsize(os.path.join(DOWNLOAD_DIR, target_file))
            except OSError:
                size = -1

            if size == last_size and size > 0:
                stable_count += 1
                if stable_count >= stable_checks:
                    return True
            else:
                stable_count = 0
                last_size = size

        time.sleep(poll_interval)

    return False


def open_popup_and_save_pdf(driver, row, info=None):
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
            os.makedirs(DOWNLOAD_DIR, exist_ok=True)
            before_files = set(os.listdir(DOWNLOAD_DIR))
        except FileNotFoundError:
            before_files = set()

        completed = False
        for attempt in range(1, PDF_SAVE_MAX_RETRIES + 1):
            try:
                pdf_btn.click()
            except Exception:
                driver.execute_script("arguments[0].click();", pdf_btn)

            completed = wait_for_new_download(before_files, timeout=PDF_SAVE_TIMEOUT_PER_TRY)
            if completed:
                break

            print(f"[경고] {attempt}번째 시도에서 다운로드 완료를 확인하지 못했습니다.")
            if attempt < PDF_SAVE_MAX_RETRIES:
                print("[안내] PDF저장 버튼을 다시 눌러 재시도합니다...")
                time.sleep(2)
                try:
                    pdf_btn = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "input[value='PDF저장']"))
                    )
                except TimeoutException:
                    print("[경고] 재시도 중 PDF저장 버튼을 다시 찾지 못했습니다. 재시도를 중단합니다.")
                    break

        if not completed:
            doc_label = f"{info['doc_no']} - {info['title']}" if info else "(문서 정보 없음)"
            print(
                f"[경고] {PDF_SAVE_MAX_RETRIES}번 시도했지만 다운로드 완료를 확인하지 못했습니다. "
                f"파일을 직접 확인해주세요. 실패한 문서: {doc_label}"
            )
    except TimeoutException:
        try:
            popup_url = driver.current_url
            popup_title = driver.title
            body_text = driver.find_element(By.TAG_NAME, "body").text[:200]
        except Exception:
            popup_url = popup_title = body_text = "(진단 정보 수집 실패)"
        print(
            "[진단] PDF저장 버튼을 찾지 못했습니다. 팝업이 정상 문서 화면이 아닐 수 있습니다.\n"
            f"       팝업 URL: {popup_url}\n"
            f"       팝업 제목: {popup_title}\n"
            f"       팝업 본문 일부: {body_text!r}"
        )
        raise
    finally:
        driver.close()
        driver.switch_to.window(main_handle)
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

    try:
        new_current = get_current_page_number(driver)
        print(f"[디버그] 이동 후 현재 페이지: {new_current}")
        if new_current != target:
            print(f"[경고] 목표 페이지({target})와 실제 페이지({new_current})가 다릅니다.")
    except (NoSuchElementException, ValueError):
        pass

    return True


def set_date_field(driver, hidden_id, date_obj):
    """등록일자 hidden input과, 화면에 보이는 readonly 표시용 input을 함께 갱신한다.
    JS로 값을 직접 넣기 때문에, 상세검색 패널이 화면에 보이지 않는 상태여도 동작한다.
    """
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

    try:
        display_el = driver.find_element(
            By.XPATH, f"//input[@id='{hidden_id}']/following-sibling::input[1]"
        )
        driver.execute_script("arguments[0].value = arguments[1];", display_el, display_str)
    except NoSuchElementException:
        pass


def set_search_dates(driver, date_start, date_end):
    """등록일자만 채워 넣는다. 양식명 등 다른 조건과 검색 실행은 사람이 직접 한다."""
    if date_start:
        set_date_field(driver, DATE_START_HIDDEN_ID, date_start)
    if date_end:
        set_date_field(driver, DATE_END_HIDDEN_ID, date_end)


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


def compute_period_label(date_start, date_end):
    """다운로드 폴더 이름에 쓸 라벨을 만든다. 보통 시작일 기준 연도를 사용한다."""
    if date_start:
        return str(date_start.year)
    if date_end:
        return str(date_end.year)
    return "전체기간"


def dump_similar_elements(driver, keyword, max_items=15):
    """정확히 일치하는 요소를 못 찾았을 때, 키워드가 포함된 요소들을 콘솔에 보여줘 원인 파악을 돕는다."""
    xpath = f"//*[self::div or self::a or self::span][contains(normalize-space(.), '{keyword}')]"
    elements = driver.find_elements(By.XPATH, xpath)
    print(f"[진단] '{keyword}' 텍스트를 포함하는 요소 {len(elements)}개 발견 (상위 {max_items}개만 표시):")
    for i, el in enumerate(elements[:max_items]):
        try:
            tag = el.tag_name
            text = el.text.strip().replace("\n", " ")
            cls = el.get_attribute("class")
            print(f"  [{i}] <{tag}> class={cls!r} text={text!r}")
        except Exception:
            continue


def click_menu_by_text(driver, text, timeout=WAIT_SECONDS, exact=True):
    """화면에 보이는 텍스트로 메뉴/트리 항목을 찾아 클릭한다."""
    if exact:
        xpath = f"//*[self::div or self::a or self::span][normalize-space(text())='{text}']"
    else:
        xpath = f"//*[self::div or self::a or self::span][contains(normalize-space(text()), '{text}')]"

    try:
        el = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
    except TimeoutException:
        if exact:
            print(f"[안내] '{text}' 정확히 일치하는 요소를 못 찾아, 포함 검색으로 재시도합니다.")
            return click_menu_by_text(driver, text, timeout=5, exact=False)
        keyword = text.strip("[]() ")[:6] or text
        dump_similar_elements(driver, keyword)
        raise

    try:
        el.click()
    except (ElementNotInteractableException, ElementClickInterceptedException):
        driver.execute_script("arguments[0].click();", el)


def navigate_to_document_list(driver):
    """로그인 직후 화면에서 전자결재 > 문서함 > 기록물등록대장까지 자동으로 이동한다."""
    driver.switch_to.default_content()

    print("[안내] '전자결재' 메뉴로 이동합니다.")
    click_menu_by_text(driver, "전자결재")
    time.sleep(2)

    print("[안내] '문서함'을 엽니다.")
    try:
        click_menu_by_text(driver, "문서함")
        time.sleep(1)
    except TimeoutException:
        print("[안내] '문서함' 항목을 찾지 못했습니다. 이미 펼쳐진 상태일 수 있어 계속 진행합니다.")

    print("[안내] '기록물등록대장(모든부서)'를 클릭합니다.")
    try:
        click_menu_by_text(driver, "기록물등록대장(모든부서)")
    except TimeoutException:
        print(
            "[오류] '기록물등록대장(모든부서)' 항목을 찾지 못했습니다. "
            "위 [진단] 목록에서 실제 텍스트를 확인해 알려주세요."
        )
        raise
    time.sleep(2)


def check_stop_requested():
    """콘솔에서 Enter 키가 눌렸는지 논블로킹으로 확인한다 (Windows/Mac/Linux 공통)."""
    stopped = False
    while _key_pressed():
        ch = _read_key()
        if ch in (b"\r", b"\n"):
            stopped = True
    if stopped:
        print("\n[안내] 중단 요청을 받았습니다. 현재 처리 중인 문서까지 마치고 안전하게 종료합니다...")
    return stopped


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


def strip_existing_prefix(filename):
    """이미 붙어있는 '1. ' 같은 일련번호 접두어를 제거한다 (재실행 시 중복 방지)."""
    return re.sub(r"^\d+\.\s+", "", filename)


def rename_files_with_sequence(folder, log_csv_path):
    """완료목록.csv에 기록된 순서(=다운로드된 순서) 그대로, 폴더 내 PDF 파일명 앞에
    '1. ', '2. ' 형태의 일련번호를 붙인다. 가장 먼저 받은 파일이 1번이 된다.
    """
    if not log_csv_path or not os.path.exists(log_csv_path):
        return

    rows = []
    with open(log_csv_path, newline="", encoding="utf-8-sig") as f:
        for row in csv.reader(f):
            if row:
                rows.append(row)  # [doc_no, title, timestamp] - 다운로드된(추가된) 순서 그대로

    try:
        existing_files = os.listdir(folder)
    except FileNotFoundError:
        return

    used_files = set()
    seq = 1
    renamed_count = 0

    for doc_no, title, _ts in rows:
        candidate = None

        # 1순위: 파일명에 문서번호가 포함된 경우
        for f in existing_files:
            if f in used_files or not f.lower().endswith(".pdf"):
                continue
            base = strip_existing_prefix(f)
            if doc_no and doc_no in base:
                candidate = f
                break

        # 2순위: 제목과 파일명 유사도로 탐색
        if candidate is None:
            best, best_score = None, 0.0
            for f in existing_files:
                if f in used_files or not f.lower().endswith(".pdf"):
                    continue
                base = os.path.splitext(strip_existing_prefix(f))[0]
                score = difflib.SequenceMatcher(None, title, base).ratio()
                if score > best_score:
                    best, best_score = f, score
            if best and best_score >= 0.3:
                candidate = best

        if candidate is None:
            continue

        used_files.add(candidate)
        original_no_prefix = strip_existing_prefix(candidate)
        new_name = f"{seq}. {original_no_prefix}"
        if new_name != candidate:
            src = os.path.join(folder, candidate)
            dst = os.path.join(folder, new_name)
            try:
                if not os.path.exists(dst):
                    os.rename(src, dst)
                    renamed_count += 1
            except OSError as e:
                print(f"[경고] 파일명 변경 실패: {candidate} -> {e}")
        seq += 1

    print(f"[안내] {folder} 폴더 내 {renamed_count}개 파일에 일련번호를 부여했습니다.")


def run_download_loop(driver, done):
    """현재 검색 결과 화면을 기준으로, 마지막 페이지까지 순회하며 다운로드한다.
    중단(Enter) 요청 시 True를 반환한다.
    """
    page = 1
    stopped = False
    while True:
        if check_stop_requested():
            return True

        print(f"--- {page} 페이지 처리 중 ---")
        rows = get_rows(driver)
        print(f"[디버그] 찾은 행 개수: {len(rows)}")
        skipped_row_error = 0
        if len(rows) == 0:
            grids = driver.find_elements(By.CSS_SELECTOR, "div.grid-content")
            tables = driver.find_elements(By.TAG_NAME, "table")
            print(f"[디버그] div.grid-content 개수: {len(grids)}, table 태그 개수: {len(tables)}")
            for i, t in enumerate(tables):
                print(f"[디버그] table[{i}] class={t.get_attribute('class')!r} id={t.get_attribute('id')!r}")

        for row in rows:
            if check_stop_requested():
                return True

            try:
                info = row_info(row)
            except Exception as e:
                try:
                    raw_text = row.text.strip().replace("\n", " | ")
                except Exception:
                    raw_text = "(텍스트도 못 읽음)"
                print(f"[경고] 행을 읽는 중 오류로 건너뜀: {e} / 원본 텍스트: {raw_text!r}")
                skipped_row_error += 1
                continue

            if info["doc_no"] in done:
                print(f"이미 처리됨, 건너뜀: {info['doc_no']}")
                continue
            try:
                open_popup_and_save_pdf(driver, row, info)
                append_done(info["doc_no"], info["title"])
                done.add(info["doc_no"])
                print(f"완료: {info['doc_no']} - {info['title']}")
            except TimeoutException:
                print(f"실패(타임아웃): {info['doc_no']} - {info['title']}")
            except Exception as e:
                print(f"실패: {info['doc_no']} - {e}")
            time.sleep(1)

        print(f"[디버그] {page}페이지 요약: 행읽기오류 {skipped_row_error}건")

        if not go_next_page(driver):
            print("마지막 페이지입니다. 이번 기간 처리를 종료합니다.")
            return False

        page += 1


def main():
    global GROUPWARE_URL, DOWNLOAD_DIR, LOG_CSV

    default_url = GROUPWARE_URL
    entered_url = input(f"그룹웨어 주소를 입력하세요 (Enter 시 기본값: {default_url}): ").strip()
    GROUPWARE_URL = entered_url if entered_url else default_url

    user_id = input("그룹웨어 아이디를 입력하세요 (직접 로그인하려면 비워두고 Enter): ").strip()
    user_pw = ""
    if user_id:
        user_pw = getpass.getpass(
            "그룹웨어 비밀번호를 입력하세요 (입력한 문자는 화면에 표시되지 않습니다): "
        )

    driver = get_driver()
    driver.get(GROUPWARE_URL)

    if user_id and user_pw:
        auto_login(driver, user_id, user_pw)
    else:
        input("브라우저에서 직접 로그인해주세요. 완료 후 Enter를 누르세요: ")

    navigate_to_document_list(driver)
    switch_to_content_frame(driver)

    while True:
        print("\n등록일자 범위를 입력하세요. 형식: YYYY-MM-DD (조건 없이 진행하려면 비워두고 Enter)")
        date_start = parse_date_input("  시작일: ")
        date_end = parse_date_input("  종료일: ")

        period_label = compute_period_label(date_start, date_end)
        DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads", f"{period_label}_PDF")
        LOG_CSV = os.path.join(DOWNLOAD_DIR, "완료목록.csv")
        os.makedirs(DOWNLOAD_DIR, exist_ok=True)
        set_download_dir(driver, DOWNLOAD_DIR)
        done = load_done_list()

        switch_to_content_frame(driver)
        if date_start or date_end:
            print(f"[안내] 등록일자 적용 중... 시작일={date_start}, 종료일={date_end}")
            set_search_dates(driver, date_start, date_end)

        input(
            "등록일자를 채워 넣었습니다. 그룹웨어 화면에서 양식명 등 나머지 조건을 "
            "직접 설정하고 검색 버튼을 눌러주세요. 준비되면 여기로 돌아와 Enter를 누르세요: "
        )

        switch_to_content_frame(driver)
        print(f"[안내] 다운로드 폴더: {DOWNLOAD_DIR}")
        print("[안내] 지금부터 자동으로 문서를 처리합니다. 중단하려면 아무 때나 Enter를 누르세요.")

        stopped = run_download_loop(driver, done)

        rename_files_with_sequence(DOWNLOAD_DIR, LOG_CSV)

        if stopped:
            print("[안내] 사용자 요청으로 중단되었습니다.")
            break

        answer = input(
            "\n이번 기간 처리가 끝났습니다. 다른 기간을 이어서 처리할까요? (y: 계속 / n: 종료) [n]: "
        ).strip().lower()
        if answer != "y":
            break

    driver.quit()
    print(f"\n완료된 목록은 {LOG_CSV} 에서 확인할 수 있습니다.")


if __name__ == "__main__":
    main()

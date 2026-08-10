# -*- coding: utf-8 -*-
"""
전체목록(xlsx) vs 자동화 완료목록(csv) 대조 스크립트

전체목록 xlsx의 '신청내역'(문서제목) 컬럼과, 자동화가 만든 완료목록.csv의
제목(B열)을 대조해서, 아직 PDF를 확보하지 못한 건을 뽑아냅니다.

사용 전 설정:
- XLSX_PATH: 전체목록 엑셀 파일 경로
- XLSX_TITLE_COLUMN: 전체목록에서 문서제목이 있는 컬럼명 (기본 '신청내역')
- CSV_PATH: 자동화가 만든 완료목록.csv 경로
- OUTPUT_PATH: 결과를 저장할 엑셀 파일 경로

출력:
- '완전일치': 정확히 같은 제목을 완료목록에서 찾은 건 (확보됨)
- '유사매칭_확인필요': 완전히 같진 않지만 매우 비슷한 제목이 있는 건 (수동 확인 권장)
- '미확보': 완료목록에서 전혀 찾지 못한 건 (수동으로 찾아야 함)
세 개 시트로 나눠서 결과 xlsx를 만듭니다.
"""

import difflib
import glob
import os
import re

import pandas as pd

# 이 스크립트 파일이 있는 폴더를 기준으로 경로를 고정 (cmd 실행 위치와 무관하게 항상 동일)
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ------------------- 설정값 -------------------
# 파일명만 적으면 이 스크립트가 있는 폴더 기준으로 찾습니다.
# 다른 폴더의 파일을 쓰려면 전체 경로를 적어주세요. (예: r"C:\Users\svc\Documents\전체목록.xlsx")
XLSX_PATH = os.path.join(SCRIPT_DIR, "전체목록.xlsx")          # ★ 실제 파일명으로 교체
XLSX_TITLE_COLUMN = "신청내역"        # ★ 전체목록에서 문서제목 컬럼명

# 완료목록.csv를 2개 이상 대조하고 싶을 때: 리스트에 경로를 나열
# (1차 검색용 폴더의 완료목록.csv, 2차 검색용 폴더의 완료목록.csv 등)
CSV_PATHS = [
    os.path.join(SCRIPT_DIR, "완료목록.csv"),
    # os.path.join(SCRIPT_DIR, "..", "출장신청서_PDF_2차검색", "완료목록.csv"),
    # r"C:\Users\svc\Downloads\다른폴더\완료목록.csv",
]

# 매번 경로를 하나하나 적기 번거로우면, 아래처럼 특정 상위 폴더 안의
# 모든 "완료목록.csv"를 하위 폴더까지 자동으로 찾게 할 수도 있음.
# 사용하려면 CSV_PATHS = [] 로 비워두고 아래 두 줄의 주석을 해제하세요.
CSV_AUTO_SEARCH_ROOT = None  # 예: r"C:\Users\svc\Downloads"
# CSV_AUTO_SEARCH_ROOT = r"C:\Users\svc\Downloads"

OUTPUT_PATH = os.path.join(SCRIPT_DIR, "미확보목록.xlsx")       # 결과 저장 파일명

SIMILARITY_THRESHOLD = 0.85  # 이 이상 유사하면 '유사매칭'으로 분류


def normalize(text):
    """공백 차이 등으로 인한 오탐을 줄이기 위한 간단 정규화."""
    if pd.isna(text):
        return ""
    text = str(text)
    text = re.sub(r"\s+", " ", text)  # 연속 공백/개행을 하나로
    return text.strip()


def load_all_done_titles():
    """CSV_PATHS(또는 CSV_AUTO_SEARCH_ROOT)에 지정된 모든 완료목록.csv를 읽어 하나로 합친다.
    같은 등록번호가 여러 파일에 중복으로 있어도 한 번만 센다.
    """
    csv_paths = list(CSV_PATHS)

    if not csv_paths and CSV_AUTO_SEARCH_ROOT:
        csv_paths = glob.glob(
            os.path.join(CSV_AUTO_SEARCH_ROOT, "**", "완료목록.csv"), recursive=True
        )
        print(f"[안내] 자동 검색으로 찾은 완료목록.csv 개수: {len(csv_paths)}")
        for p in csv_paths:
            print(f"       - {p}")

    frames = []
    for path in csv_paths:
        if not os.path.exists(path):
            print(f"[경고] 파일을 찾을 수 없어 건너뜁니다: {path}")
            continue
        df = pd.read_csv(path, header=None, encoding="utf-8-sig",
                          names=["등록번호", "제목", "처리시각"])
        df["_출처파일"] = path
        frames.append(df)

    if not frames:
        raise ValueError("읽어들인 완료목록.csv가 하나도 없습니다. CSV_PATHS 설정을 확인해주세요.")

    combined = pd.concat(frames, ignore_index=True)
    before = len(combined)
    combined = combined.drop_duplicates(subset=["등록번호"], keep="first")
    after = len(combined)
    if before != after:
        print(f"[안내] 여러 목록에 중복된 등록번호 {before - after}건은 한 번만 반영했습니다.")

    print(f"[안내] 완료목록 합계: {after}건 (파일 {len(frames)}개 합산)")
    return combined


def main():
    # 1) 전체목록 읽기
    df_all = pd.read_excel(XLSX_PATH)
    if XLSX_TITLE_COLUMN not in df_all.columns:
        raise ValueError(
            f"'{XLSX_TITLE_COLUMN}' 컬럼을 찾을 수 없습니다. "
            f"실제 컬럼명: {list(df_all.columns)}"
        )
    df_all["_정규화제목"] = df_all[XLSX_TITLE_COLUMN].apply(normalize)

    # 2) 완료목록(csv, 1개 이상) 읽어서 합치기
    df_done = load_all_done_titles()
    done_titles = [normalize(t) for t in df_done["제목"]]
    done_titles_set = set(done_titles)

    exact_rows = []
    fuzzy_rows = []
    missing_rows = []

    for _, row in df_all.iterrows():
        title = row["_정규화제목"]
        if not title:
            missing_rows.append(row)
            continue

        if title in done_titles_set:
            exact_rows.append(row)
            continue

        # 완전일치가 아니면 가장 비슷한 제목과 유사도 계산
        best_match = difflib.get_close_matches(title, done_titles, n=1, cutoff=SIMILARITY_THRESHOLD)
        if best_match:
            row = row.copy()
            row["_가장비슷한완료제목"] = best_match[0]
            fuzzy_rows.append(row)
        else:
            missing_rows.append(row)

    df_exact = pd.DataFrame(exact_rows).drop(columns=["_정규화제목"], errors="ignore")
    df_fuzzy = pd.DataFrame(fuzzy_rows).drop(columns=["_정규화제목"], errors="ignore")
    df_missing = pd.DataFrame(missing_rows).drop(columns=["_정규화제목"], errors="ignore")

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        df_missing.to_excel(writer, sheet_name="미확보", index=False)
        df_fuzzy.to_excel(writer, sheet_name="유사매칭_확인필요", index=False)
        df_exact.to_excel(writer, sheet_name="완전일치", index=False)

    total = len(df_all)
    print(f"전체 {total}건 중")
    print(f"  - 완전일치(확보됨): {len(df_exact)}건")
    print(f"  - 유사매칭(확인 필요): {len(df_fuzzy)}건")
    print(f"  - 미확보(수동으로 찾아야 함): {len(df_missing)}건")
    print(f"결과 파일: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()

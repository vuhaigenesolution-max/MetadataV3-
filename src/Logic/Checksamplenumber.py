"""
Checksamplenumber.py
--------------------
Đọc file SUM (Pool Allocation, multi-page) → xuất Excel với 3 cột: RunName, Col 3, Col 5.

Đầu vào:
    File SUM Excel với sheet chứa nhiều "page" lặp ngang.
    Mỗi page = 20 cột dữ liệu + 2 cột trống (STEP = 22).
    Trong 1 page:
        - Dòng 1 (meta):  A1=RunContent, E1=RunNameNotes, Q1=MachineNotes
        - Dòng 4 : header bảng (sẽ thay bằng "Col 1".."Col 20")
        - Dòng 5+: data, key trim theo Col 3 (cột C trong page)

Luồng:
    1. detect_pages()                  — tự dò số page có data
    2. read_one_page_flatten_merge()   — đọc 1 page, expand merged cells, trim theo Col 3
    3. parse_file_by_pages()           — duyệt mọi page → 1 DataFrame phẳng + 3 cột meta
    4. source_to_df()                  — lọc T7/MGI, trích RunName
    5. (__main__)                      — pick folder → xuất Excel chỉ 3 cột
"""

import os
import re
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string


# ── Cấu hình page layout ──────────────────────────────────
PAGE_COLS = 20            # A..T
STEP = 22                 # 20 cột + 2 cột trống
START_EXCEL_ROW = 5       # data từ dòng 5 (Excel, 1-based)


def _is_blank(x):
    if x is None:
        return True
    if isinstance(x, str) and x.strip() == "":
        return True
    return False


def detect_pages(ws, page_cols=PAGE_COLS, step=STEP, probe_rows=4):
    """Tự dò số page: quét block STEP cột, page hợp lệ nếu probe_rows đầu có data."""
    pages = []
    max_col = ws.max_column

    for col_start in range(1, max_col + 1, step):
        col_end = min(col_start + page_cols - 1, max_col)

        found = False
        for r in range(1, min(ws.max_row, probe_rows) + 1):
            for c in range(col_start, col_end + 1):
                if not _is_blank(ws.cell(r, c).value):
                    found = True
                    break
            if found:
                break

        if found:
            pages.append(col_start)
        else:
            break

    return pages


def read_one_page_flatten_merge(
    ws,
    col_start: int,
    page_cols: int = 20,
    trim_by_key: bool = True,
    key_col_letter_in_page: str = "C",
    start_excel_row: int = 5,
):
    """Đọc 1 page, expand merged cells, trim theo key column. Trả về df_page (raw, chưa cắt header)."""
    col_end = col_start + page_cols - 1
    max_row = ws.max_row
    key_offset = column_index_from_string(key_col_letter_in_page) - 1  # C → 2

    # 1) đọc raw 20 cột của page
    data = [
        [ws.cell(r, c).value for c in range(col_start, col_end + 1)]
        for r in range(1, max_row + 1)
    ]

    # 2) trim theo key column TRƯỚC khi expand merge
    if trim_by_key:
        df_raw = pd.DataFrame(data).replace(r'^\s*$', np.nan, regex=True)
        start_df_row = start_excel_row - 1
        if df_raw.shape[0] > start_df_row:
            s = df_raw.iloc[start_df_row:, key_offset]
            m = s.notna() & s.astype(str).str.strip().ne("")
            if m.any():
                last_idx = start_df_row + np.where(m.to_numpy())[0].max()
                data = data[: last_idx + 1]
            else:
                data = data[: max(4, start_df_row)]

    # 3) expand merged cells trong phạm vi page (KHÔNG ffill)
    max_row_cut = len(data)
    for mr in ws.merged_cells.ranges:
        overlap_c1 = max(mr.min_col, col_start)
        overlap_c2 = min(mr.max_col, col_end)
        if overlap_c1 > overlap_c2:
            continue
        r1 = max(mr.min_row, 1)
        r2 = min(mr.max_row, max_row_cut)
        if r1 > r2:
            continue
        top_left_val = ws.cell(mr.min_row, mr.min_col).value
        if top_left_val is None:
            continue
        for r in range(r1, r2 + 1):
            row_idx = r - 1
            for c in range(overlap_c1, overlap_c2 + 1):
                col_idx = c - col_start
                data[row_idx][col_idx] = top_left_val

    return pd.DataFrame(data).replace(r'^\s*$', np.nan, regex=True)


def parse_file_by_pages(source_file_path, sheet_name):
    """Duyệt mọi page → concat thành 1 DataFrame phẳng + 3 cột meta."""
    wb = load_workbook(source_file_path, data_only=True)
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active

    page_starts = detect_pages(ws, page_cols=PAGE_COLS, step=STEP, probe_rows=4)
    page_num = len(page_starts)

    df_lst = []
    for col_start in page_starts:
        df_page = read_one_page_flatten_merge(ws, col_start, PAGE_COLS, START_EXCEL_ROW)

        # 3 thông tin meta ở dòng 1 (index 0) của page
        run_content   = df_page.iat[0, 0]    # A1
        run_name_note = df_page.iat[0, 4]    # E1
        machine_note  = df_page.iat[0, 16]   # Q1

        # bỏ 3 dòng đầu, dòng 4 (index 3) làm header rồi bỏ ra
        df_1 = df_page.iloc[3:].copy()
        df_1.columns = df_1.iloc[0].astype(str)
        df_1 = df_1.iloc[1:].reset_index(drop=True)

        # ép về 20 cột với tên Col 1..Col 20
        df_1 = df_1.iloc[:, :PAGE_COLS]
        df_1.columns = [f"Col {j}" for j in range(1, PAGE_COLS + 1)]

        # trim đến dòng đầu Col 3 blank
        c3 = df_1["Col 3"].astype(str).str.strip()
        if c3.eq("").any():
            df_1 = df_1.iloc[:c3.eq("").idxmax()].copy()

        # giữ dòng có Col 3 và Col 1 đều khác trống
        mask = (
            df_1["Col 3"].notna() & df_1["Col 3"].astype(str).str.strip().ne("") &
            df_1["Col 1"].notna() & df_1["Col 1"].astype(str).str.strip().ne("")
        )
        df_1 = df_1.loc[mask].reset_index(drop=True)

        df_1["RunContent"]   = run_content
        df_1["RunNameNotes"] = run_name_note
        df_1["MachineNotes"] = machine_note
        df_lst.append(df_1)

    return pd.concat(df_lst, ignore_index=True) if df_lst else pd.DataFrame(), page_num


def source_to_df(source_path, sheet):
    """Đọc file SUM, trích RunName. Trả về DataFrame có Col 1..Col 20 + RunName (mọi máy)."""
    result, page_num = parse_file_by_pages(source_path, sheet)
    print("page_num =", page_num)

    df = result.copy()
    df["RunName"] = (
        df["RunContent"].astype(str)
        .str.extract(r":\s*([^,]+)", expand=False)
        .str.strip()
    )

    return df


# ══════════════════════════════════════════════════════════
# Đọc folder metadata → gộp thành 1 DataFrame
# ══════════════════════════════════════════════════════════
def read_meta_folder(folder: str, sheet_name: str = "Sample") -> pd.DataFrame:
    """Đọc tất cả file metadata trong folder → 1 DataFrame gộp.

    Quy tắc cho mỗi file:
        - RunName : regex 'metadata_(.+?)_\\d+' từ tên file
                     (fallback: tên file không đuôi nếu pattern không match)
        - Header  : Excel row 21, lấy 4 cột Excel A, B, U, V
        - Data    : Excel row 22 trở đi, cùng 4 cột A, B, U, V

    Lưu ý: A/B/U/V là CỘT EXCEL (theo vị trí), KHÔNG phải tên cột.

    Return: DataFrame gộp với cột đầu = RunName, 4 cột sau = header từ row 21.
            Các dòng mà cả 4 ô đều trống sẽ bị bỏ.
    """
    import glob as _glob

    files = sorted(
        _glob.glob(os.path.join(folder, "*.xlsx"))
        + _glob.glob(os.path.join(folder, "*.xls"))
    )
    file_pat = re.compile(r"metadata_([RP]\d{4,5})_(\d{8})", re.IGNORECASE)
    target_cols = [0, 1, 20, 21]   # 0-indexed: A=0, B=1, U=20, V=21

    df_lst = []
    for fp in files:
        fname = os.path.basename(fp)
        m = file_pat.search(fname)
        runname = m.group(1) if m else os.path.splitext(fname)[0]

        wb = load_workbook(fp, data_only=True, read_only=True)
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active

        # Header ở row 21 — lấy 4 ô tương ứng A, B, U, V
        header_row = next(ws.iter_rows(min_row=21, max_row=21, values_only=True), ())
        headers = [
            str(header_row[i]).strip()
            if i < len(header_row) and header_row[i] not in (None, "")
            else f"Col_{i + 1}"
            for i in target_cols
        ]

        # Data từ row 22 trở đi
        rows = []
        for row in ws.iter_rows(min_row=22, values_only=True):
            picked = [row[i] if i < len(row) else None for i in target_cols]
            if all(v is None or (isinstance(v, str) and v.strip() == "") for v in picked):
                continue
            rows.append(picked)

        wb.close()

        df_f = pd.DataFrame(rows, columns=headers)
        df_f.insert(0, "RunName", runname)
        df_lst.append(df_f)

    return pd.concat(df_lst, ignore_index=True) if df_lst else pd.DataFrame()


# ══════════════════════════════════════════════════════════
# Pipeline đầy đủ cho file SUM
# ══════════════════════════════════════════════════════════
_RANGE_PAT = re.compile(r"^\s*(\d+)\s*-\s*(\d+)\s*$")
_NIPT_PAT  = re.compile(r"^(?:CRH|CR|EK|TK|GS|[ETBHLVP])\d+", flags=re.IGNORECASE)


def _parse_col5(val):
    """Col 5 → (sample_orders[], loai_flags[]).
       'X-Y'        → range X..Y, không loại
       'X-Y/<excl>' → range X..Y, đánh 'Loại' cho số trong <excl>
       khác         → 1 dòng giữ nguyên giá trị
    """
    s = str(val).strip()
    if "/" in s:
        main_part, exc_part = s.split("/", 1)
    else:
        main_part, exc_part = s, ""

    m = _RANGE_PAT.match(main_part.strip())
    if not m:
        return [s], [""]
    start, end = int(m.group(1)), int(m.group(2))
    if end < start:
        return [s], [""]

    orders = list(range(start, end + 1))
    excluded = set()
    for token in exc_part.split(","):
        token = token.strip()
        if not token:
            continue
        m2 = _RANGE_PAT.match(token)
        if m2:
            a, b = int(m2.group(1)), int(m2.group(2))
            excluded.update(range(a, b + 1))
        else:
            try:
                excluded.add(int(token))
            except ValueError:
                pass

    loai = ["Loại" if o in excluded else "" for o in orders]
    return orders, loai


def _classify_col3(val):
    """Col 3 → 'NIPT' nếu match pattern, ngược lại 'other'."""
    s = str(val).strip().upper()
    if s in ("PGS", "PGTM"):
        return "NIPT"
    return "NIPT" if _NIPT_PAT.match(s) else "other"


def process_sum_file(sum_path: str, sheet: str = "Sheet1") -> pd.DataFrame:
    """Đọc + xử lý file SUM → DataFrame 6 cột:
       RunName, Col 3, Col 5, sample_order, Loại, Sample Type

    Lưu ý: Col 3 nếu chứa "(...)" sẽ bị loại bỏ phần đó (gồm cả ngoặc và nội dung bên trong),
    ví dụ: 'ABC(XYZ)' → 'ABC'.
    """
    df = source_to_df(sum_path, sheet)

    df_out = (
        df[["RunName", "Col 3", "Col 5"]]
        .replace(r"^\s*$", np.nan, regex=True)
        .dropna()
        .reset_index(drop=True)
    )

    # Clean Col 3: bỏ "(...)" và content bên trong
    df_out["Col 3"] = (
        df_out["Col 3"].astype(str)
        .str.replace(r"\s*\([^)]*\)", "", regex=True)
        .str.strip()
    )

    parsed = df_out["Col 5"].apply(_parse_col5)
    df_out["sample_order"] = parsed.apply(lambda x: x[0])
    df_out["Loại"]         = parsed.apply(lambda x: x[1])
    df_out = df_out.explode(["sample_order", "Loại"]).reset_index(drop=True)

    df_out["Sample Type"] = df_out["Col 3"].apply(_classify_col3)
    return df_out


def _norm(s):
    """Chuẩn hoá series về string đã strip + bỏ '.0' đuôi (tránh mismatch 1 vs 1.0)."""
    return s.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)


def find_meta_only(df_meta: pd.DataFrame, df_sum: pd.DataFrame) -> pd.DataFrame:
    """Trả về các dòng có trong df_meta nhưng KHÔNG match trong df_sum.

    df_meta cột (theo thứ tự):
        RunName, expNum (col A), sampleOrder (col B), Col_21 (col U), Col_22 (col V)
    df_sum cột:
        RunName, Col 3, Col 5, sample_order, Loại, Sample Type

    Rule match (mỗi df_meta row có "match" nếu thoả 1 trong 2):
        - NIPT  : tồn tại df_sum row có Sample Type='NIPT', Loại rỗng,
                  RunName khớp, Col 3 = expNum, sample_order = sampleOrder
        - other : tồn tại df_sum row có Sample Type='other',
                  RunName khớp, sample_order = sampleOrder,
                  Col 3 = Col_22 (nếu Col_22 không trống) HOẶC = Col_21 (nếu Col_22 trống)

    Output: 3 cột RunName, expNum, sampleOrder (chỉ những row chưa match).
    """
    cols = list(df_meta.columns)
    if len(cols) < 5:
        return df_meta.copy()

    rn_col, exp_col, ord_col, u_col, v_col = cols[:5]

    if df_meta.empty:
        return df_meta[[rn_col, exp_col, ord_col]].copy()

    # ── Build match keys cho df_meta
    rn_m  = _norm(df_meta[rn_col])
    exp_m = _norm(df_meta[exp_col])
    ord_m = _norm(df_meta[ord_col])
    u_m   = _norm(df_meta[u_col])
    v_m   = _norm(df_meta[v_col])

    nipt_key  = rn_m + "|" + exp_m + "|" + ord_m
    other_id  = v_m.where(v_m.ne(""), u_m)        # ưu tiên V, fallback U
    other_key = rn_m + "|" + other_id + "|" + ord_m

    # ── Build set keys từ df_sum
    rn_s  = _norm(df_sum["RunName"])
    c3_s  = _norm(df_sum["Col 3"])
    ord_s = _norm(df_sum["sample_order"])
    sum_key = rn_s + "|" + c3_s + "|" + ord_s

    nipt_mask  = (df_sum["Sample Type"] == "NIPT") & (df_sum["Loại"].astype(str).str.strip() == "")
    other_mask = df_sum["Sample Type"] == "other"

    nipt_keys  = set(sum_key[nipt_mask])
    other_keys = set(sum_key[other_mask])

    matched = nipt_key.isin(nipt_keys) | other_key.isin(other_keys)
    return df_meta.loc[~matched, [rn_col, exp_col, ord_col]].reset_index(drop=True)


# ══════════════════════════════════════════════════════════
# RUN STANDALONE — chọn folder input, xuất Excel
# ══════════════════════════════════════════════════════════
if __name__ == "__main__":
    import glob
    import sys
    import tkinter as tk
    from tkinter import filedialog

    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 200)
    pd.set_option("display.max_rows", 200)

    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title="Chọn folder chứa file SUM Excel")
    if not folder:
        print("Không chọn folder — thoát")
        sys.exit(0)

    candidates = sorted(glob.glob(os.path.join(folder, "SUM*.xlsx")))
    if not candidates:
        candidates = sorted(
            glob.glob(os.path.join(folder, "*.xlsx"))
            + glob.glob(os.path.join(folder, "*.xls"))
        )
    if not candidates:
        print(f"Không tìm thấy file Excel nào trong: {folder}")
        sys.exit(1)

    input_path = candidates[0]
    print(f"Đang xử lý: {input_path}")
    df_out = process_sum_file(input_path, sheet="Sheet1")
    print(df_out)
    print(f"\nShape: {df_out.shape}")

    in_stem  = os.path.splitext(os.path.basename(input_path))[0]
    out_path = os.path.join(folder, f"{in_stem}_solution.xlsx")
    df_out.to_excel(out_path, index=False)
    print(f"\n✓ Đã xuất: {out_path}")


# -----------------------------




"""
SampleImport.py
---------------
Xử lý file metadata Excel đã được combine, tạo 2 file CSV output:
  Output 1 : SampleImport_<tên file>.csv   — từ sheet SampleImport
  Output 2 : <tên file>_<date>.csv         — từ sheet Aviti Manifest

Luồng chính:
  1. Đọc file nguồn (phải có sheet "Sample")
  2. Xác định dòng cuối dữ liệu theo cột A của sheet Sample
  3. Đọc values-only (không formula) từ SampleImport và Aviti Manifest
  4. Làm sạch: replace 0 → ""
  5. Map test-type (cột K của SampleImport) theo file template
  6. Xuất 2 file CSV
"""

import os
import re
import logging
import pandas as pd
from openpyxl import load_workbook

log = logging.getLogger(__name__)

# ── Tên sheet ─────────────────────────────────────────────
SHEET_SAMPLE  = "Sample"
SHEET_IMPORT  = "SampleImport"
SHEET_AVITI   = "Aviti Manifest"

# ── Dòng bắt đầu dữ liệu (1-indexed, Excel) ──────────────
SAMPLE_START_ROW = 22
IMPORT_START_ROW = 24
AVITI_START_ROW  = 16

# ── Cột A (0-indexed) dùng xác định dòng cuối ─────────────
COL_A = 0

# ── Mapping test-type (lowercase key) ─────────────────────
TEST_TYPE_MAP = {
    "tsprocare": "TSPRO",
    "tsthalass": "TS3",
    "ts9.5":     "TS95",
}

# ── Regex lấy runname + date từ tên file ──────────────────
FILE_PATTERN = re.compile(r"metadata_(.+?)_(\d{8})", re.IGNORECASE)


# ══════════════════════════════════════════════════════════
# HÀM HỖ TRỢ
# ══════════════════════════════════════════════════════════

def _get_last_data_row(ws, start_row: int) -> int:
    """
    Trả về index dòng cuối (1-indexed) có giá trị khác None / "" ở cột A.
    Duyệt từ start_row xuống cuối sheet.
    """
    last = start_row - 1
    for row in ws.iter_rows(min_row=start_row, min_col=1, max_col=1):
        cell_val = row[0].value
        if cell_val is not None and str(cell_val).strip() != "":
            last = row[0].row
    return last


def _sheet_to_df(wb, sheet_name: str, start_row: int, last_row: int) -> pd.DataFrame:
    """
    Đọc dữ liệu values-only (không formula) từ một sheet của workbook.
    Trả về DataFrame, cột đặt tên theo chữ cái Excel (A, B, C, ...).
    """
    if sheet_name not in wb.sheetnames:
        log.warning(f"Sheet '{sheet_name}' không tồn tại trong file — trả về DataFrame rỗng")
        return pd.DataFrame()

    ws = wb[sheet_name]
    max_col = ws.max_column

    rows = []
    for row in ws.iter_rows(min_row=start_row, max_row=last_row,
                             min_col=1, max_col=max_col):
        rows.append([cell.value for cell in row])

    # Đặt tên cột theo chữ cái Excel: A, B, C, ... Z, AA, AB, ...
    col_names = _excel_col_names(max_col)
    df = pd.DataFrame(rows, columns=col_names)
    return df


def _excel_col_names(n: int) -> list[str]:
    """Sinh tên cột Excel: A, B, ..., Z, AA, AB, ... cho n cột."""
    names = []
    for i in range(n):
        name, col = "", i + 1
        while col:
            col, rem = divmod(col - 1, 26)
            name = chr(65 + rem) + name
        names.append(name)
    return names


def _clean_zeros(df: pd.DataFrame) -> pd.DataFrame:
    """Thay thế các ô có giá trị 0 (hoặc "0") thành chuỗi rỗng ""."""
    return df.replace({0: "", "0": ""})


def _map_test_type(val) -> str:
    """
    Chuẩn hoá giá trị test-type:
      tsprocare → TSPRO
      tsthalass → TS3
      ts9.5     → TS95
      còn lại   → UPPERCASE
    """
    if pd.isna(val) or str(val).strip() == "":
        return val
    key = str(val).strip().lower()
    return TEST_TYPE_MAP.get(key, str(val).strip().upper())


def _apply_template_mapping(df_import: pd.DataFrame,
                             template_path: str) -> pd.DataFrame:
    """
    Map cột K (test-type) của SampleImport theo file template.

    Điều kiện ghép:
      - Cột A của df_import == (col C của template) + "-" + (col D của template)
      - Cột C của df_import khớp với cột C của template

    Khi điều kiện thoả → chuẩn hoá giá trị cột K theo TEST_TYPE_MAP.
    Nếu không có template hoặc không match → vẫn uppercase cột K.
    """
    if not os.path.isfile(template_path):
        log.warning(f"Template mapping không tìm thấy: {template_path} — bỏ qua mapping")
        # Vẫn chuẩn hoá uppercase
        if "K" in df_import.columns:
            df_import = df_import.copy()
            df_import["K"] = df_import["K"].apply(_map_test_type)
        return df_import

    # Đọc template (lấy sheet đầu tiên, header ở dòng 1)
    df_tmpl = pd.read_excel(template_path, header=0, dtype=str)
    df_tmpl = df_tmpl.fillna("")

    # Cột template dùng để tạo key ghép: col C + "-" + col D (index 2, 3)
    tmpl_cols = df_tmpl.columns.tolist()
    if len(tmpl_cols) < 4:
        log.warning("Template mapping không đủ 4 cột — bỏ qua mapping")
        if "K" in df_import.columns:
            df_import = df_import.copy()
            df_import["K"] = df_import["K"].apply(_map_test_type)
        return df_import

    col_c_tmpl = tmpl_cols[2]   # cột thứ 3 (index 2)
    col_d_tmpl = tmpl_cols[3]   # cột thứ 4 (index 3)

    # Tạo key ghép trong template: "C-D"
    df_tmpl["_join_key"] = (
        df_tmpl[col_c_tmpl].str.strip() + "-" + df_tmpl[col_d_tmpl].str.strip()
    )

    # Tạo set key hợp lệ để tra nhanh
    valid_keys = set(df_tmpl["_join_key"].str.lower())

    df_import = df_import.copy()

    if "K" not in df_import.columns:
        log.warning("Không tìm thấy cột K trong SampleImport — bỏ qua mapping")
        return df_import

    def _map_row(row):
        col_a = str(row.get("A", "")).strip().lower()
        # TODO: bổ sung điều kiện "cột C match" khi có thêm thông tin cột C cần khớp
        if col_a in valid_keys:
            return _map_test_type(row["K"])
        return _map_test_type(row["K"])   # fallback: vẫn chuẩn hoá

    df_import["K"] = df_import.apply(_map_row, axis=1)
    return df_import



# ══════════════════════════════════════════════════════════
# HÀM CHÍNH
# ══════════════════════════════════════════════════════════

def process_sample_import(source_path: str,
                           output_folder: str,
                           template_path: str = "") -> dict:
    """
    Xử lý một file metadata Excel đã combine.

    Args:
        source_path   : đường dẫn file Excel đầu vào
        output_folder : thư mục lưu 2 file CSV output
        template_path : đường dẫn file template mapping (tuỳ chọn)

    Returns:
        dict với keys "csv_import" và "csv_aviti" là đường dẫn file đã xuất.
    """
    os.makedirs(output_folder, exist_ok=True)

    fname = os.path.basename(source_path)
    stem  = os.path.splitext(fname)[0]   # tên file không đuôi

    # Lấy date từ tên file (dùng cho tên output 2)
    m = FILE_PATTERN.search(fname)
    date_str = m.group(2) if m else "00000000"

    log.info(f"Xử lý file: {fname}")

    # ── 1. Mở workbook (data_only=True để lấy values thay vì formula) ──
    try:
        wb = load_workbook(source_path, data_only=True, read_only=True)
    except Exception as e:
        raise ValueError(f"Không mở được file Excel: {e}")

    # ── 2. Kiểm tra sheet Sample ──────────────────────────
    if SHEET_SAMPLE not in wb.sheetnames:
        wb.close()
        raise ValueError(f"File '{fname}' không có sheet '{SHEET_SAMPLE}' — bỏ qua")

    ws_sample = wb[SHEET_SAMPLE]

    # ── 3. Xác định dòng cuối theo cột A của sheet Sample ─
    last_row_sample = _get_last_data_row(ws_sample, SAMPLE_START_ROW)
    num_rows        = last_row_sample - SAMPLE_START_ROW + 1

    if num_rows <= 0:
        wb.close()
        log.warning(f"Sheet Sample không có dữ liệu (file: {fname})")
        return {}

    log.info(f"  Dòng dữ liệu Sample: {SAMPLE_START_ROW} → {last_row_sample} ({num_rows} dòng)")

    # Tính dòng cuối tương ứng cho các sheet khác (cùng số dòng dữ liệu)
    last_row_import = IMPORT_START_ROW + num_rows - 1
    last_row_aviti  = AVITI_START_ROW  + num_rows - 1

    # ── 4. Đọc SampleImport ───────────────────────────────
    df_import = _sheet_to_df(wb, SHEET_IMPORT, IMPORT_START_ROW, last_row_import)
    log.info(f"  SampleImport: {len(df_import)} dòng × {len(df_import.columns)} cột")

    # ── 5. Đọc Aviti Manifest ─────────────────────────────
    df_aviti = _sheet_to_df(wb, SHEET_AVITI, AVITI_START_ROW, last_row_aviti)
    log.info(f"  Aviti Manifest: {len(df_aviti)} dòng × {len(df_aviti.columns)} cột")

    wb.close()

    # ── 6. Làm sạch: 0 → "" ──────────────────────────────
    df_import = _clean_zeros(df_import)
    df_aviti  = _clean_zeros(df_aviti)

    # ── 7. Map test-type (cột K) theo template ─────────────
    if df_import is not None and not df_import.empty:
        df_import = _apply_template_mapping(df_import, template_path)

    # ── 10. Xuất Output 1: SampleImport CSV ───────────────
    csv1_name = f"SampleImport_{stem}.csv"
    csv1_path = os.path.join(output_folder, csv1_name)
    df_import.to_csv(csv1_path, index=False, encoding="utf-8-sig")
    log.info(f"  → Output 1: {csv1_name}")

    # ── 11. Xuất Output 2: Aviti Manifest CSV ─────────────
    csv2_name = f"{stem}_{date_str}.csv"
    csv2_path = os.path.join(output_folder, csv2_name)
    df_aviti.to_csv(csv2_path, index=False, encoding="utf-8-sig")
    log.info(f"  → Output 2: {csv2_name}")

    return {"csv_import": csv1_path, "csv_aviti": csv2_path}

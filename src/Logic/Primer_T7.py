"""
Primer_T7.py
------------
Kiểm tra và loại trừ primer trùng nhau giữa cột T (primer cần check)
và cột A (primer gốc) trên nhiều file Excel đầu vào.

Luồng:
  1. Đọc tất cả file được chọn (sheet "Sample", bắt đầu dòng 22)
  2. Gom toàn bộ:
       - Cột T  → primer cần kiểm tra (kèm metadata B, C, D, RunName)
       - Cột A  → tập primer gốc dùng để loại trừ
  3. Lọc: chỉ giữ T-entries KHÔNG xuất hiện trong tập A
  4. Trả về kết quả:
       - Nếu không có primer trùng  → message "no_duplicate"  (không tạo file)
       - Nếu tất cả T đều trùng     → message "all_duplicate" (không tạo file)
       - Nếu có primer hợp lệ còn lại → xuất Primer_Check_Result.xlsx

Output columns: Primer, RunName, ExpNum, SampleOrder, LabCode
"""

import os
import logging
import pandas as pd

log = logging.getLogger(__name__)

# ── Cấu hình ─────────────────────────────────────────────
SHEET_NAME  = "Sample"
START_ROW   = 22        # dòng bắt đầu dữ liệu (1-indexed Excel)
OUTPUT_NAME = "Primer_Check_Result.xlsx"

# Vị trí cột trong sheet nguồn (0-indexed)
COL_A = 0   # Primer gốc  (exclusion list)
COL_B = 1   # ExpNum
COL_C = 2   # SampleOrder
COL_D = 3   # LabCode
COL_T = 19  # Primer cần check


# ══════════════════════════════════════════════════════════
# HÀM HỖ TRỢ
# ══════════════════════════════════════════════════════════

def _read_file(fp: str) -> pd.DataFrame:
    """
    Đọc sheet 'Sample' của 1 file Excel, trả về DataFrame với 5 cột:
    A (primer gốc), B (ExpNum), C (SampleOrder), D (LabCode), T (primer check).
    Loại bỏ các dòng mà cột A trống.
    """
    df = pd.read_excel(
        fp,
        sheet_name=SHEET_NAME,
        header=None,
        skiprows=START_ROW - 1,
        usecols=[COL_A, COL_B, COL_C, COL_D, COL_T],
        dtype=str,
    )
    df.columns = ["A", "B", "C", "D", "T"]
    df = df[df["A"].notna() & (df["A"].str.strip() != "")]
    return df.reset_index(drop=True)


# ══════════════════════════════════════════════════════════
# HÀM CHÍNH
# ══════════════════════════════════════════════════════════

def process_primer_t7(
    file_paths: list,
    output_folder: str,
    progress_callback=None,
) -> dict:
    """
    Kiểm tra primer trùng lặp giữa cột T và cột A trên nhiều file Excel.

    Args:
        file_paths       : danh sách đường dẫn file Excel đầu vào
        output_folder    : folder lưu file kết quả
        progress_callback: hàm (current, total) để cập nhật progress bar UI

    Returns:
        dict với các keys:
          total       – tổng số T-primer đã kiểm tra
          duplicates  – số T-primer bị loại (trùng với A)
          valid       – số T-primer hợp lệ (giữ lại)
          output_path – đường dẫn file xuất (None nếu không tạo file)
          message     – "ok" | "no_duplicate" | "all_duplicate" | "no_data"
    """
    os.makedirs(output_folder, exist_ok=True)
    total_files = len(file_paths)

    t_rows: list[dict] = []   # T-primers kèm metadata
    a_set:  set[str]   = set()  # tập primer gốc (dùng để loại trừ)

    # ── 1. Đọc tất cả file ──────────────────────────────
    for idx, fp in enumerate(file_paths, start=1):
        fname   = os.path.basename(fp)
        runname = os.path.splitext(fname)[0]
        log.info(f"[{idx}/{total_files}] Đọc: {fname}")

        try:
            df = _read_file(fp)
        except Exception as e:
            log.error(f"  Lỗi khi đọc {fname}: {e}")
            if progress_callback:
                progress_callback(idx, total_files)
            continue

        # Gom A primers vào exclusion set (case-insensitive)
        for val in df["A"]:
            s = str(val).strip()
            if s and s.upper() != "NAN":
                a_set.add(s.upper())

        # Gom T primers kèm metadata
        for _, row in df.iterrows():
            t_val = str(row.get("T", "")).strip()
            if not t_val or t_val.upper() == "NAN":
                continue
            t_rows.append({
                "Primer":      t_val,
                "_key":        t_val.upper(),
                "RunName":     runname,
                "ExpNum":      str(row.get("B", "")).strip(),
                "SampleOrder": str(row.get("C", "")).strip(),
                "LabCode":     str(row.get("D", "")).strip(),
            })

        log.info(f"  → {len(df)} dòng hợp lệ | "
                 f"{sum(1 for r in df['T'] if str(r).strip())} T-primers")

        if progress_callback:
            progress_callback(idx, total_files)

    # ── 2. Phân loại ────────────────────────────────────
    total_checked  = len(t_rows)
    duplicate_rows = [r for r in t_rows if r["_key"] in a_set]
    valid_rows     = [r for r in t_rows if r["_key"] not in a_set]

    log.info("─" * 50)
    log.info(f"Tổng primer đã check  : {total_checked}")
    log.info(f"Primer trùng (bị loại): {len(duplicate_rows)}")
    log.info(f"Primer hợp lệ (giữ)   : {len(valid_rows)}")

    # ── 3. Xử lý các trường hợp đặc biệt ───────────────
    if total_checked == 0:
        log.warning("Không đọc được T-primer nào từ các file đầu vào.")
        return {"total": 0, "duplicates": 0, "valid": 0,
                "output_path": None, "message": "no_data"}

    if len(duplicate_rows) == 0:
        # Không có primer nào trùng → thông báo, không tạo file
        log.info("Không có primer nào trùng.")
        return {"total": total_checked, "duplicates": 0, "valid": total_checked,
                "output_path": None, "message": "no_duplicate"}

    # ── 4. Xuất file kết quả — chứa các primer BỊ TRÙNG ────
    out_cols = ["Primer", "RunName", "ExpNum", "SampleOrder", "LabCode"]
    df_out   = pd.DataFrame(duplicate_rows)[out_cols]

    out_path = os.path.join(output_folder, OUTPUT_NAME)
    df_out.to_excel(out_path, index=False, engine="openpyxl")

    log.info(f"Xuất file: {OUTPUT_NAME} ({len(df_out)} dòng trùng) → {output_folder}")

    return {
        "total":       total_checked,
        "duplicates":  len(duplicate_rows),
        "valid":       len(valid_rows),
        "output_path": out_path,
        "message":     "ok",
    }

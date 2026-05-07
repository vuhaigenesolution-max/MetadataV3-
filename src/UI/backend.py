import os
import sys
import pandas as pd

_SRC_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if _SRC_DIR not in sys.path:
    sys.path.insert(0, _SRC_DIR)

from Logic.CombieFile import process_combine_files
from Logic.Primer_T7 import process_primer_t7
from Logic.SampleImport import process_sample_import
from Logic.CheckDescription import run_check_description
from Logic.Checksamplenumber import (
    read_meta_folder, process_sum_file, find_meta_only,
)


def run_check_sample_number(meta_folder, sum_file, output_folder, progress_callback=None):
    """Chạy pipeline check sample number:
       - meta_folder  : folder chứa nhiều file metadata (Source Metadata)
       - sum_file     : file SUM Excel (File Sum)
       - output_folder: nơi xuất file kết quả

       Output:
         - meta_only.xlsx : các dòng meta KHÔNG match trong SUM (chỉ xuất khi có)
         Nếu meta khớp đầy đủ với SUM → KHÔNG xuất file, trả về meta_only_path=None.

       Return dict với đường dẫn file (nếu có) + số dòng từng bảng.
    """
    os.makedirs(output_folder, exist_ok=True)

    if progress_callback:
        progress_callback(0, 100)

    # 1. Đọc folder metadata → df_meta
    df_meta = read_meta_folder(meta_folder)
    if progress_callback:
        progress_callback(30, 100)

    # 2. Xử lý file SUM → df_sum
    df_sum = process_sum_file(sum_file, sheet="Sheet1")
    if progress_callback:
        progress_callback(60, 100)

    # 3. Tìm dòng meta KHÔNG match trong sum
    df_meta_only = find_meta_only(df_meta, df_sum)
    if progress_callback:
        progress_callback(85, 100)

    # 4. Chỉ xuất meta_only.xlsx nếu có dòng chưa match
    meta_only_path = None
    if not df_meta_only.empty:
        meta_only_path = os.path.join(output_folder, "meta_only.xlsx")
        df_meta_only.to_excel(meta_only_path, index=False)

    if progress_callback:
        progress_callback(100, 100)

    return {
        "meta_path":      None,
        "sum_path":       None,
        "meta_only_path": meta_only_path,
        "n_meta":         len(df_meta),
        "n_sum":          len(df_sum),
        "n_meta_only":    len(df_meta_only),
        "df_meta_only":   df_meta_only,
    }


def run_sample_import(source_mode, source_path, output_path,
                      nhat_ky_nam_path="", nhat_ky_bac_path="", goi_xn_path="",
                      progress_callback=None):
    import glob
    if source_mode == "file":
        files = [source_path]
    else:
        files = glob.glob(os.path.join(source_path, "*.xlsx")) + \
                glob.glob(os.path.join(source_path, "*.xls"))

    results = []
    all_j_errors   = []
    all_desc_errors = []
    total = len(files)

    for i, fp in enumerate(files, start=1):
        result = process_sample_import(
            source_path=fp,
            output_folder=output_path,
            nhat_ky_nam_path=nhat_ky_nam_path,
            nhat_ky_bac_path=nhat_ky_bac_path,
            goi_xn_path=goi_xn_path,
        )
        results.append(result)

        df_j = result.get("j_errors")
        if df_j is not None and not df_j.empty:
            all_j_errors.append(df_j)

        df_d = result.get("desc_errors")
        if df_d is not None and not df_d.empty:
            all_desc_errors.append(df_d)

        if progress_callback:
            progress_callback(i, total)

    j_report_path    = None
    desc_report_path = None

    if all_j_errors:
        df_j_report = pd.concat(all_j_errors, ignore_index=True)
        j_report_path = os.path.join(output_path, "warning_col_J_SampleProject.xlsx")
        df_j_report.to_excel(j_report_path, index=False)
        print(f"\n⚠ Cảnh báo cột J / Sample Project trống → {j_report_path}")
    else:
        df_j_report = pd.DataFrame()
        print("\n✓ Không có cảnh báo cột J / Sample Project trống.")

    if all_desc_errors:
        df_desc_report = pd.concat(all_desc_errors, ignore_index=True)
        desc_report_path = os.path.join(output_path, "warning_description.xlsx")
        df_desc_report.to_excel(desc_report_path, index=False)
        print(f"\n⚠ Description lệch bảng labcode → {desc_report_path}")
    else:
        df_desc_report = pd.DataFrame()
        print("\n✓ Tất cả Description khớp bảng labcode.")

    return {
        "file_results":       results,
        "j_error_report":     df_j_report,
        "j_report_path":      j_report_path,
        "desc_error_report":  df_desc_report,
        "desc_report_path":   desc_report_path,
    }


def run_check_desc(goi_xn_path: str, output_path: str) -> dict:
    return run_check_description(labcode_file=goi_xn_path, folder_path=output_path)


def run_backend(source_mode, source_path, template_path, output_path, sheet_template="Sample", progress_callback=None):
    if source_mode == "file":
        folder = os.path.dirname(source_path)
        files, stats = process_combine_files(
            folder=folder,
            template_path=template_path,
            output_folder=output_path,
            sheet_template=sheet_template,
            progress_callback=progress_callback,
            filter_file=os.path.basename(source_path),
        )
    else:
        files, stats = process_combine_files(
            folder=source_path,
            template_path=template_path,
            output_folder=output_path,
            sheet_template=sheet_template,
            progress_callback=progress_callback,
        )

    # T7 primer check tự động trên các file vừa xuất
    t7_result = None
    if files:
        try:
            t7_result = process_primer_t7(file_paths=files, output_folder=output_path)
        except Exception as e:
            t7_result = {"message": "error", "error": str(e),
                         "total": 0, "duplicates": 0, "valid": 0, "output_path": None}

    return {
        "files":              files,
        "collisions":         stats["total_collisions"],
        "collision_per_file": stats["collision_per_file"],
        "error_files":        stats["error_files"],
        "t7":                 t7_result,
    }

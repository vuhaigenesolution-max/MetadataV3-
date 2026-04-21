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


def run_seed_file(source_mode, source_path, template_path, output_path, progress_callback=None):
    import glob as _glob
    from Logic.SeedfilePage import process_excel

    if source_mode == "file":
        result = process_excel(source_path, template_path, output_path)
        if progress_callback:
            progress_callback(1, 1)
        return [result]

    files = _glob.glob(os.path.join(source_path, "*.xlsx")) + \
            _glob.glob(os.path.join(source_path, "*.xls"))
    results = []
    total = len(files)
    for i, fp in enumerate(files, start=1):
        csv_path = process_excel(fp, template_path, output_path)
        results.append(csv_path)
        if progress_callback:
            progress_callback(i, total)
    return results


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

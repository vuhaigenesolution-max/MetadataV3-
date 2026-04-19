import pandas as pd
import re
import glob
import os

NIPT_PACKAGES = {"TS1", "TS3", "TS95", "TS24", "TSPRO"}


def load_labcode_lookup(labcode_file: str) -> dict:
    xl = pd.ExcelFile(labcode_file)
    lookup = {}
    for sheet in xl.sheet_names:
        df = xl.parse(sheet, skiprows=1, header=0)
        col_labcode, col_so_x, *cols_goi = df.columns

        df_melt = df.melt(
            id_vars=[col_labcode, col_so_x],
            value_vars=cols_goi,
            var_name="Tên gói xét nghiệm",
            value_name="Đánh dấu",
        )
        df_result = df_melt[df_melt["Đánh dấu"] == "Y"].drop(columns="Đánh dấu")
        df_result["Nhóm xét nghiệm"] = df_result["Tên gói xét nghiệm"].apply(
            lambda x: "NIPT" if x in NIPT_PACKAGES else "CRH"
        )

        for _, row in df_result.iterrows():
            key = (str(row[col_labcode]).strip().upper(), str(row["Nhóm xét nghiệm"]).strip())
            lookup[key] = str(row["Tên gói xét nghiệm"]).strip()

    return lookup


def _extract_prefix(val, n: int = 3) -> str:
    if pd.isna(val):
        return ""
    return re.sub(r'[\d\-]', '', str(val).strip())[:n].upper()


def _get_nhom_xn(prefix: str) -> str:
    if prefix[:2] == "GS":
        return "SGNU"
    if prefix[:2] == "CR":
        return "CRH"
    if prefix[:1] in {"E", "T", "V", "H", "B", "L", "P"}:
        return "NIPT"
    return ""


def _extract_runname(filename: str) -> str:
    m = re.search(r'_(R\d+)_', filename)
    return m.group(1) if m else ""


def run_check_description(labcode_file: str, folder_path: str) -> dict:
    """
    labcode_file : đường dẫn file "Thông tin tên gói xét nghiệm.xlsx"  (goi_xn_path)
    folder_path  : folder chứa các file metadata OUTPUT cần kiểm tra   (output_path)

    Trả về dict:
        output_file  : đường dẫn file kết quả (hoặc None nếu không có lệch)
        error_rows   : số dòng lệch tổng
        file_results : list of {file, runname, errors} cho từng file
    """
    lookup = load_labcode_lookup(labcode_file)
    output_file = os.path.join(folder_path, "check_description.xlsx")

    files = [
        f for f in glob.glob(os.path.join(folder_path, "*.xlsx"))
        if os.path.basename(f) != "check_description.xlsx"
    ]

    all_errors = []
    file_results = []

    for fp in files:
        fname   = os.path.basename(fp)
        runname = _extract_runname(fname)
        try:
            df = pd.read_excel(fp, sheet_name="SampleImport", skiprows=22, header=0)
            df = df.dropna(axis=1, how="all").dropna(axis=0, how="all").reset_index(drop=True)

            if "Description" not in df.columns:
                file_results.append({"file": fname, "runname": runname, "errors": -1, "note": "no Description col"})
                continue

            col_sample_id    = df.columns[0]
            col_sample_plate = df.columns[2]

            df["Runname"]            = runname
            df["Đầu mã"]             = df[col_sample_id].apply(lambda x: _extract_prefix(x, 3))
            df["Đầu gói xét nghiệm"] = df[col_sample_plate].apply(lambda x: _extract_prefix(x, 3))
            df["Gói xét nghiệm"]     = df["Đầu gói xét nghiệm"].apply(_get_nhom_xn)

            def get_ten_xn(row):
                nhom = row["Gói xét nghiệm"]
                if nhom == "SGNU":
                    return "SGNU"
                result = lookup.get((row["Đầu mã"], nhom), "")
                if result == "":
                    return str(row["Description"]).strip() if not pd.isna(row["Description"]) else ""
                return result

            df["Description check"] = df.apply(get_ten_xn, axis=1)

            mask   = df["Description"].astype(str).str.strip() != df["Description check"].astype(str).str.strip()
            df_err = df[mask][[
                "Runname", col_sample_id, col_sample_plate, "Description", "Description check"
            ]].copy()

            n_err = len(df_err)
            file_results.append({"file": fname, "runname": runname, "errors": n_err})

            if n_err > 0:
                df_err.insert(0, "File", fname)
                all_errors.append(df_err)

        except Exception as e:
            file_results.append({"file": fname, "runname": runname, "errors": -1, "note": str(e)})

    if all_errors:
        df_out = pd.concat(all_errors, ignore_index=True)
        df_out.to_excel(output_file, index=False)
        return {"output_file": output_file, "error_rows": len(df_out), "file_results": file_results}

    return {"output_file": None, "error_rows": 0, "file_results": file_results}

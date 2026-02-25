import os
import pandas as pd
from openpyxl import load_workbook
import csv
from datetime import datetime
from SeedfilePage import process_excel

def process_folder(source_folder, template_path, output_path):
    """
    Quét toàn bộ file trong thư mục source_folder và thực thi process_excel cho từng file.
    """
    processed_files = []
    error_files = []
    for filename in os.listdir(source_folder):
        file_path = os.path.join(source_folder, filename)
        if os.path.isfile(file_path) and filename.lower().endswith(('.xlsx', '.xls')):
            try:
                print(f"Đang xử lý file: {file_path}")
                csv_path = process_excel(file_path, template_path, output_path)
                processed_files.append(csv_path)
            except Exception as e:
                msg = f"Lỗi khi xử lý {file_path}: {e}"
                print(msg)
                error_files.append(msg)
    if error_files:
        raise ValueError("Một hoặc nhiều file có lỗi:\n" + "\n".join(error_files))
    print(f"Đã xử lý {len(processed_files)} file. Kết quả lưu tại: {output_path}")
    return processed_files

# Ví dụ sử dụng (bỏ comment để chạy độc lập)
# source_folder = r"C:\path\to\your\folder"
# template_path = r"C:\path\to\template.xlsx"
# output_path = r"C:\path\to\output"
# process_folder(source_folder, template_path, output_path)

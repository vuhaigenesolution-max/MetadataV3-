# Metadata Tool V3 — Mô tả toàn bộ codebase

## Tổng quan

Metadata Tool V3 là ứng dụng desktop (Tkinter) xử lý file Excel/CSV cho quy trình chuẩn bị dữ liệu giải trình tự gen (NGS) tại Gene Solutions. Công cụ gồm hai lớp rõ ràng:

- **Logic** (`src/Logic/`): Xử lý dữ liệu thuần — không phụ thuộc UI
- **UI** (`src/UI/`): Giao diện Tkinter gọi vào lớp Logic

---

## Cấu trúc thư mục

```
Metadata_Tool V3/
├── main.py                         # Điểm khởi chạy ứng dụng
└── src/
    ├── Logic/
    │   ├── __init__.py
    │   ├── CombieFile.py           # Gộp file Excel + phát hiện collision
    │   ├── SampleImport.py         # Xuất SampleImport CSV + Manifest
    │   ├── SeedfilePage.py         # Tạo barcode.csv từ 1 file Excel
    │   ├── Folder_SeedfilePage.py  # Xử lý batch nhiều file SeedFile
    │   ├── CheckDescription.py     # Kiểm tra cột Description
    │   └── Primer_T7.py            # Kiểm tra primer trùng cột T7
    └── UI/
        ├── __init__.py
        ├── backend.py              # Điều phối tất cả workflow Logic
        ├── HomePage.py             # Trang chủ — điều hướng
        ├── CombinePage.py          # GUI bước 1: Combine File
        ├── SampleImportPage.py     # GUI bước 2: Tạo SampleImport & Manifest
        ├── SeedFilePage.py         # GUI: Tạo file mồi (barcode)
        └── primreT7.py             # GUI: Kiểm tra primer T7
```

---

## Lớp Logic

### `CombieFile.py`

**Mục đích**: Gộp nhiều file Excel nguồn thành các file metadata đầu ra, đồng thời phát hiện và đánh dấu xung đột index.

**Hàm chính**: `process_combine_files(source_paths, template_path, output_dir)`

**Luồng xử lý**:
1. Đọc từng file nguồn → trích xuất cột A–C, H–J, K–L, T (hàm `_read_source_file`)
2. Kiểm tra ký tự đặc biệt / khoảng trắng thừa (hàm `_validate_df`)
3. Nhóm các row theo `(runname, date)` → mỗi nhóm tạo 1 file đầu ra
4. Sao chép template → dán dữ liệu vào sheet SampleImport (hàm `_paste_to_template`)
5. So sánh chuỗi index trong cùng nhóm → tô vàng ô khác nhau < 3 ký tự, ghi log vào ô Q24+ (hàm `_detect_and_mark_collisions`)

**Đầu ra**: `metadata_<runname>_<date>.xlsx` cho mỗi nhóm run

---

### `SampleImport.py`

**Mục đích**: Đọc file metadata Excel đầu ra của CombieFile và xuất file CSV dùng cho hệ thống máy giải trình tự.

**Hàm chính**: `process_sample_import(source_path, output_dir, nhat_ky_nam, nhat_ky_bac, labcode_path)`

**Logic nghiệp vụ quan trọng**:
- Tự động nhận diện miền từ tiền tố runname: `R` = Nam/South, `P` = Bắc/North
- `_check_col_j_empty()`: Kiểm tra cột J bắt buộc cho kit T7/MGI (miền Nam)
- `_check_sample_project_empty()`: Kiểm tra Sample Project cho máy G99 (miền Bắc)
- `_check_description_mismatch()`: Đối chiếu cột Description với bảng labcode
- `_map_test_type()`: Chuẩn hóa tên loại xét nghiệm (ví dụ: `tsprocare` → `TSPRO`)
- `_compute_test_type()`: Suy ra loại xét nghiệm từ SampleID + mã xét nghiệm

**Đầu ra**:
- `SampleImport_<runname>.csv`
- `<runname>_YYYYMMDD.csv` (Aviti Manifest)
- File báo lỗi nếu có sai sót

---

### `SeedfilePage.py`

**Mục đích**: Tạo file `barcode.csv` từ một file Excel nguồn duy nhất.

**Hàm chính**: `process_excel(source_path, template_path, output_dir)`

**Luồng xử lý**:
1. Đọc ô T14 lấy tên run và tên barcode
2. Đọc cột T từ hàng 22 trở xuống lấy giá trị barcode
3. Tra cứu VLOOKUP trên sheet `PRIMER LOCKED` của template
4. Xuất file CSV với header `#barcodeName, #misMatch1, #misMatch2`

---

### `Folder_SeedfilePage.py`

**Mục đích**: Wrapper xử lý batch — áp dụng `SeedfilePage.process_excel()` cho toàn bộ file Excel trong một thư mục.

**Hàm chính**: `process_folder(folder_path, template_path, output_dir)`

**Đầu ra**: Danh sách đường dẫn CSV đã tạo, raise lỗi nếu bất kỳ file nào thất bại.

---

### `CheckDescription.py`

**Mục đích**: Kiểm tra cột Description trong các sheet SampleImport có khớp với bảng labcode chuẩn hay không.

**Hàm chính**: `run_check_description(folder_path, labcode_path)`

**Luồng xử lý**:
1. `load_labcode_lookup()`: Đọc file gói xét nghiệm, tạo dict tra cứu theo `(Runname, Nhóm XN)`
2. Duyệt tất cả `.xlsx` trong folder
3. So sánh từng row Description — ghi các dòng sai vào `check_description.xlsx`

**Đầu ra**: `check_description.xlsx` + thống kê số dòng lỗi theo từng file

---

### `Primer_T7.py`

**Mục đích**: Tìm primer trùng lặp giữa cột T (primer cần kiểm tra) và cột A (primer tham chiếu) qua nhiều file Excel.

**Hàm chính**: `process_primer_t7(file_paths, exclusion_path, output_dir)`

**Luồng xử lý**:
1. `_read_file()`: Đọc sheet `Sample`, trích cột A, B, C, D, T
2. So sánh từng primer trong cột T với tập primer cột A của các file khác
3. Ghi kết quả ra `Primer_Check_Result.xlsx` nếu có trùng

**Đầu ra**: Dict chứa tổng số, số trùng, số hợp lệ, đường dẫn file kết quả, thông báo.

---

## Lớp UI

### `backend.py`

**Mục đích**: Lớp trung gian — gọi Logic từ UI, chuẩn hóa tham số, gom kết quả cảnh báo.

**Hàm chính**:

| Hàm | Gọi Logic | Mô tả |
|-----|-----------|-------|
| `run_backend()` | `CombieFile` | Chạy toàn bộ workflow Combine |
| `run_sample_import()` | `SampleImport` | Batch xuất SampleImport, tổng hợp cảnh báo j_errors và desc_errors |
| `run_check_desc()` | `CheckDescription` | Kiểm tra Description |
| `run_seed_file()` | `Folder_SeedfilePage` | Batch tạo barcode CSV |

---

### `HomePage.py`

**Mục đích**: Trang chủ — giao diện điều hướng chính của ứng dụng.

**Giao diện**: 3 nút điều hướng trên nền tối (accent `#00c4a7`):
- **Combine File** → `CombinePage`
- **Tạo file SampleImport & Manifest** → `SampleImportPage`
- **Tạo file mồi** → `SeedFilePage`

---

### `CombinePage.py`

**Mục đích**: GUI bước 1 — người dùng chọn file/thư mục nguồn, template, thư mục đầu ra rồi chạy Combine.

**Tính năng UI đáng chú ý**:
- Thanh tiến trình với animation mượt (`_anim_tick()`)
- Tự động lưu config (đường dẫn, kết quả collision) vào file JSON
- Sau khi hoàn thành tự động mở cửa sổ `primreT7` để kiểm tra T7

**Persistence**: Lưu đường dẫn vào `last_paths.json`

---

### `SampleImportPage.py`

**Mục đích**: GUI bước 2 — xuất SampleImport CSV và Aviti Manifest.

**Input fields**:
- File/thư mục nguồn (metadata Excel đầu ra của Combine)
- Thư mục đích
- Nhật ký dò miền Nam *(tùy chọn)*
- Nhật ký dò miền Bắc *(tùy chọn)*
- File gói xét nghiệm / Labcode *(tùy chọn)*

**Tính năng**: Hiển thị cửa sổ kết quả tóm tắt sau khi hoàn thành; validate các file tùy chọn trước khi thực thi.

---

### `SeedFilePage.py`

**Mục đích**: GUI tạo file barcode CSV từ file Excel mồi.

**Input fields**:
- Chế độ: File đơn hoặc Thư mục
- Đường dẫn template
- Thư mục đầu ra
- Thanh tiến trình + hiển thị thời gian đã chạy

**Persistence**: Lưu đường dẫn vào `last_paths.json`

---

### `primreT7.py`

**Mục đích**: GUI kiểm tra primer trùng cột T7 — thường được mở tự động sau bước Combine.

**Tính năng UI nổi bật**:
- Danh sách file có checkbox cuộn được (Canvas-based) với tô màu hàng khi chọn
- Nút **Select All / Clear All / Refresh**
- Kết quả hiển thị trong cửa sổ tóm tắt riêng: thống kê collision, cảnh báo ký tự đặc biệt, kết quả T7
- Nút **Go back to Combie** để quay lại bước trước
- Auto-populate thư mục đầu ra từ config của `CombinePage`

---

## Luồng xử lý chính (End-to-End)

```
[Người dùng]
     │
     ▼
HomePage
     │
     ├──► CombinePage
     │         │  chọn nhiều file Excel nguồn + template
     │         ▼
     │    backend.run_backend()
     │         │
     │         ▼
     │    CombieFile.process_combine_files()
     │         │  → metadata_*.xlsx (theo nhóm run/date)
     │         │  → collision log trong Q24+
     │         ▼
     │    primreT7 (tự động mở)
     │         │  kiểm tra primer trùng cột T
     │         ▼
     │    Primer_T7.process_primer_t7()
     │         │  → Primer_Check_Result.xlsx
     │
     ├──► SampleImportPage
     │         │  chọn metadata Excel + file tham chiếu miền Nam/Bắc
     │         ▼
     │    backend.run_sample_import()
     │         │
     │         ▼
     │    SampleImport.process_sample_import()
     │         │  → SampleImport_*.csv
     │         │  → *_YYYYMMDD.csv (Aviti Manifest)
     │         │  → báo lỗi nếu thiếu dữ liệu
     │
     └──► SeedFilePage
               │  chọn file/thư mục Excel mồi + template
               ▼
          backend.run_seed_file()
               │
               ▼
          Folder_SeedfilePage → SeedfilePage.process_excel()
               │  → barcode.csv (mỗi file nguồn 1 CSV)
```

---

## Stack công nghệ

| Thư viện | Dùng cho |
|----------|----------|
| `pandas` | Đọc, lọc, xử lý dữ liệu bảng |
| `openpyxl` | Đọc/ghi Excel, tô màu ô, dán dữ liệu vào template |
| `tkinter` | Framework GUI desktop |
| `threading` | Chạy Logic trên thread riêng, giữ UI không bị đơ |
| `logging` | Ghi log nội bộ các module Logic |
| `re` | Kiểm tra ký tự đặc biệt, chuẩn hóa chuỗi |
| `json` | Lưu/đọc đường dẫn gần nhất (`last_paths.json`) |

---

## Thống kê

| Hạng mục | Số lượng |
|----------|----------|
| File Python | 13 |
| Module Logic | 6 |
| Trang UI | 5 + 1 backend + 1 homepage |
| Tổng dòng code | ~2,500 |

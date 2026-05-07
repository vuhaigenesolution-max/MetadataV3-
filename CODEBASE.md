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
├── pyrightconfig.json              # Config Pylance — khai báo extraPaths cho src/UI
└── src/
    ├── Logic/
    │   ├── __init__.py
    │   ├── CombieFile.py           # Gộp file Excel + phát hiện collision
    │   ├── SampleImport.py         # Xuất SampleImport CSV + Manifest
    │   ├── CheckDescription.py     # Kiểm tra cột Description
    │   ├── Primer_T7.py            # Kiểm tra primer trùng cột T7
    │   └── Checksamplenumber.py    # Đối soát file SUM ↔ folder metadata
    └── UI/
        ├── __init__.py
        ├── _theme.py               # Palette, fonts, ttk styles, last_paths helpers
        ├── _progress.py            # Widget SmoothProgress (bar + % + status)
        ├── app.py                  # Single-root container, lazy-init các Page
        ├── backend.py              # Điều phối tất cả workflow Logic
        ├── HomePage.py             # Trang chủ — điều hướng
        ├── CombinePage.py          # GUI bước 1: Combine File
        ├── SampleImportPage.py     # GUI bước 2: Tạo SampleImport & Manifest
        ├── SeedFilePage.py         # GUI: Check Sample Number (SUM ↔ metadata)
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
   - Bỏ qua giá trị kiểu số (int/float) — không flag số là lỗi
   - File báo lỗi `error_<runname>_<date>.xlsx`: in đủ 4 cột A B C T, **tô vàng ô bị lỗi**, cột "Loại lỗi" mô tả chi tiết
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
- `_read_header_row()` + `_rename_with_headers()`: Đọc header thật từ row 23 (SampleImport) / row 15 (Aviti) → CSV xuất ra có header thật thay vì A,B,C,...

**Đầu ra**:
- `SampleImport_<runname>.csv` (cho cả Nam và Bắc)
- `Manifest_<runname>_<yyyymmdd>.csv` (cho cả Nam và Bắc)
- File báo lỗi nếu có sai sót

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

### `Checksamplenumber.py`

**Mục đích**: Đối soát file SUM (Pool Allocation, multi-page) với folder chứa file metadata. Phát hiện sample trong metadata chưa có trong SUM.

**Hàm chính**:
- `process_sum_file(sum_path, sheet)` → DataFrame 6 cột
- `read_meta_folder(folder, sheet_name="Sample")` → DataFrame gộp metadata
- `find_meta_only(df_meta, df_sum)` → DataFrame các dòng meta chưa match

**Pipeline `process_sum_file`**:
1. `parse_file_by_pages()` → đọc các page (mỗi page 20 cột + 2 cột trống), trim merged cells
2. Lọc 3 cột `RunName, Col 3, Col 5`, drop dòng có ô trống
3. **Clean Col 3**: bỏ phần `(...)` và content bên trong (vd `ABC(XYZ)` → `ABC`)
4. Parse Col 5 → expand thành nhiều dòng:
   - `1-48` → 48 dòng `sample_order = 1..48`
   - `1-48/1, 44-46` → 48 dòng, đánh `Loại = "Loại"` cho số {1, 44, 45, 46}
   - khác → 1 dòng giữ nguyên
5. Phân loại Sample Type:
   - `NIPT`: Col 3 prefix ∈ {E,T,B,H,L,V,P + số ; EK,TK,CR,CRH,GS + số ; PGS ; PGTM}
   - `other`: còn lại

**Pipeline `read_meta_folder`**:
- Mỗi file `metadata_<R/P xxxx>_<yyyymmdd>.xlsx` trong folder:
  - RunName = regex từ tên file (R hoặc P + 4-5 số)
  - Header = Excel row 21, 4 cột A, B, U, V
  - Data = Excel row 22+, cùng 4 cột
- Concat tất cả thành 1 DataFrame, cột đầu = RunName

**Logic `find_meta_only`** — match 1 dòng meta nếu thoả 1 trong 2:
- **NIPT**: tồn tại df_sum có `Sample Type=NIPT` & `Loại=""` & cùng RunName & `Col 3 == expNum` & `sample_order == sampleOrder`
- **other**: tồn tại df_sum có `Sample Type=other` & cùng RunName & `sample_order == sampleOrder` & `Col 3 ==` (Col_22 nếu khác rỗng, ngược lại Col_21)

Output: 3 cột `RunName, expNum, sampleOrder` — các dòng meta KHÔNG match.

---

## Lớp UI

### `backend.py`

**Mục đích**: Lớp trung gian — gọi Logic từ UI, chuẩn hóa tham số, gom kết quả cảnh báo.

**Hàm chính**:

| Hàm | Gọi Logic | Mô tả |
|-----|-----------|-------|
| `run_backend()` | `CombieFile` + `Primer_T7` | Combine + auto T7 check |
| `run_sample_import()` | `SampleImport` | Batch xuất SampleImport, tổng hợp cảnh báo j_errors / desc_errors |
| `run_check_desc()` | `CheckDescription` | Kiểm tra Description |
| `run_check_sample_number()` | `Checksamplenumber` | Đối soát SUM ↔ metadata folder |

---

### `HomePage.py`

**Mục đích**: Trang chủ — giao diện điều hướng chính của ứng dụng.

**Giao diện**: 3 nút điều hướng trên nền tối (accent `#00c4a7`):
- **Combine File** → `CombinePage`
- **Tạo file SampleImport & Manifest** → `SampleImportPage`
- **Check Sample Import** → `SeedFilePage`

---

### `CombinePage.py`

**Mục đích**: GUI bước 1 — người dùng chọn file/thư mục nguồn, template, thư mục đầu ra rồi chạy Combine.

**Tính năng UI đáng chú ý**:
- Thanh tiến trình với animation mượt (`_anim_tick()`)
- Tự động lưu config (đường dẫn, kết quả collision) vào file JSON
- Sau khi hoàn thành tự động mở cửa sổ `primreT7` để kiểm tra T7

**Persistence**: Lưu đường dẫn vào `last_paths.json` — keys: `combine_source_mode`, `combine_source_path`, `combine_template_path`, `combine_output_path`

---

### `SampleImportPage.py`

**Mục đích**: GUI bước 2 — xuất SampleImport CSV và Aviti Manifest.

**Input fields**:
- File/thư mục nguồn (metadata Excel — auto-fill từ `output_path` của Combine khi chưa chọn riêng)
- Thư mục đích (riêng biệt)
- Nhật ký dò miền Nam *(tùy chọn)*
- Nhật ký dò miền Bắc *(tùy chọn)*
- File gói xét nghiệm / Labcode *(tùy chọn)*

**Persistence**: Dùng key namespace riêng để không đè lên config của CombinePage:
`sample_source_mode`, `sample_source_path`, `sample_output_path`,
`sample_nhat_ky_nam_path`, `sample_nhat_ky_bac_path`, `sample_goi_xn_path`.

**Tính năng**:
- Progress bar **thực tế** theo số file (không còn fake animation)
- Popup tổng kết hiển thị: số file CSV đã xuất + đường dẫn warning files (cột J / Description)

---

### `SeedFilePage.py` *(đã repurpose: Check Sample Number)*

**Mục đích**: GUI đối soát file SUM với folder metadata.

**Input fields** (3 row):
1. **Source Metadata**: folder chứa nhiều file metadata
2. **File Sum**: file SUM Excel
3. **Destination**: folder đầu ra

**Backend**: gọi `run_check_sample_number(meta_folder, sum_file, output_folder)`

**Output 3 file Excel**:
- `metadata_combined.xlsx`: gộp tất cả metadata
- `<sum_stem>_solution.xlsx`: SUM đã processed (RunName, Col 3, Col 5, sample_order, Loại, Sample Type)
- `meta_only.xlsx`: 3 cột (RunName, expNum, sampleOrder) — sample có trong meta nhưng chưa có trong sum

**Persistence**: Lưu vào `last_paths.json` — keys: `check_meta_path`, `check_filesum_path`, `check_output_path`

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
     │         │  chọn file/folder Excel nguồn + template
     │         ▼
     │    backend.run_backend()
     │         ├─► CombieFile.process_combine_files()
     │         │        → metadata_*.xlsx (theo nhóm run/date)
     │         │        → tô vàng cell collision + log Q24+
     │         │        → error_*.xlsx (tô vàng ô lỗi)
     │         └─► Primer_T7.process_primer_t7() (auto chain)
     │                  → Primer_Check_Result.xlsx
     │
     ├──► SampleImportPage
     │         │  Source = output Combine (auto-fill)
     │         │  + Destination riêng + ref miền Nam/Bắc + Labcode
     │         ▼
     │    backend.run_sample_import()
     │         └─► SampleImport.process_sample_import()
     │                  → SampleImport_<runname>.csv  (header thật)
     │                  → Manifest_<runname>_<yyyymmdd>.csv
     │                  → warning_col_J_SampleProject.xlsx (gộp toàn batch)
     │                  → warning_description.xlsx (gộp toàn batch)
     │
     └──► SeedFilePage  *(Check Sample Number)*
               │  Source Metadata (folder) + File Sum + Destination
               ▼
          backend.run_check_sample_number()
               ├─► Checksamplenumber.read_meta_folder()
               │        → df_meta gộp (RunName + 4 cột A/B/U/V row 21+)
               ├─► Checksamplenumber.process_sum_file()
               │        → df_sum (clean Col 3, expand Col 5, classify NIPT/other)
               └─► Checksamplenumber.find_meta_only()
                        → df_meta_only (rows meta không match sum)
               
               Output 3 file Excel:
                   metadata_combined.xlsx
                   <sum_stem>_solution.xlsx
                   meta_only.xlsx
```

---

## Stack công nghệ

| Thư viện | Dùng cho |
|----------|----------|
| `pandas` | Đọc, lọc, transform, explode, merge dữ liệu bảng |
| `openpyxl` | Đọc/ghi Excel, tô màu ô, dán dữ liệu vào template, expand merged cells |
| `tkinter` | Framework GUI desktop |
| `threading` | Chạy Logic trên thread riêng, giữ UI không bị đơ |
| `logging` | Ghi log nội bộ các module Logic |
| `re` | Regex: parse Col 5 range, classify NIPT, validate ký tự đặc biệt |
| `json` | Lưu/đọc đường dẫn gần nhất (`last_paths.json`) |

---

## Workflow `Checksamplenumber` — chi tiết

### File SUM input

Multi-page Excel layout:
- Mỗi page = **20 cột data + 2 cột trống** (STEP = 22)
- Mỗi page có:
  - Row 1 meta: A1=RunContent, E1=RunNameNotes, Q1=MachineNotes
  - Row 4: header bảng (sẽ thay bằng `Col 1`..`Col 20`)
  - Row 5+: data, trim theo cột C của page

### File metadata input

Filename: `metadata_<R/P xxxx>_<yyyymmdd>.xlsx`
- RunName extract: regex `metadata_([RP]\d{4,5})_(\d{8})`
- Sheet: `Sample`
- Header: row 21, lấy 4 cột Excel A, B, U, V (KHÔNG phải tên cột — vị trí Excel)
- Data: row 22 trở đi, cùng 4 cột

### Match rules (find_meta_only)

| df_sum row | Condition để được coi là "matched" với df_meta row |
|---|---|
| `Sample Type = NIPT, Loại = ""` | `RunName` cùng & `Col 3 == expNum` (col A) & `sample_order == sampleOrder` (col B) |
| `Sample Type = other` | `RunName` cùng & `sample_order == sampleOrder` & `Col 3 ==` (Col_22 col V nếu khác rỗng, ngược lại Col_21 col U) |

Comparison đã chuẩn hóa: `strip()` + bỏ `.0` đuôi (1 vs 1.0 coi như bằng).

---

## Thống kê

| Hạng mục | Số lượng |
|----------|----------|
| File Python | 12 |
| Module Logic | 5 |
| Trang UI | 5 + 1 backend + 1 homepage |
| Tổng dòng code | ~3,000 |

---

## Widget chia sẻ (`_progress.py`)

`SmoothProgress` (Frame): bar tween mượt + label `%` + label status text.
- API: `reset()`, `set_target(pct, status=None)`, `set_status(text)`, `finish(status="✓ Hoàn tất")`
- Easing: tăng 18% khoảng cách / tick 20ms
- Dùng chung trên: CombinePage, SampleImportPage, SeedFilePage, primreT7

---

## Pylance / IDE config (`pyrightconfig.json`)

Ở root project, khai báo `extraPaths` cho Pylance hiểu việc `main.py` đẩy `src/UI` vào `sys.path` lúc runtime:

```json
{
  "extraPaths": ["src", "src/UI"]
}
```

Nhờ đó các import `from app import App`, `from HomePage import HomePage`, `from _theme import …` không bị gạch chân đỏ.

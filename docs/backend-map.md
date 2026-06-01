# Bản đồ backend `00.Detaisublog_v26.py`

## 1. Tổng quan

`00.Detaisublog_v26.py` là backend chính của công cụ export-agsh. File này chịu trách nhiệm:

- Quét thư mục nguồn để tìm file Excel phù hợp.
- Đọc các file `.xls`, `.xlsx`, `.xlsm`.
- Tìm sheet PL và INV, đọc metadata và các dòng chi tiết.
- Kiểm tra sai lệch dữ liệu.
- Ghi file Excel tổng hợp.
- Có thể sửa trực tiếp metadata trong file nguồn nếu bật repair.
- Dùng cache và state để giảm thời gian chạy lại.
- Gọi `folder_audit_ext.py` để cập nhật sheet `Folder` và các cột audit bổ sung.

File hiện có khoảng 15.000 dòng. Đây không còn là một module đơn giản mà là tập hợp nhiều phiên bản nâng cấp được bổ sung nối tiếp.

## 2. Vì sao file nặng và bị vá chồng

Backend được phát triển theo cách giữ logic cũ rồi thêm bản nâng cấp ở phía dưới file. Một số hàm, biến và class được định nghĩa lại nhiều lần.

Ví dụ:

```python
class Processor:
    ...

class Processor(Processor):
    ...

class Processor(Processor):
    ...
```

Class mới kế thừa class có cùng tên đã tồn tại trước đó. Vì vậy phần code phía trên không thể bị coi là code chết chỉ vì đã có class mới phía dưới.

Ngoài ra còn có các kiểu override khác:

```python
CORE_RUN_V6 = Processor.run
CORE_RUN_V6 = _u90_core_run
CORE_RUN_V6 = _u92_core_run

_v72_build_detail_bundle = _u108_build_detail_bundle_safe
```

Khi đọc file, cần xem toàn bộ thứ tự từ trên xuống dưới. Giá trị cuối cùng mới là giá trị có hiệu lực khi module được nạp.

## 3. Chuỗi `class Processor(Processor)`

`Processor` ban đầu được khai báo gần dòng 1488. Sau đó file bổ sung nhiều lớp kế thừa nối tiếp, gồm các mốc chính:

| Vùng gần dòng | Vai trò |
|---|---|
| `1488` | `Processor` nền ban đầu |
| `10875` | Runtime được gom lại rõ hơn |
| `11213` đến `13706` | Các lớp nâng cấp cache, fingerprint, upsert và tối ưu |
| `14468` | Lớp `Processor` cuối đang có hiệu lực |

Lớp cuối tại gần dòng `14468` thực hiện luồng chính:

```text
Nạp cache sidecar
→ Quét file nguồn
→ So sánh fingerprint và state
→ Chỉ parse file cần xử lý
→ Quyết định có cần mở workbook output không
→ Cập nhật SUB_DETAIL / INV / PL
→ Gọi folder_audit_ext.py nếu cần
→ Lưu workbook và cache/state
```

### Lưu ý an toàn

Không xóa các lớp `Processor` cũ trước khi chứng minh rằng:

1. Lớp cuối không kế thừa method từ lớp cũ.
2. Không còn biến trung gian như `CORE_RUN_V6` trỏ vào method cũ.
3. Test output trước và sau khi xóa vẫn tương đương.

## 4. Entrypoint từ PySide6 UI

Giao diện mới nằm trong `260423.Sub_UIv17.py`.

Luồng gọi:

```text
260423.Sub_UIv17.py
→ load_backend_module()
→ nạp 00.Detaisublog_v26.py bằng importlib
→ backend.Processor(logger)
→ processor.run(folder_path, output_path, repair_options=...)
```

Các điểm neo:

- `BACKEND_PATH`: gần dòng `23` của UI.
- `load_backend_module()`: gần dòng `33`.
- Khởi tạo `backend.Processor`: gần dòng `451`.
- Tạo `backend.RepairOptions`: gần dòng `1179`.

PySide6 UI là giao diện nên ưu tiên sử dụng, nhưng backend vẫn phải giữ API:

```python
Processor(logger)
RepairOptions(...)
Processor.run(...)
```

## 5. Entrypoint Tkinter cũ

Backend vẫn chứa UI Tkinter cũ:

- `class MainWindow`: gần dòng `1665`.
- `main()`: gần dòng `2090`.
- Entrypoint cuối file: gần dòng `15279`.

Khi chạy trực tiếp:

```powershell
python "00.Detaisublog_v26.py"
```

backend gọi `main()` và mở UI Tkinter cũ.

### Lưu ý an toàn

Chưa được xóa UI Tkinter cũ. Cần chứng minh rõ:

- Không còn người dùng chạy trực tiếp backend.
- PySide6 UI thay thế đầy đủ chức năng cần thiết.
- Có phương án giữ tương thích hoặc thông báo chuyển đổi.

## 6. Nhóm logic đọc Excel

Các helper nền nằm chủ yếu ở đầu file:

| Hàm | Vai trò |
|---|---|
| `detect_file_type()` | Phân biệt `.xls` với `.xlsx` / `.xlsm` |
| `load_workbook_data()` | Đọc workbook nguồn thành cấu trúc dữ liệu dùng chung |
| `find_best_pl_sheet()` | Chọn sheet PL phù hợp |
| `find_best_inv_sheet()` | Chọn sheet INV hoặc INVOICE phù hợp |
| `detect_pl_header()` | Tìm header PL |
| `detect_inv_header()` | Tìm header INV |
| `extract_pl_total()` | Đọc tổng PL |
| `extract_inv_total()` | Đọc tổng INV |
| `extract_sheet_meta()` | Đọc metadata invoice |
| `process_one_sub_file()` | Xử lý một file nguồn |

Engine đọc file:

- `openpyxl`: dùng cho `.xlsx`, `.xlsm`.
- `xlrd`: dùng cho `.xls`.

Phần cuối file còn có các parser mở rộng cho nhiều họ file, ví dụ SA, SUB, VSF. Override cuối:

```python
_v72_build_detail_bundle = _u108_build_detail_bundle_safe
```

Nó giúp một file SA lỗi header không làm dừng toàn bộ lượt chạy.

## 7. Nhóm logic ghi output

Output là workbook `.xlsx` tổng hợp. Các sheet chính:

```text
SUB_DETAIL
LOG_SUB_DETAIL
INV
PL
Folder
```

Các hàm quan trọng:

| Hàm | Vai trò |
|---|---|
| `open_or_create_output_workbook()` | Mở output hiện có hoặc tạo workbook mới |
| `safe_save_workbook_atomic()` | Lưu tạm, tạo backup rồi thay thế output |
| `ensure_sheet()` | Tạo sheet nếu chưa tồn tại |
| `append_log_rows()` | Ghi log vào workbook |
| Các helper `_u104_*`, `_u105_*`, `_u136_*` | Upsert, lập index, xóa và ghi lại dòng thay đổi |

`safe_save_workbook_atomic()` gần dòng `1297` là vùng nhạy cảm:

```text
Lưu workbook vào file tạm
→ nếu output cũ tồn tại thì tạo file .bak
→ os.replace() file tạm vào output
→ nếu output đang bị khóa thì lưu một bản thay thế để tránh mất dữ liệu
```

Sau đó hàm này được bọc lại gần dòng `6902` để đo benchmark. Không sửa hoặc xóa wrapper khi chưa kiểm tra luồng runtime.

## 8. Nhóm repair file nguồn

Repair không chỉ sửa báo cáo. Nó có thể ghi đè file Excel nguồn.

Các vùng chính:

| Hàm | Vai trò |
|---|---|
| `RepairOptions` | Cấu hình bật/tắt repair |
| `get_repair_truth()` | Xác định dữ liệu chuẩn từ folder hoặc nhập tay |
| `set_meta_value_to_ws()` | Ghi giá trị vào ô metadata |
| `repair_excel_metadata()` | Mở và lưu đè file nguồn |

`repair_excel_metadata()` gần dòng `1397`:

- Chỉ tự sửa `.xlsx`, `.xlsm`.
- Có thể sửa `invoice_no`, `invoice_date`, `destination`.
- Lưu trực tiếp bằng `wb.save(file_path)`.
- `.xls` chỉ được báo để xử lý tay.

### Rủi ro

Repair cần được xem như chức năng ghi dữ liệu thật. Khi test:

- Chỉ dùng bản sao thư mục nguồn.
- Kiểm tra rõ repair bật hay tắt.
- Không mặc định coi chạy backend là thao tác chỉ đọc.

## 9. Nhóm cache, log và state

Backend dùng nhiều sidecar cạnh file output để tăng tốc và giữ trạng thái giữa các lần chạy.

| Nhóm | Ví dụ tên file | Vai trò |
|---|---|---|
| Width cache | `*.widthcache.json` | Ghi nhớ độ rộng cột |
| Source cache cũ | `*.sourcecache.pkl` | Cache dữ liệu đọc nguồn |
| Source cache U90 | `*.sourcecache_u90.pkl` | Cache parser nâng cấp |
| Parse cache U130 | `*.sourcecache_u130.pkl` | Cache kết quả parse |
| Fingerprint cache | `*.fingerprintcache*.json` | Xác định file nguồn thay đổi |
| Detail index | `*.detailindex_u131.pkl` | Ánh xạ file nguồn với dòng INV/PL |
| Substate | `*.substate_u134.json` | Trạng thái xử lý theo file |
| Issue snapshot | `*.issue_snapshot_u139.json` | Danh sách lỗi để khôi phục UI nhanh |
| Benchmark log | `*.runtime_bench.log` | Log đo thời gian xử lý |
| Runtime log | `detaisublog_runtime.log` | Log kỹ thuật |

Các helper quan trọng:

```text
_u90_source_cache_path()
_u130_parse_cache_path()
_u131_detail_index_path()
_u134_substate_path()
_u139_issue_snapshot_path()
```

### Rủi ro

Cache không chỉ là tối ưu tốc độ. Một số cache tham gia quyết định:

- File nào cần parse lại.
- Dòng output nào cần cập nhật hoặc xóa.
- Có cần mở và lưu workbook không.
- Danh sách lỗi nào hiển thị lại trên UI.

Không xóa hoặc đổi schema cache trong cùng commit với refactor parser.

## 10. Chỗ gọi `folder_audit_ext.py`

Backend import động module audit trong `Processor.run()` cuối, gần dòng `14693`:

```python
from folder_audit_ext import (
    build_delta_state,
    apply_delta_inplace,
    save_cache,
)
```

Luồng tích hợp:

```text
build_delta_state(folder_path, output_path)
→ xác định thay đổi thư mục và cache audit
→ quyết định có cần mở workbook không
→ apply_delta_inplace(wb_out, folder_delta_prebuilt)
→ cập nhật sheet Folder và các cột audit
→ save_cache(...) nếu phù hợp
```

`folder_audit_ext.py` cập nhật thêm:

- Sheet `Folder`.
- `Cus no`, `Cus no.1`, `Cus no.-`.
- `S. Invoice#`, `S.DateWInv#`.
- `Cus no.-STT`.
- `Vas`, `HC`.

### Rủi ro

Không tách hoặc đổi thứ tự gọi folder audit nếu chưa kiểm tra:

- Chạy lần đầu.
- Chạy lại khi không đổi file nguồn.
- Thêm file nguồn.
- Sửa file nguồn.
- Xóa file nguồn.
- Một invoice có nhiều Cus no.

## 11. Danh sách vùng không được xóa vội

Các vùng sau chỉ được xóa sau khi có bằng chứng không còn được gọi và có test tương đương:

1. Chuỗi `class Processor(Processor)`.
2. Các lần gán `CORE_RUN_V6`.
3. Override `_v72_build_detail_bundle`.
4. UI Tkinter cũ và `main()` trong backend.
5. `repair_excel_metadata()` và các helper repair.
6. `safe_save_workbook_atomic()` nền và wrapper benchmark.
7. Các helper cache/state cũ và mới.
8. Logic upsert, xóa dòng theo path key và detail index.
9. Import động `folder_audit_ext.py`.
10. Các parser SA, SUB, VSF và fallback header.

Quy tắc xóa code:

```text
Không còn gọi trực tiếp
AND không còn được kế thừa hoặc tham chiếu gián tiếp
AND test characterization vẫn pass
```

## 12. Kế hoạch làm sạch an toàn

### Giai đoạn 1: Lập bản đồ

Mục tiêu:

- Hoàn thành call graph runtime.
- Liệt kê điểm đọc và ghi file.
- Ghi rõ override cuối của từng nhóm hàm.

Phạm vi:

- Chỉ sửa tài liệu.
- Không sửa code.

Điều kiện commit:

- Sơ đồ chỉ rõ entrypoint PySide6, Tkinter, `Processor.run()`, repair, cache và folder audit.

### Giai đoạn 2: Tài liệu hóa nghiệp vụ

Mục tiêu:

- Ghi rõ schema output.
- Mô tả repair và cảnh báo sửa file nguồn.
- Ghi danh sách sidecar cache/log.
- Chuẩn hóa checklist test bằng dữ liệu sao chép.

Phạm vi:

- `README.md`.
- Thư mục `docs/`.

Điều kiện commit:

- Tài liệu khớp code hiện tại, chưa thay đổi hành vi.

### Giai đoạn 3: Dọn phần phụ không đổi logic

Mục tiêu:

- Giảm nhiễu khi đọc file.

Việc phù hợp:

- Gom import trùng đã được chứng minh.
- Sửa comment lỗi encoding.
- Thêm tiêu đề phân vùng.
- Đánh dấu vùng legacy candidate.

Chưa làm:

- Không xóa class cũ.
- Không xóa UI Tkinter.
- Không đổi helper đọc/ghi Excel.

Điều kiện commit:

- Diff không thay đổi hành vi.
- Có review tĩnh và smoke test phù hợp trước khi merge.

### Giai đoạn 4: Sửa lỗi logic cụ thể

Mục tiêu:

- Mỗi commit xử lý đúng một lỗi nghiệp vụ có dữ liệu tái hiện.

Quy trình:

```text
Tạo fixture từ bản sao dữ liệu
→ ghi nhận output sai
→ sửa phạm vi nhỏ nhất
→ so sánh output trước/sau
→ chạy lại để kiểm tra cache
```

Ví dụ:

- Invoice có nhiều Cus no.
- File SA có header không chuẩn.
- Xóa file nguồn nhưng output còn dòng cũ.

Điều kiện commit:

- Có ví dụ tái hiện và kết quả mong đợi.
- Không trộn refactor vào commit sửa lỗi.

### Giai đoạn 5: Refactor nhỏ có test

Mục tiêu:

- Giảm độ phức tạp từng phần mà không đổi output.

Thứ tự ưu tiên:

1. Tách manifest đường dẫn cache/state.
2. Tách adapter ghi output Excel.
3. Tách repair thành module rõ ràng và yêu cầu opt-in.
4. Tách Tkinter UI cũ sang file riêng nếu vẫn cần giữ.
5. Rút gọn chuỗi `Processor` từng bước sau khi có characterization test.

Checklist test:

- Chạy lần đầu với output mới.
- Chạy lại khi không đổi dữ liệu.
- Thêm, sửa, xóa file nguồn.
- Repair bật và repair tắt.
- Output đang mở trong Excel.
- Invoice có nhiều Cus no.
- UI PySide6 và entrypoint Tkinter cũ.

Điều kiện commit:

- Mỗi commit nhỏ và có thể rollback.
- Test xác nhận output và cache/state tương thích.
- Không xóa code chỉ dựa trên suy đoán.

---

Tài liệu này là bản đồ đọc code tĩnh. Nó không xác nhận kết quả runtime cho đến khi có bộ test trên bản sao dữ liệu.

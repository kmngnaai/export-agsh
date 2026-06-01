# Call graph runtime wrapper

## 1. Mục tiêu tài liệu

Tài liệu này lập bản đồ các wrapper runtime trong `00.Detaisublog_v26.py`, đặc biệt là:

- `Processor.run()`.
- Biến global `CORE_RUN_V6`.
- Các wrapper từ V6 đến U103.
- Các biến trung gian dạng `*_base`.

Mục tiêu là phân biệt luồng active hiện tại với chuỗi wrapper legacy để tránh xóa hoặc đổi thứ tự code theo suy đoán.

Đây là kết quả phân tích tĩnh từ code. Tài liệu chưa chứng minh hành vi runtime bằng bộ test đầy đủ.

## 2. Kết luận chính

- Runtime active hiện tại từ PySide6 UI đi vào `Processor.run()` cuối, gần dòng `14525`.
- Theo phân tích tĩnh, luồng PySide6 UI không thấy đi qua chuỗi wrapper hàm `_u84_processor_run` đến `_u103_processor_run`.
- Chuỗi wrapper dùng `CORE_RUN_V6` vẫn là legacy runtime rủi ro cao. Chưa được xóa code.
- Không được coi wrapper legacy là code chết chỉ vì UI hiện tại không gọi trực tiếp. Một số helper, class kế thừa hoặc entrypoint khác vẫn có thể phụ thuộc gián tiếp.

## 3. Call graph active

Luồng active từ giao diện PySide6:

```text
260423.Sub_UIv17.py
→ load_backend_module()
→ backend.Processor(logger)
→ processor.run(...)
→ Processor.run() cuối, gần dòng 14525
→ load cache sidecar
→ scan file nguồn
→ precheck fingerprint / state
→ parse candidate
→ upsert SUB_DETAIL / INV / PL
→ gọi folder_audit_ext.py
→ save workbook / cache / state
```

Điểm quan trọng:

- `Processor` cuối kế thừa các lớp `Processor(Processor)` phía trên.
- `Processor.run()` cuối có control flow riêng.
- Không thấy method `run()` cuối gọi `CORE_RUN_V6`.

## 4. Call graph legacy dùng `CORE_RUN_V6`

Luồng khái quát:

```text
wrapper legacy
→ CORE_RUN_V6(...)
→ lookup biến global tại runtime
→ _u92_core_run()
```

`CORE_RUN_V6` được gán ba lần theo thứ tự:

```python
CORE_RUN_V6 = Processor.run
CORE_RUN_V6 = _u90_core_run
CORE_RUN_V6 = _u92_core_run
```

Giá trị cuối cùng sau khi module được nạp đầy đủ:

```python
CORE_RUN_V6 = _u92_core_run
```

Ví dụ nhánh wrapper gián tiếp:

```text
_u90_processor_run()
→ _u90_processor_run_base
→ _u87_processor_run()
→ CORE_RUN_V6(...)
→ _u92_core_run()
```

## 5. Timeline wrapper V6 đến U83

| Wrapper | Dòng gần đúng | Gọi hàm chính | Chụp base wrapper | Vai trò | Rủi ro nếu xóa |
|---|---:|---|---|---|---|
| `_v6_processor_run` | `4225` | `CORE_RUN_V6`, `_v6_rebuild_detail_sheets` | Không | Rebuild detail theo V6. | Cao |
| `_v72_processor_run` | `4448` | `CORE_RUN_V6`, `_v72_rebuild_detail_sheets` | Không | Rebuild detail theo V7.2. | Cao |
| `_v73_processor_run` | `4733` | `CORE_RUN_V6`, `_v73_rebuild_detail_sheets` | Không | Bổ sung logic cột INV. | Cao |
| `_v74_processor_run` | `5190` | `CORE_RUN_V6`, `_v74_rebuild_detail_sheets_append` | Không | Chuyển sang append detail. | Cao |
| `_v75_processor_run` | `5338` | `CORE_RUN_V6`, snapshot và restore detail tabs | Không | Bảo vệ dữ liệu INV/PL cũ khi core chạy. | Cao |
| `_v76_processor_run` | `5473` | `CORE_RUN_V6`, snapshot và restore detail tabs | Không | Nâng cấp logic detail theo V7.6. | Cao |
| `_m3_processor_run` | `5582` | `CORE_RUN_V6`, `_m3_rebuild_merged_detail_rows` | Không | Merge detail INV/PL. | Cao |
| `_u79_processor_run` | `6235` | `CORE_RUN_V6`, rebuild merged detail | Không | Rebuild theo tập path OK và WARNING. | Cao |
| `_u82_processor_run` | `6374` | `CORE_RUN_V6`, `_u82_rebuild_merged_detail_rows` | Không | Tối ưu rebuild detail. | Cao |
| `_u83_processor_run` | `6534` | `CORE_RUN_V6`, `_u83_rebuild_merged_detail_rows` | Không | Tối ưu write/reset detail. | Cao |

## 6. Timeline wrapper U84 đến U103

| Wrapper | Dòng gần đúng | Gọi hàm chính | Chụp base wrapper | Vai trò | Rủi ro nếu xóa |
|---|---:|---|---|---|---|
| `_u84_processor_run` | `6889` | `CORE_RUN_V6`, rebuild detail | Không | Thêm width cache. | Cao |
| `_u85_processor_run` | `7038` | `CORE_RUN_V6`, rebuild detail, benchmark | Có biến local `_u84_processor_run_base`, nhưng không thấy dùng | Thêm benchmark. | Cao |
| `_u86_processor_run` | `7076` | `_u86_processor_run_base` | `_u85_processor_run` | Thêm marker log rồi chuyển tiếp. | Trung bình |
| `_u87_processor_run` | `7289` | `CORE_RUN_V6`, smart rebuild | Không | Thêm detail cache và smart rebuild. | Cao |
| `_u88_processor_run` | `7531` | `_u88_processor_run_base` | `_u87_processor_run` | Thêm source parse cache U88. | Cao |
| `_u90_processor_run` | `7901` | `_u90_processor_run_base` | `_u87_processor_run` | Thêm source cache U90. | Cao |
| `_u94_processor_run` | `8085` | `CORE_RUN_V6`, smart rebuild delta | Không | Tính delta đúng thời điểm. | Cao |
| `_u95_processor_run` | `8536` | `_u94_processor_run` | Không | Chuyển tiếp sau monkey-patch rebuild U95. | Cao |
| `_u96_processor_run` | `8902` | `_u94_processor_run` | Không | Chuyển tiếp sau monkey-patch block surgery U96. | Cao |
| `_u97_processor_run` | `9090` | `CORE_RUN_V6`, append-only rebuild | Không | Đổi chiến lược ghi INV/PL. | Cao |
| `_u98_processor_run` | `9537` | `CORE_RUN_V6`, file upsert | Không | Upsert theo file. | Cao |
| `_u99_processor_run` | `9894` | `CORE_RUN_V6`, upsert U99 | Không | Sửa cách xác định changed paths. | Cao |
| `_u100_processor_run` | `9995` | `CORE_RUN_V6`, upsert | Không | Bootstrap khi INV/PL rỗng. | Cao |
| `_u101_processor_run` | `10087` | `CORE_RUN_V6`, upsert U99 | Không | Sửa bootstrap và index. | Cao |
| `_u102_processor_run` | `10182` | `CORE_RUN_V6`, upsert U99 | Không | Sửa current-folder path match. | Cao |
| `_u103_processor_run` | `10264` | `CORE_RUN_V6`, upsert U99 | Không | Ưu tiên paths trả về từ core summary. | Cao |

## 7. Bảng biến base capture

| Dòng gần đúng | Biến | Giá trị được chụp | Ý nghĩa |
|---:|---|---|---|
| `7074` | `_u86_processor_run_base` | `_u85_processor_run` | U86 thêm log marker rồi gọi wrapper U85. |
| `7528` | `_u88_processor_run_base` | `_u87_processor_run` | U88 bọc smart rebuild U87 bằng source parse cache U88. |
| `7899` | `_u90_processor_run_base` | `_u87_processor_run` | U90 bọc smart rebuild U87 bằng source cache U90. |

Lưu ý:

- Các biến trên chụp wrapper, không chụp trực tiếp giá trị `CORE_RUN_V6`.
- Khi wrapper được chụp gọi `CORE_RUN_V6(...)`, nó vẫn lookup giá trị global tại runtime.

## 8. Wrapper có vẻ chỉ chuyển tiếp

Các vùng sau có vẻ ít logic hơn các wrapper lân cận, nhưng chưa có đủ bằng chứng để xóa:

| Vùng | Nhận xét |
|---|---|
| `_u86_processor_run` | Thêm một log marker rồi gọi `_u86_processor_run_base`. |
| `_u95_processor_run` | Chỉ gọi `_u94_processor_run`, nhưng đi kèm monkey-patch rebuild U95 phía trước. |
| `_u96_processor_run` | Chỉ gọi `_u94_processor_run`, nhưng đi kèm monkey-patch block surgery U96 phía trước. |
| Local `_u84_processor_run_base` trong `_u85_processor_run` | Được gán nhưng không thấy sử dụng trong thân hàm. |
| `_u88_processor_run_base` và `_u90_processor_run_base` | Cùng trỏ `_u87_processor_run`, nhưng được dùng trong hai cache context khác nhau. |

Đây chỉ là danh sách cần kiểm tra thêm. Không phải danh sách code được phép xóa.

## 9. Cảnh báo

1. Không xóa wrapper chỉ vì luồng PySide6 hiện tại không gọi trực tiếp.
2. Không xóa monkey-patch U95 hoặc U96 cùng wrapper khi chưa chứng minh tác động gián tiếp.
3. Không đổi thứ tự ba lần gán `CORE_RUN_V6`.
4. Không xóa biến `*_base` khi chưa kiểm tra toàn bộ call graph và lịch sử patch.
5. Không trộn refactor wrapper với thay đổi parser, cache schema, repair hoặc ghi Excel.

## 10. Checklist trước khi refactor wrapper

Trước và sau mỗi thay đổi, cần chạy trên bản sao dữ liệu:

- `smoke_import_backend.py` báo `PASS`.
- `check_vas_hc_multi_cus.py` báo `PASS`.
- Chạy app trên dữ liệu copy, không dùng dữ liệu thật.
- So sánh output các sheet `INV`, `PL`, `SUB_DETAIL`, `Folder`.
- Chạy lại khi không đổi dữ liệu để kiểm tra cache.
- Chạy với repair tắt.

Nên bổ sung thêm:

- Test thêm file nguồn.
- Test sửa file nguồn.
- Test xóa file nguồn.
- Test output đang mở trong Excel.
- Test UI PySide6 và entrypoint Tkinter cũ nếu vẫn còn hỗ trợ chạy trực tiếp backend.

Chỉ refactor wrapper khi kết quả trước và sau tương đương, trừ đúng thay đổi nghiệp vụ đã được mô tả rõ.

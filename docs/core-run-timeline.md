# Timeline `CORE_RUN_V6`

## 1. Mục tiêu tài liệu

Tài liệu này ghi lại các lần gán `CORE_RUN_V6` trong `00.Detaisublog_v26.py`.

`CORE_RUN_V6` là biến global trỏ đến hàm xử lý lõi. Nhiều wrapper gọi biến này tại runtime. Vì vậy, thứ tự gán quyết định hàm lõi thực tế được sử dụng.

Tài liệu được lập từ việc đọc code tĩnh. Chưa có kết luận runtime cho đến khi chạy regression test trên bản sao dữ liệu.

## 2. Timeline 3 lần gán

| Dòng gần đúng | Lệnh gán | Hàm hoặc vùng liên quan | Vai trò |
|---|---|---|---|
| `3592` | `CORE_RUN_V6 = Processor.run` | `Processor.run` nền và vùng V6 clean-room detail engine | Lưu method nền làm core ban đầu trước khi bổ sung các wrapper V6. |
| `7896` | `CORE_RUN_V6 = _u90_core_run` | `_u90_core_run`, `_u87_processor_run`, `_u90_processor_run` | Thay core nền bằng core U90 có logic source cache. Wrapper U87 benchmark/smart rebuild gọi core mới này thông qua biến global. |
| `8072` | `CORE_RUN_V6 = _u92_core_run` | `_u92_core_run`, `_u94_processor_run` và các lớp runtime phía sau | Thay core U90 bằng core U92. Đây là lần gán cuối cùng và là giá trị có hiệu lực sau khi module được nạp đầy đủ. |

Giá trị cuối cùng:

```python
CORE_RUN_V6 = _u92_core_run
```

## 3. Cơ chế wrapper gọi global `CORE_RUN_V6`

Các wrapper chủ yếu gọi tên global tại thời điểm runtime:

```python
summary = CORE_RUN_V6(...)
```

Không thấy pattern chụp trực tiếp giá trị core theo kiểu:

```python
SAVED_CORE = CORE_RUN_V6
```

Một số vùng chụp lại wrapper khác, ví dụ:

```python
_u90_processor_run_base = _u87_processor_run
```

Đây không phải bản sao của `CORE_RUN_V6`. Tuy nhiên `_u87_processor_run()` gọi biến global `CORE_RUN_V6`, nên sau lần gán cuối nó sẽ gián tiếp dùng `_u92_core_run`.

Luồng khái quát:

```text
Wrapper runtime
→ gọi CORE_RUN_V6(...)
→ lookup giá trị global hiện tại
→ dùng _u92_core_run sau khi module được nạp đầy đủ
```

## 4. Rủi ro nếu xóa hoặc đổi thứ tự

Không được coi ba lần gán là import hoặc code trùng có thể xóa cơ học.

| Thay đổi | Rủi ro |
|---|---|
| Xóa `CORE_RUN_V6 = Processor.run` | Mất mốc core nền. Có thể gây lỗi hoặc đổi hành vi nếu thứ tự khai báo và wrapper thay đổi sau này. |
| Xóa `CORE_RUN_V6 = _u90_core_run` | Có thể làm luồng U87/U90 quay về core cũ, thay đổi source cache và smart rebuild. |
| Xóa `CORE_RUN_V6 = _u92_core_run` | Giá trị cuối có thể quay về U90 hoặc core nền, làm thay đổi cách ghi output và xử lý incremental. |
| Di chuyển các lần gán | Wrapper gọi global tại runtime nên thay đổi thứ tự có thể đổi hàm lõi thực tế. |
| Refactor chung với cache hoặc parser | Khi output sai, khó xác định nguyên nhân nằm ở core, cache hay parser. |

## 5. Quy tắc an toàn trước khi sửa

1. Không xóa hoặc đổi thứ tự `CORE_RUN_V6` chỉ vì thấy nhiều lần gán cùng một biến.
2. Lập call graph cho toàn bộ wrapper gọi `CORE_RUN_V6(...)`.
3. Kiểm tra các biến trung gian chụp wrapper, ví dụ `_u90_processor_run_base`.
4. Xác nhận `Processor.run()` cuối cùng đang đi qua lớp hoặc wrapper nào.
5. Chỉ sửa một nhóm hành vi trong mỗi commit.
6. Không trộn refactor `CORE_RUN_V6` với thay đổi parser, cache schema, repair hoặc ghi Excel.
7. Chỉ test bằng bản sao dữ liệu nguồn vì backend có chế độ repair ghi đè file nguồn.
8. Giữ khả năng rollback từng commit.

## 6. Checklist test bắt buộc trước khi refactor

Trước và sau refactor, cần đối chiếu cùng một bộ dữ liệu sao chép:

- Import backend thành công, không tự mở Tkinter UI.
- Chạy lần đầu với output mới.
- Chạy lại khi không đổi dữ liệu.
- Thêm một file nguồn rồi chạy lại.
- Sửa một file nguồn rồi chạy lại.
- Xóa một file nguồn rồi chạy lại.
- Chạy với repair tắt.
- Chạy với repair bật trên bản sao dữ liệu.
- Kiểm tra trường hợp output đang mở trong Excel.
- Kiểm tra invoice có nhiều `Cus no`, gồm `INV000000491337` ngày `11/05/2026`.
- Đối chiếu số dòng và dữ liệu trên các sheet `SUB_DETAIL`, `INV`, `PL`, `Folder`.
- Đối chiếu trạng thái `OK`, `WARNING`, `ERROR`.
- Đối chiếu cột `Vas`, `HC`.
- Đối chiếu cache, state và log được tạo hoặc cập nhật.

Chỉ refactor vùng `CORE_RUN_V6` khi kết quả trước và sau tương đương, trừ đúng thay đổi nghiệp vụ đã được mô tả rõ.

# Nhật ký cleanup an toàn cho `00.Detaisublog_v26.py`

## 1. Mục tiêu cleanup

Làm cho backend dễ đọc và dễ bảo trì hơn bằng các thay đổi nhỏ, có kiểm soát. Cleanup không được làm thay đổi nghiệp vụ xuất Excel, cache, repair hoặc luồng chạy hiện tại.

## 2. Nguyên tắc cleanup

- Mỗi lần chỉ sửa hoặc xóa một vùng nhỏ.
- Phân tích tĩnh trước khi sửa.
- Regression runner phải PASS trước và sau thay đổi.
- Mỗi thay đổi cần nằm trong một commit riêng để dễ kiểm tra và rollback.

## 3. Các bước đã làm

1. Thêm section markers cho các vùng lớn trong backend.
2. Xóa các import trùng cấp module đã được xác định là an toàn.
3. Đánh dấu hai dòng self-assignment `_v72_build_detail_bundle = _v72_build_detail_bundle` là candidate.
4. Đánh dấu `_u86_processor_run` là legacy wrapper candidate.
5. Đánh dấu `_u95_processor_run` là legacy wrapper candidate.
6. Đánh dấu `_u96_processor_run` là legacy wrapper candidate.
7. Đánh dấu `_u84_processor_run_base` trong `_u85_processor_run` là local alias candidate.
8. Xóa unused local alias `_u84_processor_run_base`.

## 4. Thay đổi đã xóa thật

### `_u84_processor_run_base` trong `_u85_processor_run`

Dòng đã xóa:

```python
_u84_processor_run_base = _u84_processor_run
```

Lý do xóa:

- Đây là biến local.
- Biến chỉ được gán nhưng không được đọc, gọi hoặc trả về trong `_u85_processor_run`.
- Biến không tham gia benchmark hoặc cache.

Rủi ro: rất thấp.

Test sau khi xóa:

```powershell
tools\run-regression.ps1
```

Kết quả: PASS.

## 5. Các vùng chưa được xóa

Các vùng sau mới chỉ được đánh dấu candidate, chưa được xóa:

- Hai dòng self-assignment `_v72_build_detail_bundle = _v72_build_detail_bundle`.
- Wrapper `_u86_processor_run`.
- Wrapper `_u95_processor_run`.
- Wrapper `_u96_processor_run`.

## 6. Lưu ý

- Chưa xóa wrapper liên quan monkey-patch.
- Chưa thay đổi `CORE_RUN_V6`.
- Chưa thay đổi `Processor.run()` cuối.
- Chưa đụng vào cache, substate hoặc detail index.
- Chưa đụng vào repair file nguồn.
- Chưa đụng vào tích hợp `folder_audit_ext.py`.

Tài liệu này ghi lại tiến trình cleanup từng bước. Không dùng danh sách candidate làm bằng chứng để xóa code nếu chưa phân tích lại và chạy regression runner.

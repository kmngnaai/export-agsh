# Bộ lệnh regression test chuẩn

## 1. Mục tiêu regression test

Bộ lệnh này dùng để kiểm tra nhanh backend và output Excel sau các thay đổi nhỏ trong repo `export-agsh`.

Mục tiêu:

- Xác nhận backend vẫn import được.
- Ghi nhận số dòng trên các sheet chính.
- Kiểm tra regression Vas/HC cho invoice có nhiều `Cus no`.
- Kiểm tra dữ liệu của file nguồn đã xóa không còn trong output.
- Kiểm tra dữ liệu của file nguồn mới đã xuất hiện trong output.
- Đếm số lần xuất hiện của file/path cần theo dõi.

## 2. Khi nào cần chạy

Chạy bộ test này trước và sau khi:

- Dọn import hoặc comment.
- Sửa logic backend.
- Sửa `folder_audit_ext.py`.
- Thay đổi cache, state hoặc detail index.
- Refactor wrapper runtime.
- Thay đổi logic đọc hoặc ghi Excel.

Chỉ dùng dữ liệu test đã sao chép. Không chạy test trên dữ liệu thật nếu chưa có bản copy.

## 3. Đường dẫn output test hiện dùng

```text
D:\01.AutobyNgan\00.Build.App\11.CODEX\02_LOCAL_TEST\export-agsh-test-data\202605test\test_output_fix_vas_hc.xlsx
```

Trong các lệnh dưới đây, `<output>` là đường dẫn trên.

## 4. Bộ lệnh test chuẩn

Chạy từ thư mục gốc repo `export-agsh`:

```powershell
python tests/smoke_import_backend.py

python tests/check_output_baseline.py "D:\01.AutobyNgan\00.Build.App\11.CODEX\02_LOCAL_TEST\export-agsh-test-data\202605test\test_output_fix_vas_hc.xlsx"

python tests/check_vas_hc_multi_cus.py "D:\01.AutobyNgan\00.Build.App\11.CODEX\02_LOCAL_TEST\export-agsh-test-data\202605test\test_output_fix_vas_hc.xlsx"

python tests/check_deleted_path_absent.py "D:\01.AutobyNgan\00.Build.App\11.CODEX\02_LOCAL_TEST\export-agsh-test-data\202605test\test_output_fix_vas_hc.xlsx" "SUB-INV33333.xls"

python tests/check_path_present.py "D:\01.AutobyNgan\00.Build.App\11.CODEX\02_LOCAL_TEST\export-agsh-test-data\202605test\test_output_fix_vas_hc.xlsx" "SUB-INV3333test.xls"

python tests/check_path_occurrences.py "D:\01.AutobyNgan\00.Build.App\11.CODEX\02_LOCAL_TEST\export-agsh-test-data\202605test\test_output_fix_vas_hc.xlsx" "SUB-INV3333test.xls"
```

## 5. Kết quả đúng mong đợi

| Kiểm tra | Kết quả mong đợi |
|---|---|
| Import backend | `PASS` |
| `SUB_DETAIL` | `rows=170` |
| `INV` | `rows=1382` |
| `PL` | `rows=1106` |
| `Folder` | `rows=898` |
| Vas/HC invoice nhiều `Cus no` | Invoice total `PASS` |
| `SUB-INV33333.xls` | Absent `PASS` |
| `SUB-INV3333test.xls` | Present `PASS` |
| Số lần xuất hiện `SUB-INV3333test.xls` | `TOTAL: occurrences=37` |

## 6. Lưu ý về baseline

Số dòng và số lần xuất hiện có thể thay đổi nếu dữ liệu test thay đổi.

Khi chủ động thay đổi dữ liệu test:

1. Chạy lại bộ lệnh trên với output mới.
2. Kiểm tra kết quả thủ công.
3. Cập nhật baseline trong tài liệu này.
4. Ghi rõ lý do baseline thay đổi trong commit.

Không sửa baseline chỉ để làm test pass khi chưa hiểu nguyên nhân sai lệch.

## 7. Lưu ý an toàn dữ liệu

- Không chạy test trên dữ liệu thật nếu chưa có bản copy.
- Khi chạy app để tạo output test, dùng thư mục input-copy và output test riêng.
- Tắt repair nếu mục tiêu chỉ là kiểm tra output.
- Nếu cần test repair, chỉ bật repair trên bản sao dữ liệu.

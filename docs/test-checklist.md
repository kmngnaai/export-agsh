# Checklist test chuẩn cho repo export-agsh

## 1. Mục tiêu test

Checklist này dùng để tạo mốc so sánh trước khi sửa `00.Detaisublog_v26.py`.

Mục tiêu:

- Xác nhận chương trình đọc đúng file nguồn.
- Ghi nhận output hiện tại để so sánh sau khi sửa code.
- Kiểm tra cache và state không làm sai kết quả khi chạy lại.
- Kiểm tra hành vi khi thêm, sửa hoặc xóa file nguồn.
- Tách riêng test repair vì repair có thể sửa trực tiếp file nguồn.
- Kiểm tra trường hợp một invoice có nhiều Cus no.

Không dùng dữ liệu thật để test trực tiếp.

## 2. Chuẩn bị dữ liệu bản sao

Tạo một thư mục test riêng, ví dụ:

```text
D:\export-agsh-test\
├── input-copy\
├── output\
└── evidence\
```

Checklist chuẩn bị:

- [ ] Sao chép một bộ dữ liệu nguồn đại diện vào `input-copy\`.
- [ ] Không trỏ chương trình vào thư mục dữ liệu thật.
- [ ] Giữ nguyên tên folder và tên file trong bản sao để kiểm tra parser.
- [ ] Có ít nhất một file `.xlsx` hoặc `.xlsm`.
- [ ] Nếu hệ thống còn dùng `.xls`, thêm ít nhất một file `.xls`.
- [ ] Có ít nhất một file hợp lệ.
- [ ] Có ít nhất một file tạo `WARNING` hoặc `ERROR`.
- [ ] Có bộ dữ liệu chứa invoice `INV000000491337`, ngày `11/05/2026`, nếu cần kiểm tra nhiều Cus no.
- [ ] Chụp lại danh sách file nguồn ban đầu vào `evidence\baseline-files.txt` hoặc bảng ghi chú tương đương.

Trước mỗi nhóm test độc lập:

- [ ] Tạo lại `input-copy\` từ bản sao sạch.
- [ ] Dùng output mới.
- [ ] Không tái sử dụng cache cũ, trừ test chạy lại không đổi dữ liệu.

## 3. Cách tạo output mới

Tạo output trong thư mục test riêng:

```text
D:\export-agsh-test\output\SUB_DETAIL_OUTPUT_TEST.xlsx
```

Checklist:

- [ ] Đảm bảo file output chưa tồn tại trước test chạy lần đầu.
- [ ] Đảm bảo không có file sidecar cũ cùng tên output.
- [ ] Chọn đúng thư mục `input-copy\`.
- [ ] Chọn output trong `output\`, không đặt output trong dữ liệu thật.
- [ ] Ghi lại tên output và thời điểm test.

Các sidecar có thể xuất hiện cạnh output:

```text
*.bak
*.runtime_bench.log
*.sourcecache*.pkl
*.fingerprintcache*.json
*.detailindex*.pkl
*.substate*.json
*.issue_snapshot*.json
*.folderauditcache*.json
```

## 4. Test chạy lần đầu

Mục tiêu: tạo baseline đầy đủ với output mới và cache sạch.

Các bước:

- [ ] Dùng `input-copy\` sạch.
- [ ] Chọn output mới chưa tồn tại.
- [ ] Ghi rõ trạng thái repair đang bật hay tắt.
- [ ] Chạy chương trình một lần.
- [ ] Mở output sau khi chương trình hoàn tất.
- [ ] Ghi số dòng của các sheet `SUB_DETAIL`, `INV`, `PL`, `Folder`.
- [ ] Ghi số lượng `OK`, `WARNING`, `ERROR`.
- [ ] Ghi danh sách file lỗi nổi bật.
- [ ] Ghi danh sách sidecar cache/log được tạo.
- [ ] Lưu ảnh chụp hoặc bảng ghi chú baseline trong `evidence\`.

Kết quả mong đợi:

- [ ] Output `.xlsx` được tạo.
- [ ] Các sheet cần thiết tồn tại.
- [ ] Không mất file nguồn.
- [ ] Nếu repair tắt, file nguồn không bị sửa.

## 5. Test chạy lại khi không đổi dữ liệu

Mục tiêu: kiểm tra cache và state giúp bỏ qua công việc không cần thiết nhưng không làm thay đổi kết quả.

Các bước:

- [ ] Giữ nguyên `input-copy\`.
- [ ] Giữ nguyên output và sidecar từ test lần đầu.
- [ ] Chạy lại chương trình.
- [ ] So sánh số dòng từng sheet trước và sau.
- [ ] So sánh số lượng `OK`, `WARNING`, `ERROR`.
- [ ] Kiểm tra log/cache có ghi nhận skip hoặc zero-work phù hợp.

Kết quả mong đợi:

- [ ] Không phát sinh dòng trùng trong `SUB_DETAIL`, `INV`, `PL`.
- [ ] Số dòng nghiệp vụ không đổi.
- [ ] Kết quả lỗi không biến mất sai.
- [ ] Runtime log hoặc benchmark cho thấy file không đổi được bỏ qua khi phù hợp.

## 6. Test thêm file nguồn

Mục tiêu: kiểm tra chương trình chỉ bổ sung dữ liệu cần thiết.

Các bước:

- [ ] Bắt đầu từ baseline đã chạy ít nhất một lần.
- [ ] Sao chép thêm một file nguồn hợp lệ vào `input-copy\`.
- [ ] Ghi lại tên file mới.
- [ ] Chạy lại chương trình.
- [ ] Kiểm tra `SUB_DETAIL`, `INV`, `PL`, `Folder`.
- [ ] Kiểm tra cache/state có thêm file mới.

Kết quả mong đợi:

- [ ] File mới được nhận diện.
- [ ] Dữ liệu cũ không bị nhân đôi.
- [ ] Sheet `Folder` phản ánh file mới.
- [ ] Các chỉ số `OK`, `WARNING`, `ERROR` thay đổi hợp lý.

## 7. Test sửa file nguồn

Mục tiêu: kiểm tra fingerprint và state phát hiện file đã thay đổi.

Các bước:

- [ ] Dùng một file trong `input-copy\`, không sửa file thật.
- [ ] Ghi lại giá trị cũ cần đối chiếu.
- [ ] Sửa một giá trị nghiệp vụ có thể quan sát trong output.
- [ ] Lưu file nguồn bản sao.
- [ ] Chạy lại chương trình.
- [ ] Kiểm tra dòng tương ứng trong `SUB_DETAIL`, `INV`, `PL`.
- [ ] Kiểm tra cache/state cập nhật.

Kết quả mong đợi:

- [ ] Chỉ file đã sửa được xử lý lại khi phù hợp.
- [ ] Dòng cũ được cập nhật hoặc thay thế đúng.
- [ ] Không tạo dòng trùng.
- [ ] Không ảnh hưởng dữ liệu của file nguồn khác.

## 8. Test xóa file nguồn

Mục tiêu: kiểm tra dọn dữ liệu stale khi một file không còn trong thư mục nguồn.

Các bước:

- [ ] Bắt đầu từ baseline có file cần xóa.
- [ ] Ghi lại các dòng output thuộc file đó.
- [ ] Xóa file khỏi `input-copy\`, không xóa file thật.
- [ ] Chạy lại chương trình.
- [ ] Kiểm tra `SUB_DETAIL`, `INV`, `PL`, `Folder`.
- [ ] Kiểm tra detail index, substate và fingerprint cache.

Kết quả mong đợi:

- [ ] Dòng thuộc file đã xóa được loại khỏi output nếu logic yêu cầu.
- [ ] Sheet `Folder` không còn file đó.
- [ ] Cache/state không giữ bản ghi stale gây sai lần chạy sau.
- [ ] Dữ liệu file khác vẫn nguyên vẹn.

## 9. Test repair tắt

Mục tiêu: xác nhận chế độ chỉ đọc file nguồn.

Chuẩn bị:

- [ ] Tạo lại `input-copy\` sạch.
- [ ] Ghi checksum, thời gian sửa hoặc bản sao so sánh của file nguồn.
- [ ] Tắt mọi tùy chọn repair trên UI.

Các bước:

- [ ] Chạy với output mới.
- [ ] So sánh file nguồn trước và sau.
- [ ] Kiểm tra output và log lỗi.

Kết quả mong đợi:

- [ ] File nguồn không thay đổi nội dung.
- [ ] File sai metadata chỉ được báo `WARNING` hoặc `ERROR`.
- [ ] Output vẫn được tạo bình thường.

## 10. Test repair bật

Mục tiêu: kiểm tra chức năng sửa metadata nguồn một cách có kiểm soát.

Chỉ chạy trên bản sao dữ liệu.

Chuẩn bị:

- [ ] Tạo `input-copy\` riêng cho repair.
- [ ] Giữ bản sao trước repair trong `evidence\before-repair\`.
- [ ] Chọn một file `.xlsx` hoặc `.xlsm` có metadata sai.
- [ ] Ghi rõ repair theo tên folder hay theo giá trị nhập tay.

Các bước:

- [ ] Bật đúng tùy chọn repair.
- [ ] Chạy chương trình.
- [ ] So sánh file nguồn trước và sau.
- [ ] Kiểm tra `invoice_no`, `invoice_date`, `destination`.
- [ ] Chạy lại lần hai để kiểm tra trạng thái ổn định.

Kết quả mong đợi:

- [ ] Chỉ file dự kiến bị sửa.
- [ ] Chỉ metadata dự kiến thay đổi.
- [ ] File `.xls` không bị tự sửa.
- [ ] Log thể hiện file đã repair hoặc lý do bỏ qua.

## 11. Test output đang mở trong Excel

Mục tiêu: kiểm tra xử lý an toàn khi output bị khóa.

Các bước:

- [ ] Dùng output test, không dùng output thật.
- [ ] Mở output test bằng Excel và giữ file đang mở.
- [ ] Thay đổi một file trong `input-copy\` để buộc chương trình cần ghi output.
- [ ] Chạy lại chương trình.
- [ ] Ghi lại thông báo lỗi.
- [ ] Kiểm tra file `.bak` và file thay thế nếu được tạo.

Kết quả mong đợi:

- [ ] Chương trình không làm mất output cũ.
- [ ] Có thông báo yêu cầu đóng file Excel.
- [ ] Nếu lưu file thay thế, đường dẫn được báo rõ.
- [ ] Không để sót file tạm không cần thiết.

## 12. Test invoice có nhiều Cus no

Mục tiêu: kiểm tra tổng `Vas` và `HC` theo từng Cus no, không cộng dồn sai toàn invoice.

Dữ liệu kiểm tra:

```text
Invoice: INV000000491337
Ngày:    11/05/2026

Cus no:
- 308518255900
- 308518722800
- 308519134360
```

Các bước:

- [ ] Dùng bản sao dữ liệu có invoice trên.
- [ ] Chạy với output mới và cache sạch.
- [ ] Mở sheet `INV`.
- [ ] Lọc invoice `INV000000491337`.
- [ ] Kiểm tra các cột `Cus no.-`, `Cus no.-STT`, `Vas`, `HC`.
- [ ] Kiểm tra từng dòng `TOTAL` của mỗi Cus no.
- [ ] Cộng lại ba block để đối chiếu tổng invoice.
- [ ] Chạy lại lần hai để kiểm tra cache không làm đổi kết quả.

Kết quả mong đợi:

| Cus no | Vas | HC |
|---|---:|---:|
| `308518255900` | `285.01` | `160.00` |
| `308518722800` | `219.30` | `160.00` |
| `308519134360` | `502.64` | `480.00` |
| **Tổng invoice** | **`1,006.95`** | **`800.00`** |

Không chấp nhận:

- [ ] Mỗi dòng `TOTAL` bị gán `Vas = 1,006.95`.
- [ ] Mỗi dòng `TOTAL` bị gán `HC = 800.00`.
- [ ] `Cus no.-` hoặc `Cus no.-STT` bị trộn giữa các block.

## 13. Các kết quả cần đối chiếu

Tạo một bảng baseline cho mỗi lần chạy:

| Hạng mục | Trước chạy | Sau chạy | Ghi chú |
|---|---:|---:|---|
| Số file nguồn |  |  |  |
| Dòng `SUB_DETAIL` |  |  | Không tính header |
| Dòng `INV` |  |  | Không tính header |
| Dòng `PL` |  |  | Không tính header |
| Dòng `Folder` |  |  | Không tính header |
| `OK` |  |  |  |
| `WARNING` |  |  |  |
| `ERROR` |  |  |  |
| File được repair |  |  |  |
| Cache/log mới |  |  |  |

Checklist đối chiếu:

- [ ] Header các sheet không bị mất hoặc đổi ngoài dự kiến.
- [ ] Không có dòng trùng sau khi chạy lại.
- [ ] `Vas`, `HC` đúng theo từng Cus no.
- [ ] `Cus no`, `Cus no.1`, `Cus no.-`, `Cus no.-STT` đúng block.
- [ ] Cache/log tồn tại đúng nhu cầu và không bị commit lên Git.
- [ ] File `.bak` chỉ xuất hiện khi có output cũ cần bảo vệ.
- [ ] Không có file nguồn thật bị sửa.

## 14. Khi nào được commit

Chỉ commit thay đổi backend khi:

- [ ] Đã lưu baseline trước khi sửa code.
- [ ] Test chạy lần đầu đạt.
- [ ] Test chạy lại không đổi dữ liệu không tạo dòng trùng.
- [ ] Test thêm, sửa, xóa file nguồn đạt.
- [ ] Test repair tắt xác nhận file nguồn không đổi.
- [ ] Nếu sửa logic repair, test repair bật đạt trên bản sao.
- [ ] Nếu sửa ghi output, test output đang mở trong Excel đạt.
- [ ] Nếu liên quan audit, test invoice nhiều Cus no đạt.
- [ ] Đã kiểm tra `SUB_DETAIL`, `INV`, `PL`, `Folder`.
- [ ] Đã đối chiếu `OK`, `WARNING`, `ERROR`.
- [ ] Đã xem cache/log/state phát sinh.
- [ ] Mỗi commit chỉ xử lý một mục tiêu rõ ràng.
- [ ] Không trộn refactor lớn vào commit sửa lỗi.

Nếu chưa đạt đủ test liên quan, chưa commit thay đổi backend.

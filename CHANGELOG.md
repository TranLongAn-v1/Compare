# Changelog

## [2.7.0] - 2026-05-11
- Thêm option Đối chiếu > 7980 vs Mau 01.
- Thêm module mới modules/module_7980_mau01/* gồm UI và logic xuất Excel tô màu.
- Luồng 7980 vs Mau 01 lấy Mẫu 01 làm chuẩn, đối chiếu các trường: Mã thẻ, Họ và tên, Ngày vào, Ngày ra, Tổng chi BV với MA_THE, HO_TEN, NGAY_VAO, NGAY_RA, T_TONGCHI.
- Bổ sung chuẩn hóa tên tiếng Việt bỏ dấu trong 7980 vs Mau 01 để giảm sai lệch khi ghép cặp.
- Cập nhật parse ngày giờ cho Mẫu 01: bỏ qua hậu tố SA/CH trước khi chuyển đổi datetime.
- Bổ sung phân loại và tô màu lỗi cho 7980 vs Mau 01:
  - Có trên 7980 nhưng không có trên Mẫu 01 (vàng)
  - Lệch ngày vào/ngày ra (cam)
  - Lệch tổng chi (đỏ)
  - Lệch mã thẻ (tím)
  - Có trên Mẫu 01 nhưng không có trên 7980 (xanh)
## [2.6.0] - 2026-05-05
- Thêm menu `Xử lý > Xml1+Xml3` với 2 input `XML1` + `XML3`, kiểm tra tên file đầu vào và báo `File không hợp lệ` khi sai rule.
- Thêm luồng kết xuất đối chiếu `MA_LK + MA_HSBA` (XML1) với `MA_LK + MA_BA` (XML3), kiểm tra thêm khớp thẻ `MA_THE_BHYT/MA_THE`.
- Bổ sung bảng hiển thị + lọc đa cột cho `Xml1+Xml3` dùng chung popup `core/filter_popup.py`.
- Định dạng hiển thị `NGAY_VAO`, `NGAY_RA` ở `Xml1+Xml3` sang `dd/mm/yyyy HH:MM`.
- Cập nhật logic `Chuyên đề 72`:
  - Lọc khoảng ngày theo `NGAY_RA` của XML1 (thay vì mốc ngày từ XML3).
  - Chỉ xét XML3 theo tập `MA_LK` đã qua lọc `NGAY_RA` từ XML1.
  - Loại `MA_NHOM` thuộc `1, 3, 8, 10, 13, 14, 15` trước bước phát hiện lỗi.

## [2.5.1] - 2026-05-04
- Thêm option `Ánh xạ` trong `Đối chiếu > Mã máy` và luồng đối chiếu file ánh xạ với `DM_MAY`.
- Cập nhật rule loại bỏ toàn bộ mã `MA_MAY` cũ có chứa `.44194` khỏi kết quả option `Ánh xạ`.
- Cập nhật `Xử lý > XML3`: lọc thêm `MA_NHOM = 18`.
- Đồng bộ popup lọc giữa `Xử lý > XML3` và `Xử lý > Mã máy` bằng helper dùng chung (`core/filter_popup.py`), hỗ trợ tìm kiếm + lọc theo ngữ cảnh.

## [2.4.0] - 2026-04-30
- Thêm menu `Xử lý > XML3` với luồng tổng hợp và xuất file.
- Thêm menu `Lọc > Mã dịch vụ` với lọc trùng theo `MA_DICH_VU + TEN_DICH_VU`.
- Thêm menu `Đối chiếu > Mã máy` để so sánh `MA_MAY` với `MAMAY3176` và `MAMAY` từ DM.
- Chuẩn hóa nhiều màn hình sang tiếng Việt có dấu.
- Bổ sung quy tắc kiểm tra đầu vào chặt chẽ cho các file nghiệp vụ.


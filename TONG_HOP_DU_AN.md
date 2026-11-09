# TỔNG HỢP DỰ ÁN ĐỐI CHIẾU DỮ LIỆU BHYT

## 1) Thông tin phiên bản
- Phiên bản hiện tại: `2.7.0`
- Ngày cập nhật tài liệu: `2026-05-11`
- Nguồn phiên bản:
1. File `VERSION`.
2. File `CHANGELOG.md`.

---

## 2) Mục tiêu dự án
- Xây dựng ứng dụng desktop Python để đối chiếu và xử lý dữ liệu BHYT từ nhiều nguồn Excel/XML.
- Chuẩn hóa quy trình nhập file, kiểm tra dữ liệu đầu vào, hiển thị kết quả và xuất Excel.
- Tối ưu cho nghiệp vụ nội bộ: thao tác nhanh, dễ dò lỗi, dễ tổng hợp theo phiên bản.

---

## 3) Công nghệ và kiến trúc

### 3.1 Công nghệ
- Ngôn ngữ: `Python`.
- UI: `customtkinter` + `tkinter` + `ttk`.
- Xử lý dữ liệu: `pandas`, `numpy`.
- Xuất Excel: `openpyxl`.

### 3.2 Kiến trúc module
- Mỗi nghiệp vụ tách thành 2 lớp:
1. `ui_xxx.py`: nhận input, gọi xử lý, hiển thị bảng.
2. `logic_xxx.py`: kiểm tra dữ liệu, xử lý nghiệp vụ, xuất Excel.

### 3.3 Cấu trúc thư mục chính
- `main.py`: khởi động app, toolbar, điều hướng màn hình.
- `core/utils.py`: hàm dùng chung đọc file/chuẩn hóa.
- `core/filter_popup.py`: popup lọc dùng chung (tìm kiếm + lọc theo ngữ cảnh đa cột).
- `modules/module_7980/*`: đối chiếu 79/80 với Cổng BH.
- `modules/module_7980_mau01/*`: đối chiếu 79/80 với Mẫu 01 BH.
- `modules/module_69/*`: chuyên đề 69.
- `modules/module_72/*`: chuyên đề 72.
- `modules/module_xml3/*`: xử lý XML3 (tổng hợp cột theo điều kiện MA_NHOM).
- `modules/module_xml1_xml3/*`: xử lý kết hợp XML1 + XML3.
- `modules/module_loc_ma_dich_vu/*`: lọc trùng mã dịch vụ.
- `modules/module_ma_may/*`: đối chiếu mã máy với DM_MAY.
- `modules/module_xu_ly_ma_may/*`: dự đoán MA_MAY còn trống từ file đối chiếu mã máy.
- `build_exe.ps1`, `*.spec`, `hooks/*`, `rthook_tk.py`: đóng gói EXE.

---

## 4) Giao diện hiện tại

### 4.1 Cửa sổ chính (`main.py`)
- Kích thước mặc định: `1300x780`, min `1100x680`.
- Có hiển thị phiên bản trên title và trạng thái toolbar.
- Thanh công cụ trên cùng (hover menu):
1. `Đối chiếu`: `7980`, `7980 vs Mau 01`, `Chuyên đề 69`, `Chuyên đề 72`, `Mã máy`.
2. `Xử lý`: `XML3`, `Xml1+Xml3`, `Mã máy`.
3. `Lọc`: `Mã dịch vụ`.

### 4.2 Mẫu bố cục màn hình nghiệp vụ
- Mỗi màn hình chia 3 phần:
1. `Phần 1 - Input file`.
2. `Phần 2 - Hướng dẫn và thao tác`.
3. `Phần 3 - Hiển thị dữ liệu` (Treeview + scrollbar ngang/dọc).

---

## 5) Tổng hợp chức năng theo menu

### 5.1 Nhóm `Đối chiếu`
1. `7980`: đối chiếu 79/80 với Cổng BH, phân loại sai lệch, tô màu, lọc đa cột, xuất Excel.
2. `7980 vs Mau 01`:
- Đối chiếu file `7980a_44194...` với `Mau01BH_44194...`.
- Lấy `Mẫu 01` làm chuẩn theo các trường: `Mã thẻ`, `Họ và tên`, `Ngày vào`, `Ngày ra`, `Tổng chi BV`.
- So khớp với 7980 theo các trường tương ứng: `MA_THE`, `HO_TEN`, `NGAY_VAO`, `NGAY_RA`, `T_TONGCHI`.
- Chuẩn hóa `NGAY_VAO`, `NGAY_RA` phía 7980 theo cùng cách module `7980` (chuỗi 12 số `yyyyMMddHHmm` -> `dd/mm/yyyy HH:MM`).
- Khi đọc ngày giờ từ `Mẫu 01`, bỏ qua hậu tố `SA/CH` để parse chính xác.
- Nhóm lỗi và màu:
1. Có trên 7980 nhưng không có trên Mẫu 01 (vàng).
2. Lệch ngày vào/ngày ra (cam).
3. Lệch tổng chi (đỏ).
4. Lệch mã thẻ (tím).
5. Có trên Mẫu 01 nhưng không có trên 7980 (xanh).
3. `Chuyên đề 69`: phát hiện NGAY_KQ muộn hơn NGAY_RA (XML3/XML4 so với XML1).
4. `Chuyên đề 72`:
- Phát hiện trùng bộ thời gian dịch vụ trong cùng nhóm.
- Lọc khoảng ngày theo `NGAY_RA` của XML1 (định dạng 12 số `yyyyMMddHHmm`).
- Khi tổng hợp XML3 bỏ các dòng có `MA_NHOM` thuộc `1, 3, 8, 10, 13, 14, 15`.
5. `Mã máy`:
- Input 1: file bắt đầu `Tong_hop_XML3`, kiểm tra đúng bộ cột `MA_DICH_VU, TEN_DICH_VU, MA_MAY, MA_BA, HO_TEN` (không thừa/không thiếu).
- Input 2: file bắt đầu `DM_MAY`, dùng cột `MAMAY3176`, `MAMAY` làm chuẩn.
- Có option:
1. `Mã máy`: đối chiếu `MA_MAY` với `MAMAY3176` và `MAMAY`.
2. `Ánh xạ`: đối chiếu `MA_MAY` từ file ánh xạ với `MAMAY3176` trong `DM_MAY`.
- Ở option `Ánh xạ`, các dòng có `MA_MAY` chứa `.44194` sẽ bị loại khỏi kết quả.
- Kết quả có 3 cột cờ `Có/Không`:
1. `Khớp MAMAY3176`.
2. `Khớp MAMAY`.
3. `Không tồn tại trong DM`.

### 5.2 Nhóm `Xử lý`
1. `XML3`:
- Nhận file bắt đầu `XML3`.
- Khi tổng hợp, lấy các cột `MA_DICH_VU, TEN_DICH_VU, MA_MAY, MA_BA, HO_TEN`.
- Loại các dòng có `MA_NHOM` thuộc `3, 8, 10, 13, 14, 15, 18`.
- Có lọc theo cột trực tiếp trên bảng và xuất file.
2. `Xml1+Xml3`:
- Input 1: file bắt đầu `XML1`.
- Input 2: file bắt đầu `XML3`.
- Nếu sai rule tên file thì báo `File không hợp lệ`.
- Nút thao tác: `Kết xuất`, `Xuất file`.
- Kết xuất lấy XML1 làm chuẩn theo `MA_LK + MA_HSBA`, đối chiếu với XML3 theo `MA_LK + MA_BA`.
- Kiểm tra khớp thêm cặp thẻ `MA_THE_BHYT` (XML1) và `MA_THE` (XML3).
- Lấy cột hiển thị: `MA_HSBA, MA_BN, HO_TEN, NGAY_VAO, NGAY_RA, MA_DICH_VU, TEN_DICH_VU, MA_MAY`.
- Định dạng hiển thị `NGAY_VAO`, `NGAY_RA` thành `dd/mm/yyyy HH:MM`.
- Với phần lấy dịch vụ từ XML3, bỏ các dòng có `MA_NHOM` thuộc `3, 8, 10, 13, 14, 15, 18`.
- Hỗ trợ lọc đa cột bằng popup dùng chung `core/filter_popup.py`.
3. `Mã máy`:
- Nhận file bắt đầu `Doi_Chieu_Ma_May` (không phụ thuộc phần tên phía sau).
- Có 3 nút thao tác: `Dự đoán`, `Xuất file`, `Xóa lọc file`.
- Khi bấm `Dự đoán`: dựa trên cặp `MA_DICH_VU + TEN_DICH_VU` để điền `MA_MAY` còn trống.
- Chỉ dùng dữ liệu tham chiếu từ các dòng có `Khớp MAMAY3176 = Có` và `MA_MAY` không trống.
- Thêm cột `THAM_CHIEU_MA_MAY` để hiển thị nguồn mã máy tham chiếu.
- Hỗ trợ lọc đa cột theo ngữ cảnh hiện tại và có ô tìm kiếm trong popup lọc.
- Popup lọc của `XML3`, `Xml1+Xml3` và `Mã máy` dùng chung một logic để đảm bảo hành vi nhất quán.

### 5.3 Nhóm `Lọc`
1. `Mã dịch vụ`:
- Nhận file bắt đầu `Tong_hop_XML3`.
- Chỉ nhận khi bộ cột khớp chính xác `MA_DICH_VU, TEN_DICH_VU, MA_MAY, MA_BA, HO_TEN`.
- Nút lọc trùng dựa trên cặp `MA_DICH_VU + TEN_DICH_VU`.
- Phần 3 chỉ hiển thị `MA_DICH_VU, TEN_DICH_VU, MA_MAY`.

---

## 6) Quy tắc dữ liệu đầu vào quan trọng
1. Rule tên file và rule cột được kiểm tra ngay khi chọn file.
2. Nếu sai định dạng theo rule module, hiển thị: `File không hợp lệ`.
3. Với các màn mới, chuẩn kiểm tra cột đang dùng so khớp tập cột (set), nên không phụ thuộc thứ tự cột.
4. Với rule “không thừa, không thiếu”, file phải có đúng tập cột yêu cầu mới được nhận.
5. Riêng `Đối chiếu > 7980 vs Mau 01`, cột ngày giờ `Ngày vào`, `Ngày ra` ở `Mẫu 01` được parse theo `dd/mm/yyyy HH:MM` và bỏ qua hậu tố `SA/CH` nếu có.

---

## 7) Hiệu năng và hành vi hiển thị
1. Các bảng preview UI hiển thị tối đa `300` dòng để giảm lag.
2. Logic xử lý vẫn chạy trên toàn bộ tập dữ liệu trong bộ nhớ.
3. Xuất file có cơ chế tránh ghi đè: tự thêm hậu tố `(1)`, `(2)` nếu trùng tên.

---

## 8) Danh sách file quan trọng khi onboarding
1. `main.py`
2. `VERSION`
3. `CHANGELOG.md`
4. `core/utils.py`
5. `core/filter_popup.py`
6. `modules/module_7980/ui_7980.py`
7. `modules/module_7980/logic_7980.py`
8. `modules/module_xml3/ui_xml3.py`
9. `modules/module_xml3/logic_xml3.py`
10. `modules/module_xml1_xml3/ui_xml1_xml3.py`
11. `modules/module_xml1_xml3/logic_xml1_xml3.py`
12. `modules/module_72/ui_72.py`
13. `modules/module_72/logic_72.py`
14. `modules/module_loc_ma_dich_vu/ui_loc_ma_dich_vu.py`
15. `modules/module_loc_ma_dich_vu/logic_loc_ma_dich_vu.py`
16. `modules/module_ma_may/ui_ma_may.py`
17. `modules/module_ma_may/logic_ma_may.py`
18. `modules/module_xu_ly_ma_may/ui_xu_ly_ma_may.py`
19. `modules/module_xu_ly_ma_may/logic_xu_ly_ma_may.py`
20. `modules/module_7980_mau01/ui_7980_mau01.py`
21. `modules/module_7980_mau01/logic_7980_mau01.py`

---

## 9) Cách cập nhật tài liệu cho phiên bản sau
1. Tăng `VERSION` theo SemVer:
- `MAJOR`: thay đổi phá vỡ luồng cũ.
- `MINOR`: thêm chức năng mới.
- `PATCH`: sửa lỗi/chỉnh logic nhỏ.
2. Thêm mục mới vào `CHANGELOG.md`.
3. Cập nhật lại các mục trong file này:
- Mục `1) Thông tin phiên bản`.
- Mục `5) Tổng hợp chức năng theo menu`.
- Mục `6) Quy tắc dữ liệu đầu vào quan trọng` nếu có đổi rule.

---

## 10) Trạng thái hiện tại
- App đã có cơ chế version nội bộ (`VERSION`, title, toolbar).
- Hệ chức năng hiện gồm 3 nhóm menu: `Đối chiếu`, `Xử lý`, `Lọc`.
- Các màn mới (`XML3`, `Xml1+Xml3`, `Mã dịch vụ`, `Mã máy`, `Xử lý > Mã máy`) đã có kiểm soát input rõ ràng và xuất Excel.


import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from core.date_entry import DateEntry
import os
import subprocess # --- [GHI CHÚ SỬA ĐỔI]: THÊM THƯ VIỆN ĐỂ GỌI LỆNH HỆ THỐNG ---
from .logic_4210 import xu_ly_doi_chieu_4210

class UI_4210(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.last_output_path = None
        self.tao_giao_dien()

    def tao_giao_dien(self):
        lbl_title = ctk.CTkLabel(self, text="ĐỐI CHIẾU DANH SÁCH VIỆN PHÍ VÀ 79A/80A", font=ctk.CTkFont(size=22, weight="bold"))
        lbl_title.pack(pady=(20, 20))

        frame_file = ctk.CTkFrame(self, fg_color="transparent")
        frame_file.pack(fill="x", padx=40, pady=10)

        ctk.CTkLabel(frame_file, text="File Viện Phí:", width=100, anchor="w").grid(row=0, column=0, pady=10)
        self.entry_vp = ctk.CTkEntry(frame_file, width=300, placeholder_text="Có thể chọn nhiều file cùng lúc")
        self.entry_vp.grid(row=0, column=1, padx=10)
        ctk.CTkButton(frame_file, text="Chọn File", width=80, command=lambda: self.chon_nhieu_file(self.entry_vp)).grid(row=0, column=2)

        ctk.CTkLabel(frame_file, text="File 79/80:", width=100, anchor="w").grid(row=1, column=0, pady=10)
        self.entry_7980 = ctk.CTkEntry(frame_file, width=300, placeholder_text="Bắt buộc: 7980a_44194...")
        self.entry_7980.grid(row=1, column=1, padx=10)
        ctk.CTkButton(frame_file, text="Chọn File", width=80, command=lambda: self.chon_file(self.entry_7980, "7980")).grid(row=1, column=2)

        frame_date = ctk.CTkFrame(self, fg_color="transparent")
        frame_date.pack(fill="x", padx=40, pady=20)
        
        ctk.CTkLabel(frame_date, text="Từ ngày:").grid(row=0, column=0, padx=5)
        self.cal_tu_ngay = DateEntry(frame_date, width=12, date_pattern='dd/mm/yyyy', locale='vi_VN')
        self.cal_tu_ngay.grid(row=0, column=1, padx=10)
        
        ctk.CTkLabel(frame_date, text="Đến ngày:").grid(row=0, column=2, padx=5)
        self.cal_den_ngay = DateEntry(frame_date, width=12, date_pattern='dd/mm/yyyy', locale='vi_VN')
        self.cal_den_ngay.grid(row=0, column=3, padx=10)

        self.btn_run = ctk.CTkButton(self, text="▶ BẮT ĐẦU TỔNG HỢP", fg_color="green", command=self.chay_xu_ly)
        self.btn_run.pack(pady=(20, 10))

        self.btn_mo_thu_muc = ctk.CTkButton(self, text="📁 Mở thư mục chứa file", command=self.mo_thu_muc_chua, state="disabled")
        self.btn_mo_thu_muc.pack(pady=(0, 20))

        self.lbl_status = ctk.CTkLabel(self, text="Sẵn sàng.", font=ctk.CTkFont(slant="italic"))
        self.lbl_status.pack()

    def chon_nhieu_file(self, entry_widget):
        filepaths = filedialog.askopenfilenames(filetypes=[("Excel Viện phí", "*.xls")])
        if filepaths:
            chuoi_duong_dan = ";".join(filepaths)
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, chuoi_duong_dan)

    def chon_file(self, entry_widget, loai_file):
        filepath = filedialog.askopenfilename(filetypes=[("Tất cả các file", "*.*"), ("Excel", "*.xls *.xlsx *.csv")])
        if filepath:
            ten_file = os.path.basename(filepath)
            if loai_file == "7980" and not ten_file.startswith("7980a_44194"):
                messagebox.showerror("Sai định dạng", "Tên file phải bắt đầu bằng: '7980a_44194'")
                return
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, filepath)

    # --- [GHI CHÚ SỬA ĐỔI]: NÂNG CẤP HÀM ĐỂ TỰ ĐỘNG FOCUS VÀO FILE ---
    def mo_thu_muc_chua(self):
        if self.last_output_path and os.path.exists(self.last_output_path):
            # Dùng subprocess gọi explorer và bôi đen file
            duong_dan_chuan = os.path.normpath(self.last_output_path)
            subprocess.run(['explorer', '/select,', duong_dan_chuan])
        else:
            messagebox.showwarning("Cảnh báo", "Không tìm thấy file hoặc thư mục đã bị xóa!")

    def update_status(self, message, color):
        self.lbl_status.configure(text=message, text_color=color)
        self.update() 

    def chay_xu_ly(self):
        f_vp = self.entry_vp.get()
        f_7980 = self.entry_7980.get()
        if not f_7980 or not f_vp:
            messagebox.showwarning("Lỗi", "Vui lòng chọn đủ 2 thông tin!")
            return
        
        tu_ngay_str = self.cal_tu_ngay.get_date().strftime('%d%m%Y')
        den_ngay_str = self.cal_den_ngay.get_date().strftime('%d%m%Y')
        ten_mac_dinh = f"File_Tong_Hop_VP_7980_Tu_{tu_ngay_str}_Den_{den_ngay_str}.xlsx"
        
        file_xuat_du_kien = filedialog.asksaveasfilename(initialfile=ten_mac_dinh, defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not file_xuat_du_kien: return
        
        thanh_cong, ket_qua = xu_ly_doi_chieu_4210(f_vp, f_7980, file_xuat_du_kien, self.cal_tu_ngay.get_date(), self.cal_den_ngay.get_date(), self.update_status)
        
        if thanh_cong:
            self.last_output_path = ket_qua 
            self.btn_mo_thu_muc.configure(state="normal") 
            messagebox.showinfo("Thành công", f"Đã xuất file đối chiếu thành công!")
        else:
            messagebox.showerror("Lỗi", ket_qua)

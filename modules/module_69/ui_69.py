import os
import re
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
import pandas as pd
from core.date_entry import DateEntry

from .logic_69 import xu_ly_chuyen_de_69


class UI_69(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.last_output_path = None
        self.huong_dan_hien = True
        self._tao_layout()

    def _tao_layout(self):
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self._tao_phan_input()
        self._tao_phan_huong_dan()
        self._tao_phan_du_lieu()

    def _tao_phan_input(self):
        frame = ctk.CTkFrame(self, corner_radius=12)
        frame.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 8))
        frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(frame, text="Phần 1 - Input file", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=14, pady=(10, 12)
        )
        self.entry_xml1 = self._dong_file(frame, 1, "File XML1:", "Bat buoc: XML1_...", "xml1")
        self.entry_xml3 = self._dong_file(frame, 2, "File XML3:", "Bat buoc: XML3_...", "xml3")
        self.entry_xml4 = self._dong_file(frame, 3, "File XML4:", "Bat buoc: XML4_...", "xml4")

    def _dong_file(self, frame, row, label, placeholder, key):
        ctk.CTkLabel(frame, text=label).grid(row=row, column=0, padx=14, pady=8, sticky="w")
        entry = ctk.CTkEntry(frame, placeholder_text=placeholder)
        entry.grid(row=row, column=1, padx=10, pady=8, sticky="ew")
        ctk.CTkButton(frame, text="Chọn file", width=100, command=lambda: self.chon_file(entry, key)).grid(row=row, column=2, padx=14, pady=8)
        return entry

    def _tao_phan_huong_dan(self):
        self.frame_huong_dan = ctk.CTkFrame(self, corner_radius=12)
        self.frame_huong_dan.grid(row=1, column=0, sticky="ew", padx=12, pady=8)
        self.frame_huong_dan.grid_columnconfigure(0, weight=1)

        top = ctk.CTkFrame(self.frame_huong_dan, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 6))
        top.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(top, text="Phần 2 - Hướng dẫn và thao tác", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, sticky="w")
        ctk.CTkButton(top, text="Thu gọn / Mở rộng", width=130, command=self.toggle_huong_dan).grid(row=0, column=1, sticky="e")

        self.lbl_huong_dan = ctk.CTkLabel(
            self.frame_huong_dan,
            justify="left",
            text="Chọn đủ XML1, XML3, XML4 -> chọn ngày ra viện -> bấm Đối chiếu -> xem dữ liệu phần 3.",
        )
        self.lbl_huong_dan.grid(row=1, column=0, sticky="w", padx=14, pady=4)

        act = ctk.CTkFrame(self.frame_huong_dan, fg_color="transparent")
        act.grid(row=2, column=0, sticky="ew", padx=12, pady=(8, 10))
        ctk.CTkLabel(act, text="Từ ngày:").grid(row=0, column=0, padx=(0, 6))
        self.cal_tu_ngay = DateEntry(act, width=12, date_pattern="dd/mm/yyyy", locale="vi_VN")
        self.cal_tu_ngay.grid(row=0, column=1, padx=(0, 12))
        ctk.CTkLabel(act, text="Đến ngày:").grid(row=0, column=2, padx=(0, 6))
        self.cal_den_ngay = DateEntry(act, width=12, date_pattern="dd/mm/yyyy", locale="vi_VN")
        self.cal_den_ngay.grid(row=0, column=3, padx=(0, 12))
        self.btn_doi_chieu = ctk.CTkButton(act, text="Đối chiếu", fg_color="#16a34a", command=self.chay_xu_ly)
        self.btn_doi_chieu.grid(row=0, column=4, padx=6)
        self.btn_xuat_excel = ctk.CTkButton(act, text="Xuất excel", fg_color="#0284c7", command=self.mo_thu_muc_chua, state="disabled")
        self.btn_xuat_excel.grid(row=0, column=5, padx=6)
        self.lbl_status = ctk.CTkLabel(act, text="Sẵn sàng.")
        self.lbl_status.grid(row=0, column=6, padx=(14, 0), sticky="w")

    def _tao_phan_du_lieu(self):
        frame = ctk.CTkFrame(self, corner_radius=12)
        frame.grid(row=2, column=0, sticky="nsew", padx=12, pady=(8, 12))
        frame.grid_rowconfigure(1, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(frame, text="Phần 3 - Hiển thị dữ liệu", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, sticky="w", padx=14, pady=(10, 8))

        table_wrap = ctk.CTkFrame(frame, fg_color="transparent")
        table_wrap.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        table_wrap.grid_rowconfigure(0, weight=1)
        table_wrap.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        style.configure("Treeview", rowheight=24, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))

        self.data_table = ttk.Treeview(table_wrap, show="headings")
        self.data_table.grid(row=0, column=0, sticky="nsew")
        y_scroll = ttk.Scrollbar(table_wrap, orient="vertical", command=self.data_table.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(table_wrap, orient="horizontal", command=self.data_table.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        self.data_table.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

    def _render_dataframe(self, df):
        self.data_table.delete(*self.data_table.get_children())
        cols = [str(c) for c in df.columns]
        self.data_table["columns"] = cols
        for col in cols:
            self.data_table.heading(col, text=col)
            self.data_table.column(col, width=150, minwidth=120, stretch=False, anchor="w")
        for row in df.fillna("").astype(str).values.tolist():
            self.data_table.insert("", "end", values=row)

    def toggle_huong_dan(self):
        self.huong_dan_hien = not self.huong_dan_hien
        if self.huong_dan_hien:
            self.lbl_huong_dan.grid()
        else:
            self.lbl_huong_dan.grid_remove()

    def chon_file(self, entry_widget, loai_file):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not filepath:
            return
        ten_file = os.path.basename(filepath)
        mau_ten = {"xml1": r"^xml_?1(?:_|\.|$)", "xml3": r"^xml_?3(?:_|\.|$)", "xml4": r"^xml_?4(?:_|\.|$)"}[loai_file]
        if not re.match(mau_ten, ten_file, flags=re.IGNORECASE):
            messagebox.showerror("Sai định dạng file", f"Vui lòng chọn đúng file {loai_file.upper()}.")
            return
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, filepath)

    def mo_thu_muc_chua(self):
        if self.last_output_path and os.path.exists(self.last_output_path):
            subprocess.run(["explorer", "/select,", os.path.normpath(self.last_output_path)])
        else:
            messagebox.showwarning("Cảnh báo", "Không tìm thấy file kết quả.")

    def update_status(self, message, color):
        self.lbl_status.configure(text=message, text_color=color)
        self.update_idletasks()

    def _hien_thi_preview(self, path):
        try:
            df = pd.read_excel(path)
            self._render_dataframe(df.head(300))
        except Exception as e:
            self.data_table.delete(*self.data_table.get_children())
            self.data_table["columns"] = ["thong_bao"]
            self.data_table.heading("thong_bao", text="Thông báo")
            self.data_table.column("thong_bao", width=900, anchor="w")
            self.data_table.insert("", "end", values=[f"Không đọc được dữ liệu xem trước: {e}"])

    def chay_xu_ly(self):
        f_xml1 = self.entry_xml1.get()
        f_xml3 = self.entry_xml3.get()
        f_xml4 = self.entry_xml4.get()
        if not f_xml1 or not f_xml3 or not f_xml4:
            messagebox.showwarning("Lỗi", "Vui lòng chọn đủ 3 file XML1, XML3, XML4.")
            return
        tu_ngay_str = self.cal_tu_ngay.get_date().strftime("%d%m%Y")
        den_ngay_str = self.cal_den_ngay.get_date().strftime("%d%m%Y")
        ten_mac_dinh = f"Chuyen_De_69_Tu_{tu_ngay_str}_Den_{den_ngay_str}.xlsx"
        file_xuat = filedialog.asksaveasfilename(initialfile=ten_mac_dinh, defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not file_xuat:
            return

        self.btn_doi_chieu.configure(state="disabled")
        self.btn_xuat_excel.configure(state="disabled")
        self.update_status("Đang kiểm tra Chuyên đề 69...", "blue")
        thanh_cong, ket_qua = xu_ly_chuyen_de_69(
            f_xml1,
            f_xml3,
            f_xml4,
            "",
            "",
            file_xuat,
            self.cal_tu_ngay.get_date(),
            self.cal_den_ngay.get_date(),
            lambda message, color: self.update_status(message, color),
        )
        self.btn_doi_chieu.configure(state="normal")
        if thanh_cong:
            self.last_output_path = ket_qua
            self.btn_xuat_excel.configure(state="normal")
            self._hien_thi_preview(ket_qua)
            self.update_status("Hoàn thành.", "green")
            messagebox.showinfo("Thành công", "Đã xuất file kết quả Chuyên đề 69.")
        else:
            self.update_status("Lỗi.", "red")
            messagebox.showerror("Lỗi", ket_qua)

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from .logic_so_lieu_chuyen_vien_tuan import (
    tong_hop_so_lieu_chuyen_vien_tuan,
    xuat_excel_so_lieu_chuyen_vien_tuan,
)


class UI_SoLieuChuyenVienTuan(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.file_path = ""
        self.file_danh_muc_path = ""
        self.df_tong_hop = None
        self.df_ngoai_tru = None
        self.df_noi_tru = None
        self.dang_xu_ly = False
        self.tables = {}
        self._tao_layout()

    def _tao_layout(self):
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self._tao_phan_input()
        self._tao_phan_thao_tac()
        self._tao_phan_du_lieu()

    def _tao_phan_input(self):
        frame = ctk.CTkFrame(self, corner_radius=12)
        frame.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 8))
        frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(frame, text="Phần 1 - Input file", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=14, pady=(10, 12)
        )
        ctk.CTkLabel(frame, text="File XML13:").grid(row=1, column=0, padx=14, pady=8, sticky="w")
        self.entry_file = ctk.CTkEntry(frame, placeholder_text="Chọn file bắt đầu bằng XML13")
        self.entry_file.grid(row=1, column=1, padx=10, pady=8, sticky="ew")
        ctk.CTkButton(frame, text="Chọn file", width=100, command=self.chon_file).grid(row=1, column=2, padx=14, pady=8)
        ctk.CTkLabel(frame, text="Danh mục cơ sở KCB:").grid(row=2, column=0, padx=14, pady=8, sticky="w")
        self.entry_file_danh_muc = ctk.CTkEntry(frame, placeholder_text="Chọn file Danh-mục-cơ-sở-KCB")
        self.entry_file_danh_muc.grid(row=2, column=1, padx=10, pady=(0, 12), sticky="ew")
        ctk.CTkButton(frame, text="Chọn file", width=100, command=self.chon_file_danh_muc).grid(
            row=2, column=2, padx=14, pady=(0, 12)
        )

    def _tao_phan_thao_tac(self):
        frame = ctk.CTkFrame(self, corner_radius=12)
        frame.grid(row=1, column=0, sticky="ew", padx=12, pady=8)
        frame.grid_columnconfigure(3, weight=1)
        ctk.CTkLabel(frame, text="Phần 2 - Thao tác", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, columnspan=4, sticky="w", padx=14, pady=(10, 8)
        )
        self.btn_tong_hop = ctk.CTkButton(frame, text="Tổng hợp", fg_color="#16a34a", command=self.chay_tong_hop)
        self.btn_tong_hop.grid(row=1, column=0, padx=(14, 6), pady=(0, 12), sticky="w")
        self.btn_xuat_excel = ctk.CTkButton(frame, text="Xuất Excel", fg_color="#0284c7", state="disabled", command=self.xuat_excel)
        self.btn_xuat_excel.grid(row=1, column=1, padx=6, pady=(0, 12), sticky="w")
        self.lbl_status = ctk.CTkLabel(frame, text="Sẵn sàng.")
        self.lbl_status.grid(row=1, column=2, padx=14, pady=(0, 12), sticky="w")

    def _tao_phan_du_lieu(self):
        frame = ctk.CTkFrame(self, corner_radius=12)
        frame.grid(row=2, column=0, sticky="nsew", padx=12, pady=(8, 12))
        frame.grid_rowconfigure(1, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(frame, text="Phần 3 - Hiển thị dữ liệu", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=14, pady=(10, 8)
        )
        self.tabview = ctk.CTkTabview(frame)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        for name in ["Tổng hợp", "Ngoại trú chuyển tuyến", "Nội trú chuyển tuyến"]:
            tab = self.tabview.add(name)
            tab.grid_rowconfigure(0, weight=1)
            tab.grid_columnconfigure(0, weight=1)
            self.tables[name] = self._tao_bang(tab)

    def _tao_bang(self, master):
        table = ttk.Treeview(master, show="headings")
        table.grid(row=0, column=0, sticky="nsew")
        y_scroll = ttk.Scrollbar(master, orient="vertical", command=table.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(master, orient="horizontal", command=table.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        table.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        return table

    def chon_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")])
        if not filepath:
            return
        if not os.path.basename(filepath).casefold().startswith("xml13"):
            messagebox.showerror("Sai định dạng file", "File phải có tên bắt đầu bằng XML13.")
            return
        self.file_path = filepath
        self.entry_file.delete(0, tk.END)
        self.entry_file.insert(0, filepath)

    def chon_file_danh_muc(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not filepath:
            return
        basename = os.path.basename(filepath).casefold()
        if "danh" not in basename or "kcb" not in basename:
            messagebox.showerror("Sai định dạng file", "File danh mục phải là file danh mục cơ sở KCB.")
            return
        self.file_danh_muc_path = filepath
        self.entry_file_danh_muc.delete(0, tk.END)
        self.entry_file_danh_muc.insert(0, filepath)

    def update_status(self, message, color):
        self.lbl_status.configure(text=message, text_color=color)
        self.update_idletasks()

    def _render_table(self, table, df):
        table.delete(*table.get_children())
        cols = [str(c) for c in df.columns]
        table["columns"] = cols
        for col in cols:
            width = 240 if col in {"Họ tên", "HOTEN_NGUOI_HT", "Địa chỉ bệnh nhân", "Địa chỉ nơi chuyển đến", "Chẩn đoán", "Chỉ tiêu"} else 170
            if col == "Tên bệnh viện":
                width = 260
            table.heading(col, text=col, anchor="w")
            table.column(col, width=width, minwidth=100, stretch=False, anchor="w")
        for _, row in df.fillna("").iterrows():
            table.insert("", "end", values=[str(row[c]) for c in cols])

    def _lay_file_tu_entry(self):
        text = self.entry_file.get().strip()
        if text:
            self.file_path = text.strip('"')
        text_danh_muc = self.entry_file_danh_muc.get().strip()
        if text_danh_muc:
            self.file_danh_muc_path = text_danh_muc.strip('"')

    def chay_tong_hop(self):
        if self.dang_xu_ly:
            return
        self._lay_file_tu_entry()
        if not self.file_path:
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng chọn file XML13.")
            return
        if not os.path.basename(self.file_path).casefold().startswith("xml13"):
            messagebox.showerror("Sai định dạng file", "File phải có tên bắt đầu bằng XML13.")
            return
        if not self.file_danh_muc_path:
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng chọn file danh mục cơ sở KCB.")
            return
        self.dang_xu_ly = True
        self.btn_tong_hop.configure(state="disabled")
        self.btn_xuat_excel.configure(state="disabled")
        self.update_status("Đang tổng hợp dữ liệu...", "blue")
        threading.Thread(target=self._chay_tong_hop_nen, daemon=True).start()

    def _update_status_tu_luong_nen(self, message, color):
        self.after(0, lambda: self.update_status(message, color))

    def _chay_tong_hop_nen(self):
        try:
            ket_qua = tong_hop_so_lieu_chuyen_vien_tuan(
                self.file_path,
                self.file_danh_muc_path,
                self._update_status_tu_luong_nen,
            )
            self.after(0, lambda: self._hoan_tat_tong_hop(True, ket_qua))
        except Exception as exc:
            self.after(0, lambda: self._hoan_tat_tong_hop(False, str(exc)))

    def _hoan_tat_tong_hop(self, thanh_cong, ket_qua):
        self.dang_xu_ly = False
        self.btn_tong_hop.configure(state="normal")
        if not thanh_cong:
            self.update_status("Lỗi khi tổng hợp.", "red")
            messagebox.showerror("Lỗi", ket_qua)
            return
        self.df_tong_hop, self.df_ngoai_tru, self.df_noi_tru = ket_qua
        self._render_table(self.tables["Tổng hợp"], self.df_tong_hop)
        self._render_table(self.tables["Ngoại trú chuyển tuyến"], self.df_ngoai_tru)
        self._render_table(self.tables["Nội trú chuyển tuyến"], self.df_noi_tru)
        self.btn_xuat_excel.configure(state="normal")
        self.update_status("Tổng hợp xong.", "green")

    def xuat_excel(self):
        if self.df_tong_hop is None:
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng bấm Tổng hợp trước khi xuất Excel.")
            return
        file_xuat = filedialog.asksaveasfilename(
            initialfile="So_Lieu_Chuyen_Vien_Trong_Tuan.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not file_xuat:
            return
        try:
            file_da_xuat = xuat_excel_so_lieu_chuyen_vien_tuan(
                self.df_tong_hop.copy(),
                self.df_ngoai_tru.copy(),
                self.df_noi_tru.copy(),
                file_xuat,
            )
            messagebox.showinfo("Thành công", f"Đã xuất file:\n{file_da_xuat}")
            try:
                os.startfile(file_da_xuat)
            except OSError:
                pass
        except Exception as exc:
            messagebox.showerror("Lỗi xuất file", str(exc))

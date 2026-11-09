import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from .logic_01_bh_02_bh import doi_chieu_01_bh_02_bh, xuat_excel_01_bh_02_bh


class UI_01BH02BH(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.file_mau01_list = []
        self.file_mau02 = ""
        self.df_ket_qua = None
        self.dang_xu_ly = False
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

        ctk.CTkLabel(frame, text="File Mẫu 01-BH:").grid(row=1, column=0, padx=14, pady=8, sticky="w")
        self.entry_mau01 = ctk.CTkEntry(frame, placeholder_text="Chọn các file bắt đầu bằng Mau01")
        self.entry_mau01.grid(row=1, column=1, padx=10, pady=8, sticky="ew")
        ctk.CTkButton(frame, text="Chọn file", width=100, command=self.chon_file_mau01).grid(row=1, column=2, padx=14, pady=8)

        ctk.CTkLabel(frame, text="File Mẫu 02:").grid(row=2, column=0, padx=14, pady=8, sticky="w")
        self.entry_pdf = ctk.CTkEntry(frame, placeholder_text="Chọn file PDF/Excel bắt đầu bằng rptTH")
        self.entry_pdf.grid(row=2, column=1, padx=10, pady=8, sticky="ew")
        ctk.CTkButton(frame, text="Chọn file", width=100, command=self.chon_file_pdf).grid(row=2, column=2, padx=14, pady=8)

    def _tao_phan_thao_tac(self):
        frame = ctk.CTkFrame(self, corner_radius=12)
        frame.grid(row=1, column=0, sticky="ew", padx=12, pady=8)
        frame.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(frame, text="Phần 2 - Thao tác", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, columnspan=4, sticky="w", padx=14, pady=(10, 8)
        )
        self.btn_doi_chieu = ctk.CTkButton(frame, text="Đối chiếu", fg_color="#16a34a", command=self.chay_xu_ly)
        self.btn_doi_chieu.grid(row=1, column=0, padx=(14, 6), pady=(0, 12), sticky="w")
        self.btn_xuat_excel = ctk.CTkButton(frame, text="Xuất Excel", fg_color="#0284c7", state="disabled", command=self.xuat_excel)
        self.btn_xuat_excel.grid(row=1, column=1, padx=6, pady=(0, 12), sticky="w")
        self.lbl_status = ctk.CTkLabel(frame, text="Sẵn sàng.")
        self.lbl_status.grid(row=1, column=2, padx=14, pady=(0, 12), sticky="w")

    def _tao_phan_du_lieu(self):
        frame = ctk.CTkFrame(self, corner_radius=12)
        frame.grid(row=2, column=0, sticky="nsew", padx=12, pady=(8, 12))
        frame.grid_rowconfigure(1, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(frame, text="Phần 3 - Kết quả đối chiếu", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=14, pady=(10, 8)
        )

        wrap = ctk.CTkFrame(frame, fg_color="transparent")
        wrap.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        wrap.grid_rowconfigure(0, weight=1)
        wrap.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        style.configure("Treeview", rowheight=26, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))

        self.data_table = ttk.Treeview(wrap, show="headings")
        self.data_table.grid(row=0, column=0, sticky="nsew")
        y_scroll = ttk.Scrollbar(wrap, orient="vertical", command=self.data_table.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(wrap, orient="horizontal", command=self.data_table.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        self.data_table.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

    def _render_dataframe(self, df):
        self.data_table.delete(*self.data_table.get_children())
        cols = [str(c) for c in df.columns if str(c) != "color"]
        self.data_table["columns"] = cols
        for col in cols:
            width = 150
            if col in {"Chỉ tiêu", "Cột Mẫu 02"}:
                width = 280
            elif col in {"Tổng Mẫu 01", "Tổng Mẫu 02", "Lệch (Mẫu 01 - Mẫu 02)"}:
                width = 180
            self.data_table.heading(col, text=col, anchor="w")
            self.data_table.column(col, width=width, minwidth=90, stretch=False, anchor="w")

        self.data_table.tag_configure("red", background="#FF9999")
        for _, row in df.fillna("").iterrows():
            values = [str(row[c]) for c in cols]
            color = str(row.get("color", "none"))
            self.data_table.insert("", "end", values=values, tags=(color,) if color == "red" else ())

    def chon_file_mau01(self):
        filepaths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not filepaths:
            return
        sai = [p for p in filepaths if not os.path.basename(p).casefold().startswith("mau01")]
        if sai:
            messagebox.showerror("Sai định dạng file", "File Mẫu 01-BH phải có tên bắt đầu bằng Mau01.")
            return
        self.file_mau01_list = list(filepaths)
        self.entry_mau01.delete(0, tk.END)
        self.entry_mau01.insert(0, "; ".join(self.file_mau01_list))

    def chon_file_pdf(self):
        filepath = filedialog.askopenfilename(filetypes=[("Mẫu 02 files", "*.pdf *.xlsx *.xls"), ("PDF files", "*.pdf"), ("Excel files", "*.xlsx *.xls")])
        if not filepath:
            return
        if not os.path.basename(filepath).casefold().startswith("rptth"):
            messagebox.showerror("Sai định dạng file", "File Mẫu 02 phải có tên bắt đầu bằng rptTH.")
            return
        self.file_mau02 = filepath
        self.entry_pdf.delete(0, tk.END)
        self.entry_pdf.insert(0, filepath)

    def update_status(self, message, color):
        self.lbl_status.configure(text=message, text_color=color)
        self.update_idletasks()

    def _lay_file_tu_entry(self):
        text_mau01 = self.entry_mau01.get().strip()
        if text_mau01:
            self.file_mau01_list = [p.strip().strip('"') for p in text_mau01.split(";") if p.strip()]
        text_pdf = self.entry_pdf.get().strip()
        if text_pdf:
            self.file_mau02 = text_pdf.strip('"')

    def chay_xu_ly(self):
        if self.dang_xu_ly:
            return
        self._lay_file_tu_entry()
        if not self.file_mau01_list or not self.file_mau02:
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng chọn file Mau01 và file Mẫu 02 rptTH.")
            return
        sai_mau01 = [p for p in self.file_mau01_list if not os.path.basename(p).casefold().startswith("mau01")]
        if sai_mau01:
            messagebox.showerror("Sai định dạng file", "File Mẫu 01-BH phải có tên bắt đầu bằng Mau01.")
            return
        if not os.path.basename(self.file_mau02).casefold().startswith("rptth"):
            messagebox.showerror("Sai định dạng file", "File Mẫu 02 phải có tên bắt đầu bằng rptTH.")
            return

        self.dang_xu_ly = True
        self.btn_doi_chieu.configure(state="disabled")
        self.btn_xuat_excel.configure(state="disabled")
        self.update_status("Đang đối chiếu dữ liệu...", "blue")
        threading.Thread(target=self._chay_xu_ly_nen, daemon=True).start()

    def _update_status_tu_luong_nen(self, message, color):
        self.after(0, lambda: self.update_status(message, color))

    def _chay_xu_ly_nen(self):
        thanh_cong, ket_qua = doi_chieu_01_bh_02_bh(
            self.file_mau01_list,
            self.file_mau02,
            self._update_status_tu_luong_nen,
        )
        self.after(0, lambda: self._hoan_tat_xu_ly(thanh_cong, ket_qua))

    def _hoan_tat_xu_ly(self, thanh_cong, ket_qua):
        self.dang_xu_ly = False
        self.btn_doi_chieu.configure(state="normal")
        if thanh_cong:
            self.df_ket_qua = ket_qua
            self._render_dataframe(self.df_ket_qua)
            self.btn_xuat_excel.configure(state="normal")
        else:
            self.df_ket_qua = None
            self.update_status("Lỗi khi đối chiếu.", "red")
            messagebox.showerror("Lỗi", ket_qua)

    def xuat_excel(self):
        if self.df_ket_qua is None or self.df_ket_qua.empty:
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng bấm Đối chiếu trước khi xuất Excel.")
            return
        file_xuat = filedialog.asksaveasfilename(
            initialfile="Doi_Chieu_01BH_02BH.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not file_xuat:
            return
        try:
            file_da_xuat = xuat_excel_01_bh_02_bh(self.df_ket_qua.copy(), file_xuat)
            messagebox.showinfo("Thành công", f"Đã xuất file:\n{file_da_xuat}")
            try:
                os.startfile(file_da_xuat)
            except OSError:
                pass
        except Exception as exc:
            messagebox.showerror("Lỗi xuất file", str(exc))

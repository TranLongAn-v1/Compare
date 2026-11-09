import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from .logic_loc_ma_dich_vu import doc_va_kiem_tra_file, loc_trung_lap_ma_dich_vu, xuat_excel_loc


class UI_LocMaDichVu(ctk.CTkFrame):
    COT_HIEN_THI = ["MA_DICH_VU", "TEN_DICH_VU", "MA_MAY"]

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.df_nguon = None
        self.df_hien_thi = None
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

        ctk.CTkLabel(frame, text="File tổng hợp:").grid(row=1, column=0, padx=14, pady=8, sticky="w")
        self.entry_file = ctk.CTkEntry(
            frame,
            placeholder_text="Chọn file có đủ cột: MA_DICH_VU, TEN_DICH_VU, MA_MAY, MA_BA, HO_TEN",
        )
        self.entry_file.grid(row=1, column=1, padx=10, pady=8, sticky="ew")
        ctk.CTkButton(frame, text="Chọn file", width=100, command=self.chon_file).grid(row=1, column=2, padx=14, pady=8)

    def _tao_phan_huong_dan(self):
        frame = ctk.CTkFrame(self, corner_radius=12)
        frame.grid(row=1, column=0, sticky="ew", padx=12, pady=8)
        frame.grid_columnconfigure(0, weight=1)

        top = ctk.CTkFrame(frame, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 6))
        top.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(top, text="Phần 2 - Hướng dẫn và thao tác", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, sticky="w"
        )

        self.lbl_huong_dan = ctk.CTkLabel(
            frame,
            justify="left",
            text="Bấm 'Lọc những giá trị trùng lặp' để loại bỏ dòng trùng theo cặp MA_DICH_VU + TEN_DICH_VU.",
        )
        self.lbl_huong_dan.grid(row=1, column=0, sticky="w", padx=14, pady=4)

        action = ctk.CTkFrame(frame, fg_color="transparent")
        action.grid(row=2, column=0, sticky="ew", padx=12, pady=(8, 10))

        self.btn_loc = ctk.CTkButton(action, text="Lọc những giá trị trùng lặp", fg_color="#16a34a", command=self.chay_loc)
        self.btn_loc.grid(row=0, column=0, padx=6)
        self.btn_xuat = ctk.CTkButton(action, text="Xuất file", fg_color="#0284c7", state="disabled", command=self.xuat_file)
        self.btn_xuat.grid(row=0, column=1, padx=6)
        self.lbl_status = ctk.CTkLabel(action, text="Sẵn sàng.")
        self.lbl_status.grid(row=0, column=2, padx=(14, 0), sticky="w")

    def _tao_phan_du_lieu(self):
        frame = ctk.CTkFrame(self, corner_radius=12)
        frame.grid(row=2, column=0, sticky="nsew", padx=12, pady=(8, 12))
        frame.grid_rowconfigure(1, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(frame, text="Phần 3 - Hiển thị dữ liệu", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=14, pady=(10, 8)
        )

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

    def update_status(self, message, color):
        self.lbl_status.configure(text=message, text_color=color)
        self.update_idletasks()

    def _render_dataframe(self, df):
        df = df[[c for c in self.COT_HIEN_THI if c in df.columns]].copy()
        self.data_table.delete(*self.data_table.get_children())
        cols = [str(c) for c in df.columns]
        self.data_table["columns"] = cols
        for col in cols:
            self.data_table.heading(col, text=col, anchor="w")
            self.data_table.column(col, width=180, minwidth=120, stretch=False, anchor="w")

        for _, row in df.fillna("").iterrows():
            self.data_table.insert("", "end", values=[str(row[c]) for c in cols])

    def chon_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not filepath:
            return

        try:
            self.update_status("Đang kiểm tra cột file...", "blue")
            self.df_nguon = doc_va_kiem_tra_file(filepath)
            self.entry_file.delete(0, tk.END)
            self.entry_file.insert(0, filepath)
            self.df_hien_thi = self.df_nguon.copy()
            self._render_dataframe(self.df_hien_thi.head(300))
            self.update_status(f"Đã nạp file hợp lệ: {len(self.df_nguon)} dòng.", "green")
        except Exception as e:
            self.df_nguon = None
            self.df_hien_thi = None
            self.update_status("File không hợp lệ.", "red")
            messagebox.showerror("Sai định dạng file", "File không hợp lệ")

    def chay_loc(self):
        if self.df_nguon is None:
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng chọn file hợp lệ trước.")
            return

        self.btn_loc.configure(state="disabled")
        self.btn_xuat.configure(state="disabled")
        try:
            self.update_status("Đang lọc giá trị trùng lặp...", "blue")
            self.df_hien_thi = loc_trung_lap_ma_dich_vu(self.df_nguon)
            self._render_dataframe(self.df_hien_thi.head(300))
            self.btn_xuat.configure(state="normal")
            self.update_status(f"Lọc xong: còn {len(self.df_hien_thi)} dòng.", "green")
        except Exception as e:
            self.update_status("Lọc thất bại.", "red")
            messagebox.showerror("Lỗi", str(e))
        finally:
            self.btn_loc.configure(state="normal")

    def xuat_file(self):
        if self.df_hien_thi is None or self.df_hien_thi.empty:
            messagebox.showwarning("Không có dữ liệu", "Không có dữ liệu để xuất.")
            return

        file_xuat = filedialog.asksaveasfilename(
            initialfile="Loc_Ma_Dich_Vu.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not file_xuat:
            return

        try:
            file_da_xuat = xuat_excel_loc(self.df_hien_thi.copy(), file_xuat)
            self.update_status("Xuất file thành công.", "green")
            messagebox.showinfo("Thành công", f"Đã xuất file:\n{file_da_xuat}")
        except Exception as e:
            self.update_status("Xuất file thất bại.", "red")
            messagebox.showerror("Lỗi xuất file", str(e))

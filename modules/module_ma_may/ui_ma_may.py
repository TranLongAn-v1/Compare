import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from .logic_ma_may import (
    doc_dm_may,
    doc_file_anh_xa,
    doc_xml3_tong_hop,
    doi_chieu_anh_xa,
    doi_chieu_ma_may,
    xuat_excel_ma_may,
)


class UI_MaMay(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.che_do = "ma_may"
        self.df_file_1 = None
        self.df_file_2 = None
        self.df_ket_qua = None
        self._tao_layout()

    def _tao_layout(self):
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self._tao_phan_input()
        self._tao_phan_huong_dan()
        self._tao_phan_du_lieu()
        self._cap_nhat_giao_dien_theo_che_do()

    def _tao_phan_input(self):
        frame = ctk.CTkFrame(self, corner_radius=12)
        frame.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 8))
        frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(frame, text="Phần 1 - Input file", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=14, pady=(10, 12)
        )

        mode_wrap = ctk.CTkFrame(frame, fg_color="transparent")
        mode_wrap.grid(row=1, column=0, columnspan=3, sticky="w", padx=14, pady=(0, 8))
        ctk.CTkLabel(mode_wrap, text="Option:").grid(row=0, column=0, padx=(0, 8), sticky="w")
        self.mode_selector = ctk.CTkSegmentedButton(
            mode_wrap,
            values=["Mã máy", "Ánh xạ"],
            command=self._chon_che_do,
            width=260,
        )
        self.mode_selector.grid(row=0, column=1, sticky="w")
        self.mode_selector.set("Mã máy")

        self.lbl_file_1 = ctk.CTkLabel(frame, text="")
        self.lbl_file_1.grid(row=2, column=0, padx=14, pady=8, sticky="w")
        self.entry_file_1 = ctk.CTkEntry(frame, placeholder_text="")
        self.entry_file_1.grid(row=2, column=1, padx=10, pady=8, sticky="ew")
        ctk.CTkButton(frame, text="Chọn file", width=100, command=self.chon_file_1).grid(row=2, column=2, padx=14, pady=8)

        self.lbl_file_2 = ctk.CTkLabel(frame, text="")
        self.lbl_file_2.grid(row=3, column=0, padx=14, pady=8, sticky="w")
        self.entry_file_2 = ctk.CTkEntry(frame, placeholder_text="")
        self.entry_file_2.grid(row=3, column=1, padx=10, pady=8, sticky="ew")
        ctk.CTkButton(frame, text="Chọn file", width=100, command=self.chon_file_2).grid(row=3, column=2, padx=14, pady=8)

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

        self.lbl_huong_dan = ctk.CTkLabel(frame, justify="left", text="")
        self.lbl_huong_dan.grid(row=1, column=0, sticky="w", padx=14, pady=4)

        action = ctk.CTkFrame(frame, fg_color="transparent")
        action.grid(row=2, column=0, sticky="ew", padx=12, pady=(8, 10))

        self.btn_doi_chieu = ctk.CTkButton(action, text="Đối chiếu", fg_color="#16a34a", command=self.chay_doi_chieu)
        self.btn_doi_chieu.grid(row=0, column=0, padx=6)
        self.btn_xuat_file = ctk.CTkButton(action, text="Xuất file", fg_color="#0284c7", command=self.xuat_file, state="disabled")
        self.btn_xuat_file.grid(row=0, column=1, padx=6)
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

    def _chon_che_do(self, value):
        self.che_do = "anh_xa" if value == "Ánh xạ" else "ma_may"
        self.df_file_1 = None
        self.df_file_2 = None
        self.df_ket_qua = None
        self.entry_file_1.delete(0, tk.END)
        self.entry_file_2.delete(0, tk.END)
        self.btn_xuat_file.configure(state="disabled")
        self.data_table.delete(*self.data_table.get_children())
        self.data_table["columns"] = []
        self._cap_nhat_giao_dien_theo_che_do()
        self.update_status("Đã đổi option đối chiếu.", "blue")

    def _cap_nhat_giao_dien_theo_che_do(self):
        if self.che_do == "anh_xa":
            self.lbl_file_1.configure(text="File DM_MAY:")
            self.entry_file_1.configure(placeholder_text="Bắt đầu bằng DM_MAY")
            self.lbl_file_2.configure(text="File Ánh xạ:")
            self.entry_file_2.configure(placeholder_text="Đủ cột: STT, TEN_TB, ..., ID")
            self.lbl_huong_dan.configure(
                text=(
                    "Bấm Đối chiếu để lấy MA_MAY từ file Ánh xạ đối chiếu với MAMAY3176 trong DM_MAY. "
                    "Những mã MA_MAY có chứa .44194 sẽ tự động bị loại khỏi kết quả."
                )
            )
        else:
            self.lbl_file_1.configure(text="File Tong_hop_XML3:")
            self.entry_file_1.configure(placeholder_text="Bắt đầu bằng Tong_hop_XML3 và đúng bộ cột chuẩn")
            self.lbl_file_2.configure(text="File DM_MAY:")
            self.entry_file_2.configure(placeholder_text="Bắt đầu bằng DM_MAY")
            self.lbl_huong_dan.configure(
                text=(
                    "Bấm Đối chiếu để so sánh MA_MAY (Tong_hop_XML3) với MAMAY3176 và MAMAY (DM_MAY). "
                    "Kết quả trả về 3 cột Có/Không: Khớp MAMAY3176, Khớp MAMAY, Không tồn tại trong DM."
                )
            )

    def update_status(self, message, color):
        self.lbl_status.configure(text=message, text_color=color)
        self.update_idletasks()

    def _render_dataframe(self, df):
        self.data_table.delete(*self.data_table.get_children())
        cols = [str(c) for c in df.columns]
        self.data_table["columns"] = cols
        for col in cols:
            self.data_table.heading(col, text=col, anchor="w")
            self.data_table.column(col, width=180, minwidth=120, stretch=False, anchor="w")

        for _, row in df.fillna("").iterrows():
            self.data_table.insert("", "end", values=[str(row[c]) for c in cols])

    def chon_file_1(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path:
            return
        try:
            if self.che_do == "anh_xa":
                self.df_file_1 = doc_dm_may(path)
                mo_ta = "DM_MAY"
            else:
                self.df_file_1 = doc_xml3_tong_hop(path)
                mo_ta = "Tong_hop_XML3"
            self.entry_file_1.delete(0, tk.END)
            self.entry_file_1.insert(0, path)
            self.update_status(f"Đã nhận file {mo_ta}: {len(self.df_file_1)} dòng.", "green")
        except Exception:
            self.df_file_1 = None
            self.update_status("File không hợp lệ.", "red")
            messagebox.showerror("Sai định dạng file", "File không hợp lệ")

    def chon_file_2(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path:
            return
        try:
            if self.che_do == "anh_xa":
                self.df_file_2 = doc_file_anh_xa(path)
                mo_ta = "Ánh xạ"
            else:
                self.df_file_2 = doc_dm_may(path)
                mo_ta = "DM_MAY"
            self.entry_file_2.delete(0, tk.END)
            self.entry_file_2.insert(0, path)
            self.update_status(f"Đã nhận file {mo_ta}: {len(self.df_file_2)} dòng.", "green")
        except Exception:
            self.df_file_2 = None
            self.update_status("File không hợp lệ.", "red")
            messagebox.showerror("Sai định dạng file", "File không hợp lệ")

    def chay_doi_chieu(self):
        if self.df_file_1 is None or self.df_file_2 is None:
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng chọn đủ 2 file hợp lệ trước khi đối chiếu.")
            return

        self.btn_doi_chieu.configure(state="disabled")
        self.btn_xuat_file.configure(state="disabled")
        try:
            self.update_status("Đang đối chiếu MA_MAY...", "blue")
            if self.che_do == "anh_xa":
                self.df_ket_qua = doi_chieu_anh_xa(self.df_file_1, self.df_file_2)
            else:
                self.df_ket_qua = doi_chieu_ma_may(self.df_file_1, self.df_file_2)
            self._render_dataframe(self.df_ket_qua.head(300))
            self.btn_xuat_file.configure(state="normal")
            self.update_status(f"Đối chiếu xong: {len(self.df_ket_qua)} dòng.", "green")
        except Exception as e:
            self.df_ket_qua = None
            self.update_status("Đối chiếu thất bại.", "red")
            messagebox.showerror("Lỗi", str(e))
        finally:
            self.btn_doi_chieu.configure(state="normal")

    def xuat_file(self):
        if self.df_ket_qua is None or self.df_ket_qua.empty:
            messagebox.showwarning("Không có dữ liệu", "Không có dữ liệu để xuất.")
            return

        ten_mac_dinh = "Doi_Chieu_Anh_Xa.xlsx" if self.che_do == "anh_xa" else "Doi_Chieu_Ma_May.xlsx"
        file_xuat = filedialog.asksaveasfilename(
            initialfile=ten_mac_dinh,
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not file_xuat:
            return

        try:
            file_da_xuat = xuat_excel_ma_may(self.df_ket_qua.copy(), file_xuat)
            self.update_status("Xuất file thành công.", "green")
            messagebox.showinfo("Thành công", f"Đã xuất file:\n{file_da_xuat}")
        except Exception as e:
            self.update_status("Xuất file thất bại.", "red")
            messagebox.showerror("Lỗi xuất file", str(e))

import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from core.filter_popup import lay_gia_tri_loc_theo_ngu_canh, mo_popup_loc_cot
from .logic_chuan_doan_xml3_pttt_tuoi import (
    COT_HIEN_THI,
    LOI_NGUOI_LON_KE_DICH_VU_NHI,
    LOI_TRE_EM_KE_DICH_VU_NGUOI_LON,
    chuan_doan_xml3_pttt_tuoi,
    xuat_excel_chuan_doan_xml3_pttt_tuoi,
)


class UI_ChuanDoanXML3PTTTTuoi(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.df_ket_qua = None
        self.df_hien_thi = None
        self.current_columns = []
        self.column_filters = {}
        self.filter_popup = None
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

        ctk.CTkLabel(frame, text="File XML3:").grid(row=1, column=0, padx=14, pady=7, sticky="w")
        self.entry_xml3 = ctk.CTkEntry(frame, placeholder_text="Bắt buộc: XML3_...")
        self.entry_xml3.grid(row=1, column=1, padx=10, pady=7, sticky="ew")
        ctk.CTkButton(frame, text="Chọn file", width=100, command=self.chon_file_xml3).grid(row=1, column=2, padx=14, pady=7)

        ctk.CTkLabel(frame, text="File XML1:").grid(row=2, column=0, padx=14, pady=7, sticky="w")
        self.entry_xml1 = ctk.CTkEntry(frame, placeholder_text="Bắt buộc: XML1_...")
        self.entry_xml1.grid(row=2, column=1, padx=10, pady=7, sticky="ew")
        ctk.CTkButton(frame, text="Chọn file", width=100, command=self.chon_file_xml1).grid(row=2, column=2, padx=14, pady=7)

        ctk.CTkLabel(frame, text="File PTTT:").grid(row=3, column=0, padx=14, pady=7, sticky="w")
        self.entry_pttt = ctk.CTkEntry(frame, placeholder_text="Danh mục pttt (1).xlsx")
        self.entry_pttt.grid(row=3, column=1, padx=10, pady=7, sticky="ew")
        ctk.CTkButton(frame, text="Chọn file", width=100, command=self.chon_file_pttt).grid(row=3, column=2, padx=14, pady=7)

    def _tao_phan_huong_dan(self):
        self.frame_huong_dan = ctk.CTkFrame(self, corner_radius=12)
        self.frame_huong_dan.grid(row=1, column=0, sticky="ew", padx=12, pady=8)
        self.frame_huong_dan.grid_columnconfigure(0, weight=1)

        top = ctk.CTkFrame(self.frame_huong_dan, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 6))
        top.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(top, text="Phần 2 - Hướng dẫn và thao tác", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, sticky="w"
        )
        ctk.CTkButton(top, text="Thu gọn / Mở rộng", width=130, command=self.toggle_huong_dan).grid(row=0, column=1, sticky="e")

        self.lbl_huong_dan = ctk.CTkLabel(
            self.frame_huong_dan,
            justify="left",
            text=(
                "Bấm Chuẩn đoán để ghép NGAY_SINH từ XML1 vào XML3 theo MA_LK, đối chiếu danh mục PTTT "
                "và kiểm tra dịch vụ Nhi khoa cho trẻ dưới 16 tuổi."
            ),
        )
        self.lbl_huong_dan.grid(row=1, column=0, sticky="w", padx=14, pady=4)

        action = ctk.CTkFrame(self.frame_huong_dan, fg_color="transparent")
        action.grid(row=2, column=0, sticky="ew", padx=12, pady=(8, 10))

        self.btn_chuan_doan = ctk.CTkButton(action, text="Chuẩn đoán", fg_color="#16a34a", command=self.chay_chuan_doan)
        self.btn_chuan_doan.grid(row=0, column=0, padx=6)
        self.btn_xuat_file = ctk.CTkButton(action, text="Xuất file", fg_color="#0284c7", command=self.xuat_file, state="disabled")
        self.btn_xuat_file.grid(row=0, column=1, padx=6)
        self.btn_xoa_tat_ca_loc = ctk.CTkButton(action, text="Xóa tất cả lọc", fg_color="#64748b", command=self.xoa_tat_ca_loc, state="disabled")
        self.btn_xoa_tat_ca_loc.grid(row=0, column=2, padx=6)
        self.lbl_status = ctk.CTkLabel(action, text="Sẵn sàng.")
        self.lbl_status.grid(row=0, column=3, padx=(14, 0), sticky="w")

    def _tao_phan_du_lieu(self):
        frame = ctk.CTkFrame(self, corner_radius=12)
        frame.grid(row=2, column=0, sticky="nsew", padx=12, pady=(8, 12))
        frame.grid_rowconfigure(1, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(frame, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=14, pady=(10, 8))
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(header, text="Phần 3 - Hiển thị dữ liệu", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, sticky="w"
        )
        self.lbl_tong_so = ctk.CTkLabel(
            header,
            text="Tổng số dòng: 0 | Lỗi: 0",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=("#92400e", "#fde68a"),
        )
        self.lbl_tong_so.grid(row=0, column=1, sticky="e")

        table_wrap = ctk.CTkFrame(frame, fg_color="transparent")
        table_wrap.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        table_wrap.grid_rowconfigure(0, weight=1)
        table_wrap.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        style.configure("Treeview", rowheight=24, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))

        self.data_table = ttk.Treeview(table_wrap, show="headings")
        self.data_table.grid(row=0, column=0, sticky="nsew")
        self.data_table.bind("<Button-1>", self._on_tree_click)
        self.data_table.tag_configure("loi_nguoi_lon", background="#d8b4fe", foreground="#111827")
        self.data_table.tag_configure("loi_tre_em", background="#fdba74", foreground="#111827")

        y_scroll = ttk.Scrollbar(table_wrap, orient="vertical", command=self.data_table.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(table_wrap, orient="horizontal", command=self.data_table.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        self.data_table.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

    def toggle_huong_dan(self):
        self.huong_dan_hien = not self.huong_dan_hien
        if self.huong_dan_hien:
            self.lbl_huong_dan.grid()
        else:
            self.lbl_huong_dan.grid_remove()

    def _chon_file(self, entry, prefix=None):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")])
        if not filepath:
            return
        if prefix:
            ten_file = os.path.basename(filepath)
            if not re.match(rf"^{re.escape(prefix)}(?:_|\.|$)", ten_file, flags=re.IGNORECASE):
                messagebox.showerror("Sai định dạng file", f"Vui lòng chọn file có tên bắt đầu bằng {prefix.upper()}")
                return
        entry.delete(0, tk.END)
        entry.insert(0, filepath)

    def chon_file_xml3(self):
        self._chon_file(self.entry_xml3, "xml3")

    def chon_file_xml1(self):
        self._chon_file(self.entry_xml1, "xml1")

    def chon_file_pttt(self):
        self._chon_file(self.entry_pttt)

    def update_status(self, message, color):
        self.lbl_status.configure(text=message, text_color=color)
        self.update_idletasks()

    def _tao_text_tong_so(self, df):
        if df is None or df.empty:
            return "Tổng số dòng: 0 | Lỗi: 0"
        so_loi = int((df["LOI"] != "").sum()) if "LOI" in df.columns else 0
        return f"Tổng số dòng: {len(df)} | Lỗi: {so_loi}"

    def _render_dataframe(self, df):
        cols_hien_thi = [c for c in COT_HIEN_THI if c in df.columns]
        df = df[cols_hien_thi].copy()
        self.data_table.delete(*self.data_table.get_children())
        cols = [str(c) for c in df.columns]
        self.current_columns = cols
        self.data_table["columns"] = cols

        for col in cols:
            self.data_table.heading(col, text=f"{col} ▼", anchor="w")
            if col == "TEN_DICH_VU":
                width = 360
            elif col in {"HO_TEN", "LOI"}:
                width = 260 if col == "HO_TEN" else 390
            else:
                width = 165
            self.data_table.column(col, width=width, minwidth=120, stretch=False, anchor="w")

        for _, row in df.fillna("").iterrows():
            values = [str(row[c]) for c in cols]
            loi = row.get("LOI", "")
            if loi == LOI_NGUOI_LON_KE_DICH_VU_NHI:
                tags = ("loi_nguoi_lon",)
            elif loi == LOI_TRE_EM_KE_DICH_VU_NGUOI_LON:
                tags = ("loi_tre_em",)
            else:
                tags = ()
            self.data_table.insert("", "end", values=values, tags=tags)

    def _on_tree_click(self, event):
        region = self.data_table.identify("region", event.x, event.y)
        if region != "heading":
            return
        col_id = self.data_table.identify_column(event.x)
        if not col_id or not col_id.startswith("#"):
            return
        idx = int(col_id[1:]) - 1
        if idx < 0 or idx >= len(self.current_columns):
            return
        self._mo_popup_loc_col(self.current_columns[idx], event.x_root, event.y_root)
        return "break"

    def _mo_popup_loc_col(self, col_name, x_root, y_root):
        if self.df_hien_thi is None or self.df_hien_thi.empty:
            return
        values = lay_gia_tri_loc_theo_ngu_canh(self.df_ket_qua, self.column_filters, col_name)

        def on_apply(selected_set):
            self.column_filters[col_name] = selected_set
            self._ap_dung_loc_nhieu_cot()

        self.filter_popup = mo_popup_loc_cot(
            self,
            col_name,
            x_root,
            y_root,
            values,
            self.column_filters,
            on_apply,
            popup_cu=self.filter_popup,
        )

    def _ap_dung_loc_nhieu_cot(self):
        if self.df_ket_qua is None or self.df_ket_qua.empty:
            return
        df = self.df_ket_qua.copy()
        for col_name, selected_set in self.column_filters.items():
            if selected_set is None or col_name not in df.columns:
                continue
            series = df[col_name].fillna("").astype(str)
            df = df.loc[series.isin(selected_set)]
        self.df_hien_thi = df
        self._render_dataframe(self.df_hien_thi.head(300))
        tong_so = self._tao_text_tong_so(self.df_hien_thi)
        self.lbl_tong_so.configure(text=tong_so)
        self.update_status(f"Đang lọc. {tong_so}.", "blue")

    def xoa_tat_ca_loc(self):
        if self.df_ket_qua is None:
            return
        self.column_filters = {c: None for c in self.current_columns}
        self.df_hien_thi = self.df_ket_qua.copy()
        self._render_dataframe(self.df_hien_thi.head(300))
        tong_so = self._tao_text_tong_so(self.df_hien_thi)
        self.lbl_tong_so.configure(text=tong_so)
        self.update_status(f"Đã xóa toàn bộ lọc. {tong_so}.", "green")

    def chay_chuan_doan(self):
        file_xml3 = self.entry_xml3.get().strip()
        file_xml1 = self.entry_xml1.get().strip()
        file_pttt = self.entry_pttt.get().strip()
        if not file_xml3 or not file_xml1 or not file_pttt:
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng chọn đủ file XML3, XML1 và PTTT.")
            return

        self.btn_chuan_doan.configure(state="disabled")
        self.btn_xuat_file.configure(state="disabled")
        self.btn_xoa_tat_ca_loc.configure(state="disabled")

        try:
            self.df_ket_qua = chuan_doan_xml3_pttt_tuoi(file_xml3, file_xml1, file_pttt, callback_status=self.update_status)
            self.df_hien_thi = self.df_ket_qua.copy()
            self.column_filters = {str(c): None for c in COT_HIEN_THI}
            self._render_dataframe(self.df_hien_thi.head(300))
            self.btn_xuat_file.configure(state="normal")
            self.btn_xoa_tat_ca_loc.configure(state="normal")
            tong_so = self._tao_text_tong_so(self.df_ket_qua)
            self.lbl_tong_so.configure(text=tong_so)
            self.update_status(f"Chuẩn đoán xong. {tong_so}.", "green")
        except Exception as e:
            self.df_ket_qua = None
            self.df_hien_thi = None
            self.column_filters = {}
            self.lbl_tong_so.configure(text="Tổng số dòng: 0 | Lỗi: 0")
            self.update_status("Chuẩn đoán thất bại.", "red")
            messagebox.showerror("Lỗi", str(e))
        finally:
            self.btn_chuan_doan.configure(state="normal")

    def xuat_file(self):
        if self.df_hien_thi is None or self.df_hien_thi.empty:
            messagebox.showwarning("Không có dữ liệu", "Không có dữ liệu để xuất.")
            return

        file_xuat = filedialog.asksaveasfilename(
            initialfile="Chuan_Doan_XML3_PTTT_Tuoi.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not file_xuat:
            return

        try:
            file_da_xuat = xuat_excel_chuan_doan_xml3_pttt_tuoi(self.df_hien_thi.copy(), file_xuat)
            self.update_status("Xuất file thành công.", "green")
            messagebox.showinfo("Thành công", f"Đã xuất file:\n{file_da_xuat}")
        except Exception as e:
            self.update_status("Xuất file thất bại.", "red")
            messagebox.showerror("Lỗi xuất file", str(e))

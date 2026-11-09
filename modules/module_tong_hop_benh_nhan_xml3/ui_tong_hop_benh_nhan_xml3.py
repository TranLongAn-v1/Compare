import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from core.filter_popup import lay_gia_tri_loc_theo_ngu_canh, mo_popup_loc_cot
from .logic_tong_hop_benh_nhan_xml3 import (
    COT_HIEN_THI,
    COT_LOC_TRUNG_THEO_NGAY,
    loc_trung_benh_nhan,
    tong_hop_benh_nhan_xml3,
    xuat_excel_tong_hop_benh_nhan,
)


class UI_TongHopBenhNhanXML3(ctk.CTkFrame):
    def __init__(
        self,
        master,
        ma_nhom_can_lay=13,
        ten_hien_thi="bệnh nhân tới khám",
        ten_dich_vu_chua=None,
        ten_dich_vu_chua_theo_ma_nhom=None,
        ten_file_xuat="Tong_Hop_Benh_Nhan_Toi_Kham.xlsx",
        **kwargs,
    ):
        super().__init__(master, **kwargs)
        self.ma_nhom_can_lay = ma_nhom_can_lay
        self.ten_hien_thi = ten_hien_thi
        self.ten_dich_vu_chua = ten_dich_vu_chua
        self.ten_dich_vu_chua_theo_ma_nhom = ten_dich_vu_chua_theo_ma_nhom or {}
        self.ten_file_xuat = ten_file_xuat
        self.df_ket_qua = None
        self.df_hien_thi = None
        self.current_columns = []
        self.column_filters = {}
        self.filter_popup = None
        self.huong_dan_hien = True
        self._tao_layout()

    def _tao_text_tong_so(self, df, che_do="tong_hop"):
        if df is None or df.empty:
            return "Tổng số dòng: 0 | Bệnh nhân: 0"

        so_dong = len(df)
        cot_dem_benh_nhan = COT_LOC_TRUNG_THEO_NGAY if "_NGAY_TH_YL_NGAY" in df.columns else ["MA_BA", "HO_TEN", "MA_THE"]
        if not all(c in df.columns for c in cot_dem_benh_nhan):
            return f"Tổng số dòng: {so_dong}"

        so_benh_nhan = len(df[cot_dem_benh_nhan].drop_duplicates())
        if che_do == "loc_trung":
            return f"Bệnh nhân duy nhất: {so_benh_nhan} | Dòng sau lọc trùng: {so_dong}"
        if isinstance(self.ma_nhom_can_lay, (list, tuple, set)):
            dieu_kien = f"MA_NHOM = {', '.join(str(v) for v in self.ma_nhom_can_lay)}"
        else:
            dieu_kien = f"MA_NHOM = {self.ma_nhom_can_lay}"
        if self.ten_dich_vu_chua:
            cum_tu_loc = self.ten_dich_vu_chua if isinstance(self.ten_dich_vu_chua, (list, tuple, set)) else [self.ten_dich_vu_chua]
            dieu_kien += f", TEN_DICH_VU có {', '.join([repr(cum_tu) for cum_tu in cum_tu_loc])}"
        if self.ten_dich_vu_chua_theo_ma_nhom:
            dieu_kien += ", lọc TEN_DICH_VU theo nhóm"
        return f"Tổng số dòng {dieu_kien}: {so_dong} | Bệnh nhân: {so_benh_nhan}"

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

        ctk.CTkLabel(frame, text="File dữ liệu:").grid(row=1, column=0, padx=14, pady=8, sticky="w")
        self.entry_file = ctk.CTkEntry(
            frame,
            placeholder_text="Chọn file có MA_BA, HO_TEN, MA_THE, NGAY_TH_YL, SO_PHIEU/SOPHIEU, TEN_DICH_VU",
        )
        self.entry_file.grid(row=1, column=1, padx=10, pady=8, sticky="ew")
        ctk.CTkButton(frame, text="Chọn file", width=100, command=self.chon_file).grid(row=1, column=2, padx=14, pady=8)

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
                f"Bấm Tổng hợp để lấy dữ liệu {self.ten_hien_thi} theo MA_NHOM = {self.ma_nhom_can_lay}. "
                "Bấm Lọc trùng bệnh nhân để loại dòng trùng và giữ danh sách bệnh nhân duy nhất theo MA_BA + HO_TEN + MA_THE + ngày trong NGAY_TH_YL."
            ),
        )
        self.lbl_huong_dan.grid(row=1, column=0, sticky="w", padx=14, pady=4)

        action = ctk.CTkFrame(self.frame_huong_dan, fg_color="transparent")
        action.grid(row=2, column=0, sticky="ew", padx=12, pady=(8, 10))

        self.btn_tong_hop = ctk.CTkButton(action, text="Tổng hợp", fg_color="#16a34a", command=self.chay_tong_hop)
        self.btn_tong_hop.grid(row=0, column=0, padx=6)
        self.btn_loc_trung = ctk.CTkButton(
            action,
            text="Lọc trùng bệnh nhân",
            fg_color="#f59e0b",
            text_color="#111827",
            command=self.chay_loc_trung,
            state="disabled",
        )
        self.btn_loc_trung.grid(row=0, column=1, padx=6)
        self.btn_xuat_file = ctk.CTkButton(action, text="Xuất file", fg_color="#0284c7", command=self.xuat_file, state="disabled")
        self.btn_xuat_file.grid(row=0, column=2, padx=6)
        self.btn_xoa_tat_ca_loc = ctk.CTkButton(
            action,
            text="Xóa tất cả lọc",
            fg_color="#64748b",
            command=self.xoa_tat_ca_loc,
            state="disabled",
        )
        self.btn_xoa_tat_ca_loc.grid(row=0, column=3, padx=6)
        self.lbl_status = ctk.CTkLabel(action, text="Sẵn sàng.")
        self.lbl_status.grid(row=0, column=4, padx=(14, 0), sticky="w")

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
            text="Tổng số dòng: 0 | Bệnh nhân: 0",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=("#0f766e", "#5eead4"),
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

    def chon_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")])
        if not filepath:
            return
        self.entry_file.delete(0, tk.END)
        self.entry_file.insert(0, filepath)

    def update_status(self, message, color):
        self.lbl_status.configure(text=message, text_color=color)
        self.update_idletasks()

    def _render_dataframe(self, df):
        cols_hien_thi = [c for c in COT_HIEN_THI if c in df.columns]
        df = df[cols_hien_thi].copy()
        self.data_table.delete(*self.data_table.get_children())
        cols = [str(c) for c in df.columns]
        self.current_columns = cols
        self.data_table["columns"] = cols

        for col in cols:
            self.data_table.heading(col, text=f"{col} ▼", anchor="w")
            width = 260 if col == "TEN_DICH_VU" else 190
            self.data_table.column(col, width=width, minwidth=120, stretch=False, anchor="w")

        for _, row in df.fillna("").iterrows():
            self.data_table.insert("", "end", values=[str(row[c]) for c in cols])

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
        self.update_status(f"Đang lọc: {len(self.df_hien_thi)} dòng.", "blue")

    def xoa_tat_ca_loc(self):
        if self.df_ket_qua is None:
            return
        self.column_filters = {c: None for c in self.current_columns}
        self.df_hien_thi = self.df_ket_qua.copy()
        self._render_dataframe(self.df_hien_thi.head(300))
        self.update_status(f"Đã xóa toàn bộ lọc. Tổng {len(self.df_hien_thi)} dòng.", "green")

    def chay_tong_hop(self):
        path_file = self.entry_file.get().strip()
        if not path_file:
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng chọn file dữ liệu.")
            return

        self.btn_tong_hop.configure(state="disabled")
        self.btn_loc_trung.configure(state="disabled")
        self.btn_xuat_file.configure(state="disabled")
        self.btn_xoa_tat_ca_loc.configure(state="disabled")

        try:
            self.df_ket_qua = tong_hop_benh_nhan_xml3(
                path_file,
                ma_nhom_can_lay=self.ma_nhom_can_lay,
                ten_dich_vu_chua=self.ten_dich_vu_chua,
                ten_dich_vu_chua_theo_ma_nhom=self.ten_dich_vu_chua_theo_ma_nhom,
                callback_status=self.update_status,
            )
            self.df_hien_thi = self.df_ket_qua.copy()
            self.column_filters = {str(c): None for c in COT_HIEN_THI}
            self._render_dataframe(self.df_hien_thi.head(300))
            self.btn_loc_trung.configure(state="normal")
            self.btn_xuat_file.configure(state="normal")
            self.btn_xoa_tat_ca_loc.configure(state="normal")
            tong_so = self._tao_text_tong_so(self.df_ket_qua)
            self.lbl_tong_so.configure(text=tong_so)
            self.update_status(f"Tổng hợp xong. {tong_so}.", "green")
        except Exception as e:
            self.df_ket_qua = None
            self.df_hien_thi = None
            self.column_filters = {}
            self.update_status("Tổng hợp thất bại.", "red")
            messagebox.showerror("Lỗi", str(e))
        finally:
            self.btn_tong_hop.configure(state="normal")

    def chay_loc_trung(self):
        if self.df_ket_qua is None or self.df_ket_qua.empty:
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng tổng hợp dữ liệu trước.")
            return

        self.btn_loc_trung.configure(state="disabled")
        try:
            self.update_status("Đang lọc trùng bệnh nhân...", "blue")
            self.df_ket_qua = loc_trung_benh_nhan(self.df_ket_qua)
            self.df_hien_thi = self.df_ket_qua.copy()
            self.column_filters = {str(c): None for c in COT_HIEN_THI}
            self._render_dataframe(self.df_hien_thi.head(300))
            tong_so = self._tao_text_tong_so(self.df_hien_thi, che_do="loc_trung")
            self.lbl_tong_so.configure(text=tong_so)
            self.update_status(f"Lọc trùng xong. {tong_so}.", "green")
        except Exception as e:
            self.update_status("Lọc trùng thất bại.", "red")
            messagebox.showerror("Lỗi", str(e))
        finally:
            self.btn_loc_trung.configure(state="normal")

    def xuat_file(self):
        if self.df_hien_thi is None or self.df_hien_thi.empty:
            messagebox.showwarning("Không có dữ liệu", "Không có dữ liệu để xuất.")
            return

        file_xuat = filedialog.asksaveasfilename(
            initialfile=self.ten_file_xuat,
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not file_xuat:
            return

        try:
            file_da_xuat = xuat_excel_tong_hop_benh_nhan(self.df_hien_thi.copy(), file_xuat)
            self.update_status("Xuất file thành công.", "green")
            messagebox.showinfo("Thành công", f"Đã xuất file:\n{file_da_xuat}")
        except Exception as e:
            self.update_status("Xuất file thất bại.", "red")
            messagebox.showerror("Lỗi xuất file", str(e))

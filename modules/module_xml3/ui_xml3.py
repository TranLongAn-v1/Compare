import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from core.filter_popup import lay_gia_tri_loc_theo_ngu_canh, mo_popup_loc_cot
from .logic_xml3 import tong_hop_xml3, xuat_excel_xml3


class UI_XML3(ctk.CTkFrame):
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
        frame_input = ctk.CTkFrame(self, corner_radius=12)
        frame_input.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 8))
        frame_input.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(frame_input, text="Phần 1 - Input file", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=14, pady=(10, 12)
        )

        ctk.CTkLabel(frame_input, text="File XML3:").grid(row=1, column=0, padx=14, pady=8, sticky="w")
        self.entry_xml3 = ctk.CTkEntry(frame_input, placeholder_text="Bắt buộc: XML3_...")
        self.entry_xml3.grid(row=1, column=1, padx=10, pady=8, sticky="ew")
        ctk.CTkButton(frame_input, text="Chọn file", width=100, command=self.chon_file_xml3).grid(row=1, column=2, padx=14, pady=8)

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
                "Bấm Tổng hợp để lấy các cột MA_DICH_VU, TEN_DICH_VU, MA_MAY, MA_BA, SOPHIEU, HO_TEN, NGAY_TH_YL. "
                "Bỏ qua các dòng có MA_NHOM thuộc 3, 8, 10, 13, 14, 15, 18."
            ),
        )
        self.lbl_huong_dan.grid(row=1, column=0, sticky="w", padx=14, pady=4)

        action = ctk.CTkFrame(self.frame_huong_dan, fg_color="transparent")
        action.grid(row=2, column=0, sticky="ew", padx=12, pady=(8, 10))

        self.btn_tong_hop = ctk.CTkButton(action, text="Tổng hợp", fg_color="#16a34a", command=self.chay_tong_hop)
        self.btn_tong_hop.grid(row=0, column=0, padx=6)
        self.btn_xuat_file = ctk.CTkButton(action, text="Xuất file", fg_color="#0284c7", command=self.xuat_file, state="disabled")
        self.btn_xuat_file.grid(row=0, column=1, padx=6)
        self.btn_xoa_tat_ca_loc = ctk.CTkButton(action, text="Xóa tất cả lọc", fg_color="#64748b", command=self.xoa_tat_ca_loc, state="disabled")
        self.btn_xoa_tat_ca_loc.grid(row=0, column=2, padx=6)
        self.lbl_status = ctk.CTkLabel(action, text="Sẵn sàng.")
        self.lbl_status.grid(row=0, column=3, padx=(14, 0), sticky="w")

    def _tao_phan_du_lieu(self):
        frame_data = ctk.CTkFrame(self, corner_radius=12)
        frame_data.grid(row=2, column=0, sticky="nsew", padx=12, pady=(8, 12))
        frame_data.grid_rowconfigure(1, weight=1)
        frame_data.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(frame_data, text="Phần 3 - Hiển thị dữ liệu", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=14, pady=(10, 8)
        )

        table_wrap = ctk.CTkFrame(frame_data, fg_color="transparent")
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

    def chon_file_xml3(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not filepath:
            return
        ten_file = os.path.basename(filepath)
        if not re.match(r"^xml3(?:_|\.|$)", ten_file, flags=re.IGNORECASE):
            messagebox.showerror("Sai định dạng file", "Vui lòng chọn file có tên bắt đầu bằng XML3")
            return
        self.entry_xml3.delete(0, tk.END)
        self.entry_xml3.insert(0, filepath)

    def update_status(self, message, color):
        self.lbl_status.configure(text=message, text_color=color)
        self.update_idletasks()

    def _render_dataframe(self, df):
        self.data_table.delete(*self.data_table.get_children())
        cols = [str(c) for c in df.columns]
        self.current_columns = cols
        self.data_table["columns"] = cols
        for col in cols:
            self.data_table.heading(col, text=f"{col} ▼", anchor="w")
            self.data_table.column(col, width=180, minwidth=120, stretch=False, anchor="w")

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
        file_xml3 = self.entry_xml3.get().strip()
        if not file_xml3:
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng chọn file XML3.")
            return

        self.btn_tong_hop.configure(state="disabled")
        self.btn_xuat_file.configure(state="disabled")
        self.btn_xoa_tat_ca_loc.configure(state="disabled")

        try:
            self.df_ket_qua = tong_hop_xml3(file_xml3, callback_status=self.update_status)
            self.df_hien_thi = self.df_ket_qua.copy()
            self.column_filters = {str(c): None for c in self.df_ket_qua.columns}
            self._render_dataframe(self.df_hien_thi.head(300))
            self.btn_xuat_file.configure(state="normal")
            self.btn_xoa_tat_ca_loc.configure(state="normal")
            self.update_status(f"Tổng hợp xong: {len(self.df_ket_qua)} dòng.", "green")
        except Exception as e:
            self.df_ket_qua = None
            self.df_hien_thi = None
            self.column_filters = {}
            self.update_status("Tổng hợp thất bại.", "red")
            messagebox.showerror("Lỗi", str(e))
        finally:
            self.btn_tong_hop.configure(state="normal")

    def xuat_file(self):
        if self.df_hien_thi is None or self.df_hien_thi.empty:
            messagebox.showwarning("Không có dữ liệu", "Không có dữ liệu để xuất.")
            return

        file_xuat = filedialog.asksaveasfilename(
            initialfile="Tong_hop_XML3.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not file_xuat:
            return

        try:
            file_da_xuat = xuat_excel_xml3(self.df_hien_thi.copy(), file_xuat)
            self.update_status("Xuất file thành công.", "green")
            messagebox.showinfo("Thành công", f"Đã xuất file:\n{file_da_xuat}")
        except Exception as e:
            self.update_status("Xuất file thất bại.", "red")
            messagebox.showerror("Lỗi xuất file", str(e))

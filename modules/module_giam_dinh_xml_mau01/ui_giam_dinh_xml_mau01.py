import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
from core.date_entry import DateEntry

from .logic_giam_dinh_xml_mau01 import tao_du_lieu_doi_chieu_giam_dinh_xml_mau01, xuat_excel_tong_hop


class UI_GiamDinhXMLMau01(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.df_ket_qua = None
        self.df_hien_thi = None
        self.huong_dan_hien = True
        self.current_columns = []
        self.column_filters = {}
        self.filter_popup = None
        self.dang_xu_ly = False
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
        ctk.CTkLabel(frame_input, text="Phan 1 - Input file", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=14, pady=(10, 12)
        )

        ctk.CTkLabel(frame_input, text="File Giam dinh XML:").grid(row=1, column=0, padx=14, pady=8, sticky="w")
        self.entry_giam_dinh_xml = ctk.CTkEntry(frame_input, placeholder_text="Bat buoc: TKDLXMLDADNTT...")
        self.entry_giam_dinh_xml.grid(row=1, column=1, padx=10, pady=8, sticky="ew")
        ctk.CTkButton(frame_input, text="Chon file", width=100, command=lambda: self.chon_file(self.entry_giam_dinh_xml, "giam_dinh_xml")).grid(
            row=1, column=2, padx=14, pady=8
        )

        ctk.CTkLabel(frame_input, text="File Mau 01:").grid(row=2, column=0, padx=14, pady=8, sticky="w")
        self.entry_mau01 = ctk.CTkEntry(frame_input, placeholder_text="Bat buoc: MAU01BH...")
        self.entry_mau01.grid(row=2, column=1, padx=10, pady=8, sticky="ew")
        ctk.CTkButton(frame_input, text="Chon file", width=100, command=lambda: self.chon_file(self.entry_mau01, "mau01")).grid(
            row=2, column=2, padx=14, pady=8
        )

    def _tao_phan_huong_dan(self):
        self.frame_huong_dan = ctk.CTkFrame(self, corner_radius=12)
        self.frame_huong_dan.grid(row=1, column=0, sticky="ew", padx=12, pady=8)
        self.frame_huong_dan.grid_columnconfigure(0, weight=1)

        top = ctk.CTkFrame(self.frame_huong_dan, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 6))
        top.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(top, text="Phan 2 - Huong dan va thao tac", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, sticky="w"
        )
        ctk.CTkButton(top, text="Thu gon / Mo rong", width=130, command=self.toggle_huong_dan).grid(row=0, column=1, sticky="e")

        self.lbl_huong_dan = ctk.CTkLabel(
            self.frame_huong_dan,
            justify="left",
            text="Bam vao tieu de cot (co dau ▼) de loc nhu Excel va co the chon nhieu gia tri.",
        )
        self.lbl_huong_dan.grid(row=1, column=0, sticky="w", padx=14, pady=4)

        action = ctk.CTkFrame(self.frame_huong_dan, fg_color="transparent")
        action.grid(row=2, column=0, sticky="ew", padx=12, pady=(8, 10))
        ctk.CTkLabel(action, text="Tu ngay:").grid(row=0, column=0, padx=(0, 6))
        self.cal_tu_ngay = DateEntry(action, width=12, date_pattern="dd/mm/yyyy", locale="vi_VN")
        self.cal_tu_ngay.grid(row=0, column=1, padx=(0, 12))
        ctk.CTkLabel(action, text="Den ngay:").grid(row=0, column=2, padx=(0, 6))
        self.cal_den_ngay = DateEntry(action, width=12, date_pattern="dd/mm/yyyy", locale="vi_VN")
        self.cal_den_ngay.grid(row=0, column=3, padx=(0, 12))
        self.btn_doi_chieu = ctk.CTkButton(action, text="Doi chieu", fg_color="#16a34a", command=self.chay_xu_ly)
        self.btn_doi_chieu.grid(row=0, column=4, padx=6)
        self.btn_xuat_excel = ctk.CTkButton(action, text="Xuat excel", fg_color="#0284c7", command=self.xuat_excel, state="disabled")
        self.btn_xuat_excel.grid(row=0, column=5, padx=6)
        self.btn_xoa_tat_ca_loc = ctk.CTkButton(action, text="Xoa tat ca loc", fg_color="#64748b", command=self.xoa_tat_ca_loc, state="disabled")
        self.btn_xoa_tat_ca_loc.grid(row=0, column=6, padx=6)
        self.lbl_status = ctk.CTkLabel(action, text="San sang.")
        self.lbl_status.grid(row=0, column=7, padx=(14, 0), sticky="w")

    def _tao_phan_du_lieu(self):
        frame_data = ctk.CTkFrame(self, corner_radius=12)
        frame_data.grid(row=2, column=0, sticky="nsew", padx=12, pady=(8, 12))
        frame_data.grid_rowconfigure(1, weight=1)
        frame_data.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(frame_data, text="Phan 3 - Hien thi du lieu", font=ctk.CTkFont(size=16, weight="bold")).grid(
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

    def _render_dataframe(self, df):
        self.data_table.delete(*self.data_table.get_children())
        cols = [str(c) for c in df.columns if str(c) != "color"]
        self.current_columns = cols
        self.data_table["columns"] = cols
        for col in cols:
            self.data_table.heading(col, text=f"{col} ▼", anchor="w")
            self.data_table.column(col, width=150, minwidth=120, stretch=False, anchor="w")

        self.data_table.tag_configure("blue", background="#B3D9FF")
        self.data_table.tag_configure("yellow", background="#FFFF99")
        self.data_table.tag_configure("red", background="#FF9999")
        self.data_table.tag_configure("orange", background="#FFD580")
        self.data_table.tag_configure("purple", background="#E6CCFF")

        display_df = df.copy()
        if "color" not in display_df.columns:
            display_df["color"] = "none"

        for _, row in display_df.fillna("").iterrows():
            row_color = str(row.get("color", "none"))
            values = [str(row[c]) for c in cols]
            if row_color in {"blue", "yellow", "red", "orange", "purple"}:
                self.data_table.insert("", "end", values=values, tags=(row_color,))
            else:
                self.data_table.insert("", "end", values=values)

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
        if self.df_ket_qua is None or self.df_ket_qua.empty:
            return
        if self.filter_popup is not None and self.filter_popup.winfo_exists():
            self.filter_popup.destroy()

        popup = tk.Toplevel(self)
        popup.title(f"Loc cot: {col_name}")
        popup.geometry(f"340x400+{x_root}+{y_root}")
        popup.transient(self.winfo_toplevel())
        self.filter_popup = popup

        tk.Label(popup, text=f"Gia tri cua cot '{col_name}':", anchor="w").pack(fill="x", padx=10, pady=(10, 6))

        values = sorted({str(v) for v in self.df_ket_qua[col_name].fillna("").astype(str).tolist()})
        listbox = tk.Listbox(popup, selectmode="multiple")
        listbox.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        for val in values:
            listbox.insert("end", val)

        selected_set = self.column_filters.get(col_name)
        if selected_set is None:
            listbox.select_set(0, "end")
        else:
            for i, val in enumerate(values):
                if val in selected_set:
                    listbox.select_set(i)

        btn_frame = tk.Frame(popup)
        btn_frame.pack(fill="x", padx=10, pady=(0, 10))

        def chon_tat_ca():
            listbox.select_set(0, "end")

        def bo_chon():
            listbox.selection_clear(0, "end")

        def ap_dung():
            selected = {values[i] for i in listbox.curselection()}
            self.column_filters[col_name] = None if len(selected) == len(values) else selected
            self._ap_dung_loc_nhieu_cot()
            popup.destroy()

        tk.Button(btn_frame, text="Chon tat ca", command=chon_tat_ca).pack(side="left")
        tk.Button(btn_frame, text="Bo chon", command=bo_chon).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Ap dung", command=ap_dung).pack(side="right")

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
        self.update_status(f"Dang loc: {len(self.df_hien_thi)} dong.", "blue")

    def xoa_tat_ca_loc(self):
        if self.df_ket_qua is None:
            return
        self.column_filters = {c: None for c in self.current_columns}
        self.df_hien_thi = self.df_ket_qua.copy()
        self._render_dataframe(self.df_hien_thi.head(300))
        self.update_status(f"Da xoa toan bo loc. Tong {len(self.df_hien_thi)} dong.", "green")

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
        ten_file_chuan = ten_file.casefold()
        if loai_file == "giam_dinh_xml" and not ten_file_chuan.startswith("TKDLXMLDADNTT".casefold()):
            messagebox.showerror("Sai dinh dang file", "Ten file Giam dinh XML phai bat dau bang TKDLXMLDADNTT")
            return
        if loai_file == "mau01" and not ten_file_chuan.startswith("MAU01BH".casefold()):
            messagebox.showerror("Sai dinh dang file", "Ten file Mau 01 phai bat dau bang MAU01BH")
            return
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, filepath)

    def xuat_excel(self):
        if self.df_ket_qua is None or self.df_ket_qua.empty:
            messagebox.showwarning("Thieu du lieu", "Vui long bam Doi chieu truoc khi xuat Excel.")
            return
        if self.df_hien_thi is None or self.df_hien_thi.empty:
            messagebox.showwarning("Khong co du lieu", "Khong co dong nao de xuat theo bo loc hien tai.")
            return

        tu_ngay_str = self.cal_tu_ngay.get_date().strftime("%d%m%Y")
        den_ngay_str = self.cal_den_ngay.get_date().strftime("%d%m%Y")
        ten_mac_dinh = f"File_Tong_Hop_Giam_Dinh_XML_vs_Mau01_Tu_{tu_ngay_str}_Den_{den_ngay_str}.xlsx"
        thu_muc_xuat = filedialog.askdirectory(title="Chon thu muc xuat file Excel")
        if not thu_muc_xuat:
            return
        file_xuat = os.path.join(thu_muc_xuat, ten_mac_dinh)

        try:
            file_da_xuat = xuat_excel_tong_hop(self.df_hien_thi.copy(), file_xuat)
            messagebox.showinfo("Thanh cong", "Da xuat file Excel thanh cong.")
            try:
                os.startfile(file_da_xuat)
            except OSError:
                pass
        except Exception as e:
            messagebox.showerror("Loi xuat file", str(e))

    def update_status(self, message, color):
        self.lbl_status.configure(text=message, text_color=color)
        self.update_idletasks()

    def chay_xu_ly(self):
        if self.dang_xu_ly:
            return
        f_giam_dinh_xml = self.entry_giam_dinh_xml.get()
        f_mau01 = self.entry_mau01.get()
        if not f_giam_dinh_xml or not f_mau01:
            messagebox.showwarning("Loi", "Vui long chon du 2 file.")
            return
        tu_ngay = self.cal_tu_ngay.get_date()
        den_ngay = self.cal_den_ngay.get_date()

        if self.filter_popup is not None and self.filter_popup.winfo_exists():
            self.filter_popup.destroy()
        self.dang_xu_ly = True
        self.btn_doi_chieu.configure(state="disabled")
        self.btn_xuat_excel.configure(state="disabled")
        self.btn_xoa_tat_ca_loc.configure(state="disabled")
        self.update_status("Dang doi chieu du lieu...", "blue")

        threading.Thread(
            target=self._chay_xu_ly_nen,
            args=(f_giam_dinh_xml, f_mau01, tu_ngay, den_ngay),
            daemon=True,
        ).start()

    def _update_status_tu_luong_nen(self, message, color):
        self.after(0, lambda: self.update_status(message, color))

    def _chay_xu_ly_nen(self, f_giam_dinh_xml, f_mau01, tu_ngay, den_ngay):
        thanh_cong, ket_qua = tao_du_lieu_doi_chieu_giam_dinh_xml_mau01(
            f_giam_dinh_xml,
            f_mau01,
            tu_ngay,
            den_ngay,
            self._update_status_tu_luong_nen,
        )
        self.after(0, lambda: self._hoan_tat_xu_ly(thanh_cong, ket_qua))

    def _hoan_tat_xu_ly(self, thanh_cong, ket_qua):
        self.dang_xu_ly = False
        self.btn_doi_chieu.configure(state="normal")
        if thanh_cong:
            self.df_ket_qua = ket_qua
            self.df_hien_thi = self.df_ket_qua.copy()
            self.column_filters = {str(c): None for c in self.df_ket_qua.columns if str(c) != "color"}
            self._render_dataframe(self.df_hien_thi.head(300))
            self.btn_xuat_excel.configure(state="normal")
            self.btn_xoa_tat_ca_loc.configure(state="normal")
            self.update_status("Doi chieu xong. Bam tieu de cot de loc.", "green")
        else:
            self.df_ket_qua = None
            self.df_hien_thi = None
            self.column_filters = {}
            self.update_status("Loi khi doi chieu.", "red")
            messagebox.showerror("Loi", ket_qua)

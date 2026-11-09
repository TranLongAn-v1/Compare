# modules/module_c79a/ui_c79a.py
import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from core.date_entry import DateEntry

from .logic_c79a import tao_du_lieu_doi_chieu_c79a, xuat_excel_c79a

class UI_C79a(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.last_output_path = None
        self.df_ket_qua = None
        self.df_hien_thi = None
        self.huong_dan_hien = True
        self.current_columns = []
        self.column_filters = {}
        self.filter_popup = None
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
        
        ctk.CTkLabel(frame_input, text="Phần 1 - Input file đối chiếu C79a", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=14, pady=(10, 12)
        )

        ctk.CTkLabel(frame_input, text="File chuẩn (C79a):").grid(row=1, column=0, padx=14, pady=8, sticky="w")
        self.entry_std = ctk.CTkEntry(frame_input, placeholder_text="Chọn file chuẩn (ví dụ: C79a_44194_2026_I.xlsx)")
        self.entry_std.grid(row=1, column=1, padx=10, pady=8, sticky="ew")
        ctk.CTkButton(frame_input, text="Chọn file", width=100, command=lambda: self.chon_file(self.entry_std, "std")).grid(
            row=1, column=2, padx=14, pady=8
        )

        ctk.CTkLabel(frame_input, text="File Cổng GD (C79a-HD):").grid(row=2, column=0, padx=14, pady=8, sticky="w")
        self.entry_por = ctk.CTkEntry(frame_input, placeholder_text="Chọn file cổng giám định (ví dụ: C79a-HD_44194_2026_I.xlsx)")
        self.entry_por.grid(row=2, column=1, padx=10, pady=8, sticky="ew")
        ctk.CTkButton(frame_input, text="Chọn file", width=100, command=lambda: self.chon_file(self.entry_por, "por")).grid(
            row=2, column=2, padx=14, pady=8
        )

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
            text="Lấy file chuẩn làm chuẩn để so sánh với file cổng giám định.\n"
                 "- Bệnh nhân có ở file chuẩn mà thiếu ở cổng giám định: tô màu xanh da trời.\n"
                 "- Bệnh nhân có ở cổng giám định mà thiếu ở file chuẩn: tô màu vàng.\n"
                 "- Trùng khớp bệnh nhân nhưng chênh lệch các cột chi phí: tô màu tím và liệt kê các cột lệch.\n"
                 "- Bấm vào tiêu đề cột (Họ và tên, Mã thẻ BHYT, Ghi chú) để lọc dữ liệu nhanh như Excel.",
        )
        self.lbl_huong_dan.grid(row=1, column=0, sticky="w", padx=14, pady=4)

        action = ctk.CTkFrame(self.frame_huong_dan, fg_color="transparent")
        action.grid(row=2, column=0, sticky="ew", padx=12, pady=(8, 10))
        
        ctk.CTkLabel(action, text="Từ ngày:").grid(row=0, column=0, padx=(0, 6))
        self.cal_tu_ngay = DateEntry(action, width=12, date_pattern="dd/mm/yyyy", locale="vi_VN")
        self.cal_tu_ngay.grid(row=0, column=1, padx=(0, 12))
        
        ctk.CTkLabel(action, text="Đến ngày:").grid(row=0, column=2, padx=(0, 6))
        self.cal_den_ngay = DateEntry(action, width=12, date_pattern="dd/mm/yyyy", locale="vi_VN")
        self.cal_den_ngay.grid(row=0, column=3, padx=(0, 12))
        
        self.btn_doi_chieu = ctk.CTkButton(action, text="Đối chiếu", fg_color="#16a34a", command=self.chay_xu_ly)
        self.btn_doi_chieu.grid(row=0, column=4, padx=6)
        
        self.btn_xuat_excel = ctk.CTkButton(action, text="Xuất excel", fg_color="#0284c7", command=self.xuat_excel, state="disabled")
        self.btn_xuat_excel.grid(row=0, column=5, padx=6)
        
        self.btn_xoa_tat_ca_loc = ctk.CTkButton(action, text="Xóa tất cả lọc", fg_color="#64748b", command=self.xoa_tat_ca_loc, state="disabled")
        self.btn_xoa_tat_ca_loc.grid(row=0, column=6, padx=6)
        
        self.lbl_status = ctk.CTkLabel(action, text="Sẵn sàng.")
        self.lbl_status.grid(row=0, column=7, padx=(14, 0), sticky="w")

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

    def _render_dataframe(self, df):
        self.data_table.delete(*self.data_table.get_children())
        
        # Lấy các cột chính để hiển thị giao diện theo yêu cầu
        cols = ["STT", "Họ và tên", "Mã thẻ BHYT", "Mã ĐK", "Ngày vào", "Ghi chú"]
        self.current_columns = cols
        self.data_table["columns"] = cols
        
        # Thiết lập cột STT nhỏ hơn
        self.data_table.heading("STT", text="STT", anchor="w")
        self.data_table.column("STT", width=60, minwidth=50, stretch=False, anchor="w")
        
        for col in cols[1:]:
            self.data_table.heading(col, text=f"{col} ▼", anchor="w")
            if col == "Ghi chú":
                self.data_table.column(col, width=500, minwidth=300, stretch=True, anchor="w")
            elif col == "Ngày vào":
                self.data_table.column(col, width=120, minwidth=100, stretch=False, anchor="w")
            elif col == "Mã ĐK":
                self.data_table.column(col, width=100, minwidth=80, stretch=False, anchor="w")
            else:
                self.data_table.column(col, width=200, minwidth=150, stretch=False, anchor="w")

        # Thiết lập màu sắc dòng
        self.data_table.tag_configure("blue", background="#B3D9FF")    # Xanh da trời
        self.data_table.tag_configure("yellow", background="#FFFF99")  # Vàng
        self.data_table.tag_configure("purple", background="#E6CCFF")  # Tím

        display_df = df.copy()
        if "color" not in display_df.columns:
            display_df["color"] = "none"

        for _, row in display_df.fillna("").iterrows():
            row_color = str(row.get("color", "none"))
            values = [str(row[c]) for c in cols]
            if row_color in {"blue", "yellow", "purple"}:
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
            
        col_name = self.current_columns[idx]
        if col_name == "STT":
            return
            
        self._mo_popup_loc_col(col_name, event.x_root, event.y_root)
        return "break"

    def _mo_popup_loc_col(self, col_name, x_root, y_root):
        if self.df_ket_qua is None or self.df_ket_qua.empty:
            return
        if self.filter_popup is not None and self.filter_popup.winfo_exists():
            self.filter_popup.destroy()

        popup_width, popup_height = 360, 420
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        right_margin = 24
        bottom_margin = 48

        x_pos = x_root
        y_pos = y_root
        if x_pos + popup_width > screen_w - right_margin:
            x_pos = max(8, screen_w - right_margin - popup_width)
        if y_pos + popup_height > screen_h - bottom_margin:
            y_pos = max(8, screen_h - bottom_margin - popup_height)

        popup = tk.Toplevel(self)
        popup.title(f"Lọc cột: {col_name}")
        popup.geometry(f"{popup_width}x{popup_height}+{x_pos}+{y_pos}")
        popup.transient(self.winfo_toplevel())
        self.filter_popup = popup

        tk.Label(popup, text=f"Giá trị của cột '{col_name}':", anchor="w").pack(fill="x", padx=10, pady=(10, 6))

        values = sorted({str(v) for v in self.df_ket_qua[col_name].fillna("").astype(str).tolist()})
        listbox = tk.Listbox(popup, selectmode="multiple")
        listbox.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        
        # Thêm thanh cuộn cho Listbox
        scrollbar = tk.Scrollbar(listbox, orient="vertical")
        scrollbar.config(command=listbox.yview)
        listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        
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

        tk.Button(btn_frame, text="Chọn tất cả", command=chon_tat_ca).pack(side="left")
        tk.Button(btn_frame, text="Bỏ chọn", command=bo_chon).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Áp dụng", command=ap_dung).pack(side="right")

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
        
        # Đánh lại STT sau khi lọc trên giao diện
        display_df = self.df_hien_thi.copy()
        display_df["STT"] = range(1, len(display_df) + 1)
        
        self._render_dataframe(display_df.head(500))
        self.update_status(f"Đang lọc: {len(self.df_hien_thi)} dòng.", "blue")

    def xoa_tat_ca_loc(self):
        if self.df_ket_qua is None:
            return
        self.column_filters = {c: None for c in self.current_columns if c != "STT"}
        self.df_hien_thi = self.df_ket_qua.copy()
        self._render_dataframe(self.df_hien_thi.head(500))
        self.update_status(f"Đã xóa toàn bộ lọc. Tổng {len(self.df_hien_thi)} dòng.", "green")

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
        ten_file_lower = ten_file.lower()
        
        if loai_file == "std":
            if "c79a" not in ten_file_lower or "hd" in ten_file_lower:
                messagebox.showerror("Sai định dạng file", "Tên file chuẩn phải chứa 'C79a' và không chứa 'HD'")
                return
        elif loai_file == "por":
            if "c79a" not in ten_file_lower or "hd" not in ten_file_lower:
                messagebox.showerror("Sai định dạng file", "Tên file cổng giám định phải chứa cả 'C79a' và 'HD'")
                return
                
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, filepath)

    def xuat_excel(self):
        if self.df_ket_qua is None or self.df_ket_qua.empty:
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng bấm Đối chiếu trước khi xuất Excel.")
            return
        if self.df_hien_thi is None or self.df_hien_thi.empty:
            messagebox.showwarning("Không có dữ liệu", "Không có dòng nào để xuất theo bộ lọc hiện tại.")
            return

        tu_ngay_str = self.cal_tu_ngay.get_date().strftime("%d%m%Y")
        den_ngay_str = self.cal_den_ngay.get_date().strftime("%d%m%Y")
        ten_mac_dinh = f"Doi_Chieu_C79a_Tu_{tu_ngay_str}_Den_{den_ngay_str}.xlsx"
        
        thu_muc_xuat = filedialog.askdirectory(title="Chọn thư mục xuất file Excel")
        if not thu_muc_xuat:
            return
        file_xuat = os.path.join(thu_muc_xuat, ten_mac_dinh)

        try:
            file_da_xuat = xuat_excel_c79a(self.df_hien_thi.copy(), file_xuat)
            self.last_output_path = file_da_xuat
            messagebox.showinfo("Thành công", f"Đã xuất file Excel thành công:\n{os.path.basename(file_da_xuat)}")
            try:
                os.startfile(file_da_xuat)
            except OSError:
                pass
        except Exception as e:
            messagebox.showerror("Lỗi xuất file", str(e))

    def update_status(self, message, color):
        self.lbl_status.configure(text=message, text_color=color)
        self.update_idletasks()

    def chay_xu_ly(self):
        f_std = self.entry_std.get()
        f_por = self.entry_por.get()
        if not f_std or not f_por:
            messagebox.showwarning("Lỗi", "Vui lòng chọn đủ 2 file.")
            return

        self.btn_doi_chieu.configure(state="disabled")
        self.btn_xuat_excel.configure(state="disabled")
        self.btn_xoa_tat_ca_loc.configure(state="disabled")
        self.update_status("Đang đối chiếu dữ liệu...", "blue")

        thanh_cong, ket_qua = tao_du_lieu_doi_chieu_c79a(
            f_std,
            f_por,
            self.cal_tu_ngay.get_date(),
            self.cal_den_ngay.get_date(),
            self.update_status,
        )
        self.btn_doi_chieu.configure(state="normal")
        
        if thanh_cong:
            self.df_ket_qua = ket_qua
            self.df_hien_thi = self.df_ket_qua.copy()
            self.column_filters = {str(c): None for c in ["Họ và tên", "Mã thẻ BHYT", "Mã ĐK", "Ngày vào", "Ghi chú"]}
            self._render_dataframe(self.df_hien_thi.head(500))
            self.btn_xuat_excel.configure(state="normal")
            self.btn_xoa_tat_ca_loc.configure(state="normal")
            self.update_status("Đối chiếu xong. Bấm tiêu đề cột để lọc.", "green")
        else:
            self.df_ket_qua = None
            self.df_hien_thi = None
            self.column_filters = {}
            self.update_status("Lỗi khi đối chiếu.", "red")
            messagebox.showerror("Lỗi", ket_qua)

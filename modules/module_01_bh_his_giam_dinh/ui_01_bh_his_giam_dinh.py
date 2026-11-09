import os
import re
import threading
import tkinter as tk
from decimal import Decimal
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from .logic_01_bh_his_giam_dinh import doi_chieu_01_bh_his_giam_dinh, xuat_excel_01_bh_his_giam_dinh


class UI_01BHHISGiamDinh(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.file_his = ""
        self.file_giam_dinh_list = []
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

        ctk.CTkLabel(frame, text="File 01-BH HIS:").grid(row=1, column=0, padx=14, pady=8, sticky="w")
        self.entry_his = ctk.CTkEntry(frame, placeholder_text="Chọn file bắt đầu bằng 01_MAU_01BH")
        self.entry_his.grid(row=1, column=1, padx=10, pady=8, sticky="ew")
        ctk.CTkButton(frame, text="Chọn file", width=100, command=self.chon_file_his).grid(row=1, column=2, padx=14, pady=8)

        ctk.CTkLabel(frame, text="File 01-BH giám định:").grid(row=2, column=0, padx=14, pady=8, sticky="w")
        self.entry_giam_dinh = ctk.CTkEntry(frame, placeholder_text="Chọn các file bắt đầu bằng Mau01")
        self.entry_giam_dinh.grid(row=2, column=1, padx=10, pady=8, sticky="ew")
        ctk.CTkButton(frame, text="Chọn file", width=100, command=self.chon_file_giam_dinh).grid(row=2, column=2, padx=14, pady=8)

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

        ctk.CTkLabel(frame, text="Phần 3 - Tổng và bệnh nhân lệch", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=14, pady=(10, 8)
        )

        wrap = ctk.CTkFrame(frame, fg_color="transparent")
        wrap.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        wrap.grid_rowconfigure(0, weight=1)
        wrap.grid_columnconfigure(0, weight=1)

        self.table_canvas = tk.Canvas(wrap, bg="#ffffff", highlightthickness=0)
        self.table_canvas.grid(row=0, column=0, sticky="nsew")
        y_scroll = ttk.Scrollbar(wrap, orient="vertical", command=self.table_canvas.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(wrap, orient="horizontal", command=self.table_canvas.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        self.table_canvas.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.table_canvas.bind("<MouseWheel>", self._on_table_mousewheel)

    def _render_dataframe(self, df):
        self.table_canvas.delete("all")
        cols = [str(c) for c in df.columns if str(c) != "color"]
        col_widths = [self._column_width(col) for col in cols]

        header_h = 32
        row_h = 28
        total_w = sum(col_widths)
        total_h = header_h + len(df) * row_h
        self.table_canvas.configure(scrollregion=(0, 0, total_w, total_h))

        x = 0
        for col, width in zip(cols, col_widths):
            self.table_canvas.create_rectangle(x, 0, x + width, header_h, fill="#e2e8f0", outline="#cbd5e1")
            self.table_canvas.create_text(
                x + 6,
                header_h / 2,
                text=self._rut_gon_text(col, width),
                anchor="w",
                font=("Segoe UI", 10, "bold"),
                fill="#0f172a",
            )
            x += width

        row_fills = {"red": "#FF9999", "yellow": "#FFFF99", "blue": "#B3D9FF"}
        for row_idx, (_, row) in enumerate(df.fillna("").iterrows()):
            y1 = header_h + row_idx * row_h
            y2 = y1 + row_h
            bg = row_fills.get(str(row.get("color", "none")), "#ffffff")
            highlight_cols = self._cot_can_to_cam(cols, row)

            x = 0
            for col_idx, (col, width) in enumerate(zip(cols, col_widths)):
                is_highlight = col_idx in highlight_cols
                text_color = "#4C1D95" if is_highlight else "#111827"
                text_font = ("Segoe UI", 10, "bold") if is_highlight else ("Segoe UI", 10)
                self.table_canvas.create_rectangle(x, y1, x + width, y2, fill=bg, outline="#cbd5e1")
                self.table_canvas.create_text(
                    x + 6,
                    y1 + row_h / 2,
                    text=self._rut_gon_text(str(row[col]), width),
                    anchor="w",
                    font=text_font,
                    fill=text_color,
                )
                x += width

    def _column_width(self, col):
        if col in {"Họ tên / Chỉ tiêu", "Trạng thái", "Nguồn giám định"}:
            return 240
        if col in {"Mã thẻ", "Ngày vào", "Ngày ra"}:
            return 155
        if "Lệch" in col or "HIS" in col or "giám định" in col:
            return 165
        return 145

    def _on_table_mousewheel(self, event):
        self.table_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _rut_gon_text(self, text, width):
        max_chars = max(4, int((width - 12) / 7))
        return text if len(text) <= max_chars else text[: max_chars - 1] + "…"

    def _parse_tien_hien_thi(self, val):
        text = re.sub(r"[^0-9,.\-]", "", str(val or "").strip().replace(" ", ""))
        if not text:
            return Decimal("0")
        if "," in text and "." in text:
            text = text.replace(".", "").replace(",", ".")
        elif "," in text:
            text = text.replace(",", ".")
        elif text.count(".") > 1:
            text = text.replace(".", "")
        elif text.count(".") == 1:
            before, after = text.split(".", 1)
            if len(after) == 3 and before.replace("-", "").isdigit():
                text = before + after
        try:
            return Decimal(text)
        except Exception:
            return Decimal("0")

    def _cot_can_to_cam(self, cols, row):
        orange_cols = set()
        for idx, col in enumerate(cols):
            if not (col.startswith("L") and "T_" in col):
                continue
            if abs(self._parse_tien_hien_thi(row[col])) == 0:
                continue
            orange_cols.add(idx)
            if idx - 1 >= 0:
                orange_cols.add(idx - 1)
            if idx - 2 >= 0:
                orange_cols.add(idx - 2)
        return orange_cols

    def chon_file_his(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not filepath:
            return
        if not os.path.basename(filepath).casefold().startswith("01_mau_01bh"):
            messagebox.showerror("Sai định dạng file", "File HIS phải có tên bắt đầu bằng 01_MAU_01BH.")
            return
        self.file_his = filepath
        self.entry_his.delete(0, tk.END)
        self.entry_his.insert(0, filepath)

    def chon_file_giam_dinh(self):
        filepaths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not filepaths:
            return
        sai = [p for p in filepaths if not os.path.basename(p).casefold().startswith("mau01")]
        if sai:
            messagebox.showerror("Sai định dạng file", "File giám định phải có tên bắt đầu bằng Mau01.")
            return
        self.file_giam_dinh_list = list(filepaths)
        self.entry_giam_dinh.delete(0, tk.END)
        self.entry_giam_dinh.insert(0, "; ".join(self.file_giam_dinh_list))

    def update_status(self, message, color):
        self.lbl_status.configure(text=message, text_color=color)
        self.update_idletasks()

    def _lay_file_tu_entry(self):
        text_his = self.entry_his.get().strip()
        if text_his:
            self.file_his = text_his.strip('"')
        text_gd = self.entry_giam_dinh.get().strip()
        if text_gd:
            self.file_giam_dinh_list = [p.strip().strip('"') for p in text_gd.split(";") if p.strip()]

    def chay_xu_ly(self):
        if self.dang_xu_ly:
            return
        self._lay_file_tu_entry()
        if not self.file_his or not self.file_giam_dinh_list:
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng chọn file HIS và file giám định.")
            return
        if not os.path.basename(self.file_his).casefold().startswith("01_mau_01bh"):
            messagebox.showerror("Sai định dạng file", "File HIS phải có tên bắt đầu bằng 01_MAU_01BH.")
            return
        sai_gd = [p for p in self.file_giam_dinh_list if not os.path.basename(p).casefold().startswith("mau01")]
        if sai_gd:
            messagebox.showerror("Sai định dạng file", "File giám định phải có tên bắt đầu bằng Mau01.")
            return

        self.dang_xu_ly = True
        self.btn_doi_chieu.configure(state="disabled")
        self.btn_xuat_excel.configure(state="disabled")
        self.update_status("Đang đối chiếu dữ liệu...", "blue")
        threading.Thread(target=self._chay_xu_ly_nen, daemon=True).start()

    def _update_status_tu_luong_nen(self, message, color):
        self.after(0, lambda: self.update_status(message, color))

    def _chay_xu_ly_nen(self):
        thanh_cong, ket_qua = doi_chieu_01_bh_his_giam_dinh(
            self.file_his,
            self.file_giam_dinh_list,
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
            initialfile="Doi_Chieu_01BH_HIS_vs_Giam_Dinh.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not file_xuat:
            return
        try:
            file_da_xuat = xuat_excel_01_bh_his_giam_dinh(self.df_ket_qua.copy(), file_xuat)
            messagebox.showinfo("Thành công", f"Đã xuất file:\n{file_da_xuat}")
            try:
                os.startfile(file_da_xuat)
            except OSError:
                pass
        except Exception as exc:
            messagebox.showerror("Lỗi xuất file", str(exc))

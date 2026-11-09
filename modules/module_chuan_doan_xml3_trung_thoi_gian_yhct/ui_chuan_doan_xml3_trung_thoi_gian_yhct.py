import tkinter as tk
import threading
import queue
import unicodedata
from tkinter import filedialog, messagebox

import customtkinter as ctk
import pandas as pd

from modules.module_chuan_doan_xml3_so_luot_yhct.ui_chuan_doan_xml3_so_luot_yhct import UI_ChuanDoanXML3SoLuotYHCT

from .logic_chuan_doan_xml3_trung_thoi_gian_yhct import (
    CHE_DO_LOAI_TRU_TRUNG,
    CHE_DO_XUAT_TOAN_TAT_CA,
    COT_KET_QUA,
    _khoa_dich_vu,
    chuan_doan_xml3_trung_thoi_gian_yhct,
    phan_tich_file_xml3,
    xuat_excel_chuan_doan_xml3_trung_thoi_gian_yhct,
)


def _text_tim_kiem(value):
    value = unicodedata.normalize("NFD", str(value).casefold()).replace("đ", "d")
    return "".join(ch for ch in value if not unicodedata.combining(ch))


class UI_ChuanDoanXML3TrungThoiGianYHCT(UI_ChuanDoanXML3SoLuotYHCT):
    def __init__(self, master, **kwargs):
        self.cot_hien_thi = COT_KET_QUA
        self.df_nguon = None
        self.danh_muc = None
        self.file_da_phan_tich = None
        self.popup_dich_vu = None
        self.dong_dich_vu = {}
        self.search_var = None
        self.che_do_chuan_doan = None
        self._job_queue = queue.Queue()
        super().__init__(master, **kwargs)

    def _tao_phan_huong_dan(self):
        self.frame_huong_dan = ctk.CTkFrame(self, corner_radius=12)
        self.frame_huong_dan.grid(row=1, column=0, sticky="ew", padx=12, pady=(2, 5))
        self.frame_huong_dan.grid_columnconfigure(0, weight=1)

        top = ctk.CTkFrame(self.frame_huong_dan, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=12, pady=(8, 2))
        top.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(top, text="Phần 2 - Phân tích và chọn DVKT", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, sticky="w"
        )
        ctk.CTkButton(top, text="Thu gọn / Mở rộng", width=130, command=self.toggle_huong_dan).grid(
            row=0, column=1, sticky="e"
        )

        self.frame_tuy_chon = ctk.CTkFrame(self.frame_huong_dan, fg_color="transparent")
        self.frame_tuy_chon.grid(row=1, column=0, sticky="ew", padx=12, pady=2)
        self.frame_tuy_chon.grid_columnconfigure(0, weight=1)
        self.lbl_danh_muc = ctk.CTkLabel(
            self.frame_tuy_chon,
            text="Chưa phân tích file. Bấm Phân tích file để lấy danh mục MA_NHOM = 18.",
            anchor="w",
        )
        self.lbl_danh_muc.grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(
            self.frame_tuy_chon,
            text="Thời điểm kết thúc là dự kiến theo số phút; XML3 không có giờ kết thúc thực tế.",
            anchor="w", text_color=("gray35", "gray75"),
        ).grid(row=1, column=0, sticky="ew")

        action = ctk.CTkFrame(self.frame_huong_dan, fg_color="transparent")
        action.grid(row=2, column=0, sticky="ew", padx=12, pady=(4, 8))
        self.btn_phan_tich = ctk.CTkButton(
            action, text="Phân tích file", fg_color="#7c3aed", command=self.phan_tich_file
        )
        self.btn_phan_tich.grid(row=0, column=0, padx=4)
        self.btn_mo_danh_muc = ctk.CTkButton(
            action, text="Chọn dịch vụ...", command=self.mo_danh_muc, state="disabled"
        )
        self.btn_mo_danh_muc.grid(row=0, column=1, padx=4)
        self.btn_chuan_doan = ctk.CTkButton(
            action, text="Chuẩn đoán", fg_color="#16a34a",
            command=self.chay_chuan_doan, state="disabled",
        )
        self.btn_chuan_doan.grid(row=0, column=2, padx=4)
        self.btn_xuat_file = ctk.CTkButton(
            action, text="Xuất file", fg_color="#0284c7", command=self.xuat_file, state="disabled"
        )
        self.btn_xuat_file.grid(row=0, column=3, padx=4)
        self.btn_xoa_tat_ca_loc = ctk.CTkButton(
            action, text="Xóa tất cả lọc", fg_color="#64748b",
            command=self.xoa_tat_ca_loc, state="disabled",
        )
        self.btn_xoa_tat_ca_loc.grid(row=0, column=4, padx=4)
        self.lbl_status = ctk.CTkLabel(action, text="Sẵn sàng.")
        self.lbl_status.grid(row=0, column=5, padx=(12, 0), sticky="w")

    def chon_file(self):
        truoc = self.entry_file.get().strip()
        super().chon_file()
        if self.entry_file.get().strip() != truoc:
            self._xoa_phan_tich()

    def _xoa_phan_tich(self):
        if self.popup_dich_vu is not None:
            self.popup_dich_vu.destroy()
            self.popup_dich_vu = None
        self.df_nguon = None
        self.danh_muc = None
        self.file_da_phan_tich = None
        self.dong_dich_vu = {}
        self.search_var = None
        self.che_do_chuan_doan = None
        self.lbl_danh_muc.configure(text="Chưa phân tích file. Bấm Phân tích file để lấy danh mục MA_NHOM = 18.")
        self.btn_mo_danh_muc.configure(state="disabled")
        self.btn_chuan_doan.configure(state="disabled")
        self.btn_xuat_file.configure(state="disabled")
        self.btn_xoa_tat_ca_loc.configure(state="disabled")
        self.df_ket_qua = None
        self.df_hien_thi = None
        self.column_filters = {}
        self.lbl_tong_so.configure(text="Tổng số dòng: 0 | Lỗi: 0")
        self._render_dataframe(pd.DataFrame(columns=COT_KET_QUA))

    def _tao_popup_dich_vu(self):
        popup = ctk.CTkToplevel(self)
        popup.withdraw()
        popup.title("Chọn DVKT nhóm 18 và thời gian thực hiện")
        popup.geometry("1190x680")
        popup.minsize(850, 450)
        popup.grid_columnconfigure(0, weight=1)
        popup.grid_rowconfigure(2, weight=1)
        popup.protocol("WM_DELETE_WINDOW", popup.withdraw)
        self.popup_dich_vu = popup

        toolbar = ctk.CTkFrame(popup)
        toolbar.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 6))
        toolbar.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(toolbar, text="Tìm mã hoặc tên DVKT:").grid(row=0, column=0, padx=8, pady=8)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._loc_dich_vu)
        self.entry_tim_kiem = ctk.CTkEntry(
            toolbar, textvariable=self.search_var, placeholder_text="Ví dụ: 08.0432 hoặc xoa bop"
        )
        self.entry_tim_kiem.grid(row=0, column=1, sticky="ew", padx=8, pady=8)
        ctk.CTkButton(toolbar, text="Chọn kết quả lọc", width=135,
                      command=lambda: self._chon_hang_loat(True, chi_hien_thi=True)).grid(row=0, column=2, padx=4)
        ctk.CTkButton(toolbar, text="Chọn tất cả", width=110,
                      command=lambda: self._chon_hang_loat(True)).grid(row=0, column=3, padx=4)
        ctk.CTkButton(toolbar, text="Bỏ chọn tất cả", width=125,
                      command=lambda: self._chon_hang_loat(False)).grid(row=0, column=4, padx=8)

        header = ctk.CTkFrame(popup, fg_color="transparent")
        header.grid(row=1, column=0, sticky="ew", padx=20)
        header.grid_columnconfigure(2, weight=1)
        for col, (ten, width) in enumerate((
            ("Chọn", 55), ("Mã dịch vụ", 150), ("Tên dịch vụ", 540),
            ("Phút", 75), ("Số dòng", 80),
        )):
            ctk.CTkLabel(header, text=ten, width=width, anchor="w",
                         font=ctk.CTkFont(weight="bold")).grid(row=0, column=col, sticky="w", padx=4)

        list_frame = ctk.CTkScrollableFrame(popup)
        list_frame.grid(row=2, column=0, sticky="nsew", padx=12, pady=4)
        for item in self.danh_muc:
            khoa = _khoa_dich_vu(item["ma"], item["ten"])
            co_mac_dinh = item["phut_mac_dinh"] is not None
            row_frame = ctk.CTkFrame(list_frame, fg_color="transparent")
            row_frame.pack(fill="x", pady=1)
            row_frame.grid_columnconfigure(2, weight=1)
            var = tk.BooleanVar(value=co_mac_dinh)
            ctk.CTkCheckBox(row_frame, text="", width=55, variable=var,
                            command=self._cau_hinh_thay_doi).grid(row=0, column=0, sticky="w", padx=4)
            ctk.CTkLabel(row_frame, text=item["ma"], width=150, anchor="w").grid(
                row=0, column=1, sticky="w", padx=4
            )
            ctk.CTkLabel(row_frame, text=item["ten"], width=540, anchor="w").grid(
                row=0, column=2, sticky="ew", padx=4
            )
            entry = ctk.CTkEntry(row_frame, width=75)
            if co_mac_dinh:
                entry.insert(0, str(item["phut_mac_dinh"]))
            entry.grid(row=0, column=3, sticky="w", padx=4)
            entry.bind("<KeyRelease>", self._cau_hinh_thay_doi)
            entry.bind("<FocusOut>", self._cau_hinh_thay_doi, add="+")
            ctk.CTkLabel(row_frame, text=str(item["so_luot"]), width=80, anchor="w").grid(
                row=0, column=4, sticky="w", padx=4
            )
            self.dong_dich_vu[khoa] = (var, entry, row_frame, item)

        footer = ctk.CTkFrame(popup, fg_color="transparent")
        footer.grid(row=3, column=0, sticky="ew", padx=12, pady=(3, 10))
        footer.grid_columnconfigure(0, weight=1)
        self.lbl_tong_chon = ctk.CTkLabel(footer, text="", anchor="w")
        self.lbl_tong_chon.grid(row=0, column=0, sticky="ew")
        ctk.CTkButton(footer, text="Xong", command=popup.withdraw).grid(row=0, column=1, padx=6)
        self._cap_nhat_so_chon()

    def mo_danh_muc(self):
        if self.popup_dich_vu is None:
            return
        self.popup_dich_vu.deiconify()
        self.popup_dich_vu.lift()
        self.entry_tim_kiem.focus_set()

    def _loc_dich_vu(self, *_args):
        if self.search_var is None:
            return
        query = _text_tim_kiem(self.search_var.get().strip())
        so_hien = 0
        for khoa, (_, _, row_frame, item) in self.dong_dich_vu.items():
            row_frame.pack_forget()
            if query in _text_tim_kiem(f"{item['ma']} {item['ten']}"):
                row_frame.pack(fill="x", pady=1)
                so_hien += 1
        self._cap_nhat_so_chon(so_hien)

    def _chon_hang_loat(self, gia_tri, chi_hien_thi=False):
        truy_van = _text_tim_kiem(self.search_var.get().strip()) if chi_hien_thi and self.search_var else ""
        for var, _, _, item in self.dong_dich_vu.values():
            khop = truy_van in _text_tim_kiem(f"{item['ma']} {item['ten']}")
            if not chi_hien_thi or khop:
                var.set(gia_tri)
            elif gia_tri:
                # Chọn kết quả lọc có nghĩa là chỉ giữ lại các dịch vụ đang khớp.
                var.set(False)
        self._cau_hinh_thay_doi()

    def _cap_nhat_so_chon(self, so_hien=None):
        if self.danh_muc is None:
            return
        so_chon = sum(var.get() for var, _, _, _ in self.dong_dich_vu.values())
        self.lbl_danh_muc.configure(
            text=f"{len(self.danh_muc)} mã–tên nhóm 18 | Đã chọn: {so_chon}. Bấm Chọn dịch vụ... để tìm và sửa phút."
        )
        if self.popup_dich_vu is not None:
            if so_hien is None:
                so_hien = sum(bool(row.winfo_manager()) for _, _, row, _ in self.dong_dich_vu.values())
            self.lbl_tong_chon.configure(
                text=f"Hiển thị {so_hien}/{len(self.danh_muc)} dịch vụ | Đã chọn {so_chon}. "
                     "Dịch vụ không có phút mặc định cần nhập phút trước khi chuẩn đoán."
            )

    def _cau_hinh_thay_doi(self, _event=None):
        self._cap_nhat_so_chon()
        if self.df_ket_qua is None:
            return
        self.df_ket_qua = None
        self.df_hien_thi = None
        self.column_filters = {}
        self.btn_xuat_file.configure(state="disabled")
        self.btn_xoa_tat_ca_loc.configure(state="disabled")
        self.lbl_tong_so.configure(text="Tổng số dòng: 0 | Lỗi: 0")
        self._render_dataframe(pd.DataFrame(columns=COT_KET_QUA))
        self.update_status("Đã đổi lựa chọn/thời lượng; hãy chuẩn đoán lại.", "blue")

    def phan_tich_file(self):
        path_file = self.entry_file.get().strip()
        if not path_file:
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng chọn file XML3.")
            return
        self.btn_phan_tich.configure(state="disabled")
        self.btn_chuan_doan.configure(state="disabled")
        self.update_status("Đang phân tích file XML3...", "blue")
        def worker():
            try:
                self._job_queue.put(("analyze", *phan_tich_file_xml3(path_file)))
            except Exception as exc:
                self._job_queue.put(("error", "analyze", exc))
        threading.Thread(target=worker, daemon=True).start()
        self.after(100, self._poll_job)

    def _poll_job(self):
        try:
            job = self._job_queue.get_nowait()
        except queue.Empty:
            self.after(100, self._poll_job)
            return
        if job[0] == "analyze":
            _, df, danh_muc = job
            self._xoa_phan_tich()
            self.df_nguon, self.danh_muc = df, danh_muc
            self.file_da_phan_tich = self.entry_file.get().strip()
            self._tao_popup_dich_vu()
            self.btn_mo_danh_muc.configure(state="normal")
            self.btn_chuan_doan.configure(state="normal")
            self.btn_phan_tich.configure(state="normal")
            self.update_status(f"Đã phân tích {len(danh_muc)} mã–tên DVKT nhóm 18.", "green")
        elif job[0] == "diagnose":
            _, result = job
            self.df_ket_qua = result
            self.df_hien_thi = result.copy()
            self.column_filters = {str(c): None for c in COT_KET_QUA}
            self._render_dataframe(self.df_hien_thi.head(300))
            self.btn_xuat_file.configure(state="normal" if not result.empty else "disabled")
            self.btn_xoa_tat_ca_loc.configure(state="normal" if not result.empty else "disabled")
            tong_so = self._tao_text_tong_so(result)
            self.lbl_tong_so.configure(text=tong_so)
            self.btn_chuan_doan.configure(state="normal")
            ten_che_do = (
                "Loại trừ dịch vụ trùng" if self.che_do_chuan_doan == CHE_DO_LOAI_TRU_TRUNG
                else "Xuất toán tất cả"
            )
            self.update_status(f"Chuẩn đoán xong ({ten_che_do}). {tong_so}.", "green")
        elif job[0] == "export":
            _, output = job
            self.btn_xuat_file.configure(state="normal")
            self.update_status("Xuất file thành công.", "green")
            messagebox.showinfo("Thành công", f"Đã xuất file:\n{output}")
        else:
            _, action, exc = job
            if action == "analyze":
                self._xoa_phan_tich()
                self.btn_phan_tich.configure(state="normal")
            elif action == "diagnose":
                self.btn_chuan_doan.configure(state="normal")
                self.df_ket_qua = self.df_hien_thi = None
            else:
                self.btn_xuat_file.configure(state="normal")
            self.update_status(f"{action.capitalize()} thất bại.", "red")
            messagebox.showerror("Lỗi", str(exc))

    def _tao_text_tong_so(self, df):
        if df is None or df.empty:
            return "Tổng số dòng: 0 | Lỗi: 0"
        so_loi = int(df["SO_DVKT_TRUNG"].gt(0).sum())
        return f"Tổng số dòng: {len(df)} | Lỗi: {so_loi}"

    def chay_chuan_doan(self):
        path_file = self.entry_file.get().strip()
        if self.df_nguon is None or self.danh_muc is None or self.file_da_phan_tich != path_file:
            messagebox.showwarning("Chưa phân tích", "Hãy phân tích lại file XML3 đang chọn trước khi chuẩn đoán.")
            return
        self.btn_chuan_doan.configure(state="disabled")
        popup = ctk.CTkToplevel(self)
        popup.title("Chọn chế độ chuẩn đoán")
        popup.geometry("620x245")
        popup.resizable(False, False)
        popup.transient(self.winfo_toplevel())
        frame = ctk.CTkFrame(popup)
        frame.pack(fill="both", expand=True, padx=14, pady=14)
        ctk.CTkLabel(
            frame, text="Chọn cách tính xuất toán", font=ctk.CTkFont(size=17, weight="bold")
        ).pack(pady=(14, 8))
        ctk.CTkLabel(
            frame,
            text="Loại trừ: cùng người thực hiện, cùng mã và tên dịch vụ, trùng thời gian thì chỉ giữ lượt bắt đầu sớm hơn.",
            wraplength=565, justify="left",
        ).pack(padx=16, pady=(0, 3))
        ctk.CTkLabel(
            frame, text="Xuất toán tất cả: giữ mọi lượt dịch vụ có trùng thời gian.",
            wraplength=565, justify="left",
        ).pack(padx=16, pady=(0, 12))

        def chon_che_do(che_do=None):
            popup.grab_release()
            popup.destroy()
            if che_do is None:
                self.btn_chuan_doan.configure(state="normal")
            else:
                self._bat_dau_chuan_doan(che_do)

        nut = ctk.CTkFrame(frame, fg_color="transparent")
        nut.pack()
        ctk.CTkButton(
            nut, text="Loại trừ dịch vụ trùng", width=220,
            command=lambda: chon_che_do(CHE_DO_LOAI_TRU_TRUNG),
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            nut, text="Xuất toán tất cả", width=180,
            command=lambda: chon_che_do(CHE_DO_XUAT_TOAN_TAT_CA),
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            frame, text="Hủy", width=100, fg_color="#64748b", command=chon_che_do,
        ).pack(pady=(12, 0))
        popup.protocol("WM_DELETE_WINDOW", chon_che_do)
        popup.grab_set()
        popup.focus_set()

    def _bat_dau_chuan_doan(self, che_do):
        path_file = self.entry_file.get().strip()
        phut = {khoa: entry.get().strip() for khoa, (_, entry, _, _) in self.dong_dich_vu.items()
                if entry.get().strip()}
        chon = {khoa for khoa, (var, _, _, _) in self.dong_dich_vu.items() if var.get()}
        self.che_do_chuan_doan = che_do
        self.btn_xuat_file.configure(state="disabled")
        self.btn_xoa_tat_ca_loc.configure(state="disabled")
        self.update_status("Đang kiểm tra trùng thời gian...", "blue")
        df_nguon, danh_muc = self.df_nguon, self.danh_muc
        def worker():
            try:
                result = chuan_doan_xml3_trung_thoi_gian_yhct(
                    path_file, phut_theo_dich_vu=phut, dich_vu_chon=chon,
                    df_nguon=df_nguon, danh_muc=danh_muc, che_do=che_do,
                )
                self._job_queue.put(("diagnose", result))
            except Exception as exc:
                self._job_queue.put(("error", "diagnose", exc))
        threading.Thread(target=worker, daemon=True).start()
        self.after(100, self._poll_job)

    def xuat_file(self):
        if self.df_hien_thi is None or self.df_hien_thi.empty:
            messagebox.showwarning("Không có dữ liệu", "Không có dữ liệu để xuất.")
            return
        file_xuat = filedialog.asksaveasfilename(
            initialfile=(
                "Chuan_Doan_XML3_Trung_Thoi_Gian_YHCT_Loai_Tru_Trung.xlsx"
                if self.che_do_chuan_doan == CHE_DO_LOAI_TRU_TRUNG
                else "Chuan_Doan_XML3_Trung_Thoi_Gian_YHCT_Xuat_Toan_Tat_Ca.xlsx"
            ),
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not file_xuat:
            return
        self.btn_xuat_file.configure(state="disabled")
        self.update_status("Đang ghi file Excel...", "blue")
        data = self.df_hien_thi.copy()
        def worker():
            try:
                output = xuat_excel_chuan_doan_xml3_trung_thoi_gian_yhct(data, file_xuat)
                self._job_queue.put(("export", output))
            except Exception as exc:
                self._job_queue.put(("error", "export", exc))
        threading.Thread(target=worker, daemon=True).start()
        self.after(100, self._poll_job)

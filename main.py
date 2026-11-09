import tkinter as tk
from pathlib import Path

import customtkinter as ctk

from modules.module_7980.ui_7980 import UI_7980
from modules.module_01_bh_02_bh.ui_01_bh_02_bh import UI_01BH02BH
from modules.module_01_bh_his_giam_dinh.ui_01_bh_his_giam_dinh import UI_01BHHISGiamDinh
from modules.module_giam_dinh_xml_mau01.ui_giam_dinh_xml_mau01 import UI_GiamDinhXMLMau01
from modules.module_xml01_bh_mau_chuan.ui_xml01_bh_mau_chuan import UI_XML01BHMauChuan
from modules.module_giam_dinh_01_bh_mau_chuan.ui_giam_dinh_01_bh_mau_chuan import UI_GiamDinh01BHMauChuan
from modules.module_xml3.ui_xml3 import UI_XML3
from modules.module_xml1_xml3.ui_xml1_xml3 import UI_XML1_XML3
from modules.module_xu_ly_ma_may.ui_xu_ly_ma_may import UI_XuLyMaMay
from modules.module_chuan_doan_xml3_kham_t7_cn.ui_chuan_doan_xml3_kham_t7_cn import UI_ChuanDoanXML3KhamT7CN
from modules.module_chuan_doan_xml3_phong_xet_nghiem.ui_chuan_doan_xml3_phong_xet_nghiem import UI_ChuanDoanXML3PhongXetNghiem
from modules.module_chuan_doan_xml3_pttt_tuoi.ui_chuan_doan_xml3_pttt_tuoi import UI_ChuanDoanXML3PTTTTuoi
from modules.module_chuan_doan_xml3_pttt_ngay_kq.ui_chuan_doan_xml3_pttt_ngay_kq import UI_ChuanDoanXML3PTTTNgayKQ
from modules.module_chuan_doan_xml3_trung_dvkt.ui_chuan_doan_xml3_trung_dvkt import UI_ChuanDoanXML3TrungDVKT
from modules.module_chuan_doan_xml3_so_luot_yhct.ui_chuan_doan_xml3_so_luot_yhct import UI_ChuanDoanXML3SoLuotYHCT
from modules.module_chuan_doan_xml3_trung_thoi_gian_yhct.ui_chuan_doan_xml3_trung_thoi_gian_yhct import UI_ChuanDoanXML3TrungThoiGianYHCT
from modules.module_chuan_doan_xml3_ngay_giuong_noi_tru.ui_chuan_doan_xml3_ngay_giuong_noi_tru import UI_ChuanDoanXML3NgayGiuongNoiTru


ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


def doc_phien_ban():
    version_file = Path(__file__).resolve().parent / "VERSION"
    try:
        version = version_file.read_text(encoding="utf-8").strip()
        return version or "0.0.0"
    except OSError:
        return "0.0.0"


class AppChinh(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.phien_ban = doc_phien_ban()
        self.title(f"Phần mềm Quản lý BHYT - v{self.phien_ban}")
        self.geometry("1300x780")
        self.minsize(1100, 680)

        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._tao_thanh_cong_cu()
        self._tao_vung_noi_dung()
        self.chon_menu("7980")

    def _tao_thanh_cong_cu(self):
        toolbar = ctk.CTkFrame(self, height=56, corner_radius=0, fg_color=("#e2e8f0", "#1e293b"))
        toolbar.grid(row=0, column=0, sticky="ew")
        toolbar.grid_propagate(False)
        toolbar.grid_columnconfigure(10, weight=1)

        lbl_brand = ctk.CTkLabel(
            toolbar,
            text="BHYT TOOLBAR",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=("#0f172a", "#f8fafc"),
        )
        lbl_brand.grid(row=0, column=0, padx=(16, 24), pady=10)

        menu_button = tk.Menubutton(
            toolbar,
            text="Đối chiếu",
            relief="flat",
            bg="#cbd5e1",
            activebackground="#94a3b8",
            font=("Segoe UI", 11, "bold"),
            padx=16,
            pady=8,
        )
        menu_button.grid(row=0, column=1, padx=(0, 10), pady=8, sticky="w")

        doi_chieu_menu = tk.Menu(menu_button, tearoff=0)
        doi_chieu_menu.add_command(label="7980", command=lambda: self.chon_menu("7980"))
        doi_chieu_menu.add_command(label="01-bh giám định với 02-bh HIS", command=lambda: self.chon_menu("01_bh_02_bh"))
        doi_chieu_menu.add_command(label="01-bh HIS với 01-bh giám định", command=lambda: self.chon_menu("01_bh_his_giam_dinh"))
        doi_chieu_menu.add_command(label="Giám định xml với Mau 01", command=lambda: self.chon_menu("giam_dinh_xml_mau01"))
        doi_chieu_menu.add_command(label="Xml 01-bh với mẫu chuẩn", command=lambda: self.chon_menu("xml01_bh_mau_chuan"))
        doi_chieu_menu.add_command(label="Giám định 01-bh với mẫu chuẩn", command=lambda: self.chon_menu("giam_dinh_01_bh_mau_chuan"))
        menu_button.configure(menu=doi_chieu_menu)

        def hien_menu_hover(_event):
            x = menu_button.winfo_rootx()
            y = menu_button.winfo_rooty() + menu_button.winfo_height()
            doi_chieu_menu.tk_popup(x, y)

        menu_button.bind("<Enter>", hien_menu_hover)

        xu_ly_button = tk.Menubutton(
            toolbar,
            text="Xử lý",
            relief="flat",
            bg="#0ea5e9",
            activebackground="#38bdf8",
            fg="#ffffff",
            font=("Segoe UI", 11, "bold"),
            padx=16,
            pady=8,
        )
        xu_ly_button.grid(row=0, column=2, padx=(0, 10), pady=8, sticky="w")

        xu_ly_menu = tk.Menu(xu_ly_button, tearoff=0)
        xu_ly_menu.add_command(label="XML3", command=lambda: self.chon_menu("xml3"))
        xu_ly_menu.add_command(label="Xml1+Xml3", command=lambda: self.chon_menu("xml1_xml3"))
        xu_ly_menu.add_command(label="Mã máy", command=lambda: self.chon_menu("xu_ly_ma_may"))
        xu_ly_button.configure(menu=xu_ly_menu)

        def hien_menu_xu_ly_hover(_event):
            x = xu_ly_button.winfo_rootx()
            y = xu_ly_button.winfo_rooty() + xu_ly_button.winfo_height()
            xu_ly_menu.tk_popup(x, y)

        xu_ly_button.bind("<Enter>", hien_menu_xu_ly_hover)

        chuan_doan_button = tk.Menubutton(
            toolbar,
            text="Chuẩn đoán",
            relief="flat",
            bg="#8b5cf6",
            activebackground="#a78bfa",
            fg="#ffffff",
            font=("Segoe UI", 11, "bold"),
            padx=16,
            pady=8,
        )
        chuan_doan_button.grid(row=0, column=3, padx=(0, 10), pady=8, sticky="w")

        chuan_doan_menu = tk.Menu(chuan_doan_button, tearoff=0)
        chuan_doan_menu.add_command(
            label="Xml3-Phòng xét nghiệm",
            command=lambda: self.chon_menu("chuan_doan_xml3_phong_xet_nghiem"),
        )
        chuan_doan_menu.add_command(
            label="Xml3-Khám t7,cn",
            command=lambda: self.chon_menu("chuan_doan_xml3_kham_t7_cn"),
        )
        chuan_doan_menu.add_command(
            label="Xml3-PTTT dành cho người trên và dưới 16 tuổi",
            command=lambda: self.chon_menu("chuan_doan_xml3_pttt_tuoi"),
        )
        chuan_doan_menu.add_command(
            label="Xml3-ngày trả kết quả có sau ngày ra viện",
            command=lambda: self.chon_menu("chuan_doan_xml3_pttt_ngay_kq"),
        )
        chuan_doan_menu.add_command(
            label="Xml3-trùng dịch vụ kỹ thuật",
            command=lambda: self.chon_menu("chuan_doan_xml3_trung_dvkt"),
        )
        chuan_doan_menu.add_command(
            label="Xml3-Số lượt thực hiện DVKT của Bs.YHCT",
            command=lambda: self.chon_menu("chuan_doan_xml3_so_luot_yhct"),
        )
        chuan_doan_menu.add_command(
            label="Xml3-Trùng thời gian DVKT của Bs.YHCT",
            command=lambda: self.chon_menu("chuan_doan_xml3_trung_thoi_gian_yhct"),
        )
        chuan_doan_menu.add_command(
            label="Xml3-Ngày giường nội trú",
            command=lambda: self.chon_menu("chuan_doan_xml3_ngay_giuong_noi_tru"),
        )
        chuan_doan_button.configure(menu=chuan_doan_menu)

        def hien_menu_chuan_doan_hover(_event):
            x = chuan_doan_button.winfo_rootx()
            y = chuan_doan_button.winfo_rooty() + chuan_doan_button.winfo_height()
            chuan_doan_menu.tk_popup(x, y)

        chuan_doan_button.bind("<Enter>", hien_menu_chuan_doan_hover)


        self.lbl_trang_thai = ctk.CTkLabel(toolbar, text=f"Sẵn sàng | v{self.phien_ban}", text_color=("#334155", "#cbd5e1"))
        self.lbl_trang_thai.grid(row=0, column=11, padx=16, pady=10, sticky="e")

    def _tao_vung_noi_dung(self):
        self.content_frame = ctk.CTkFrame(self, fg_color=("#f8fafc", "#0f172a"))
        self.content_frame.grid(row=1, column=0, sticky="nsew", padx=12, pady=12)
        self.content_frame.grid_rowconfigure(0, weight=1)
        self.content_frame.grid_columnconfigure(0, weight=1)

        self.frame_7980 = UI_7980(self.content_frame, corner_radius=12)
        self.frame_01_bh_02_bh = UI_01BH02BH(self.content_frame, corner_radius=12)
        self.frame_01_bh_his_giam_dinh = UI_01BHHISGiamDinh(self.content_frame, corner_radius=12)
        self.frame_giam_dinh_xml_mau01 = UI_GiamDinhXMLMau01(self.content_frame, corner_radius=12)
        self.frame_xml01_bh_mau_chuan = UI_XML01BHMauChuan(self.content_frame, corner_radius=12)
        self.frame_giam_dinh_01_bh_mau_chuan = UI_GiamDinh01BHMauChuan(self.content_frame, corner_radius=12)
        self.frame_xml3 = UI_XML3(self.content_frame, corner_radius=12)
        self.frame_xu_ly_ma_may = UI_XuLyMaMay(self.content_frame, corner_radius=12)
        self.frame_xml1_xml3 = UI_XML1_XML3(self.content_frame, corner_radius=12)
        self.frame_chuan_doan_xml3_phong_xet_nghiem = UI_ChuanDoanXML3PhongXetNghiem(self.content_frame, corner_radius=12)
        self.frame_chuan_doan_xml3_kham_t7_cn = UI_ChuanDoanXML3KhamT7CN(self.content_frame, corner_radius=12)
        self.frame_chuan_doan_xml3_pttt_tuoi = UI_ChuanDoanXML3PTTTTuoi(self.content_frame, corner_radius=12)
        self.frame_chuan_doan_xml3_pttt_ngay_kq = UI_ChuanDoanXML3PTTTNgayKQ(self.content_frame, corner_radius=12)
        self.frame_chuan_doan_xml3_trung_dvkt = UI_ChuanDoanXML3TrungDVKT(self.content_frame, corner_radius=12)
        self.frame_chuan_doan_xml3_trung_thoi_gian_yhct = UI_ChuanDoanXML3TrungThoiGianYHCT(self.content_frame, corner_radius=12)
        self.frame_chuan_doan_xml3_so_luot_yhct = UI_ChuanDoanXML3SoLuotYHCT(self.content_frame, corner_radius=12)
        self.frame_chuan_doan_xml3_ngay_giuong_noi_tru = UI_ChuanDoanXML3NgayGiuongNoiTru(self.content_frame, corner_radius=12)

        self.menu_map = {
            "7980": self.frame_7980,
            "01_bh_02_bh": self.frame_01_bh_02_bh,
            "01_bh_his_giam_dinh": self.frame_01_bh_his_giam_dinh,
            "giam_dinh_xml_mau01": self.frame_giam_dinh_xml_mau01,
            "xml01_bh_mau_chuan": self.frame_xml01_bh_mau_chuan,
            "giam_dinh_01_bh_mau_chuan": self.frame_giam_dinh_01_bh_mau_chuan,
            "xml3": self.frame_xml3,
            "xu_ly_ma_may": self.frame_xu_ly_ma_may,
            "xml1_xml3": self.frame_xml1_xml3,
            "chuan_doan_xml3_phong_xet_nghiem": self.frame_chuan_doan_xml3_phong_xet_nghiem,
            "chuan_doan_xml3_kham_t7_cn": self.frame_chuan_doan_xml3_kham_t7_cn,
            "chuan_doan_xml3_pttt_tuoi": self.frame_chuan_doan_xml3_pttt_tuoi,
            "chuan_doan_xml3_pttt_ngay_kq": self.frame_chuan_doan_xml3_pttt_ngay_kq,
            "chuan_doan_xml3_trung_dvkt": self.frame_chuan_doan_xml3_trung_dvkt,
            "chuan_doan_xml3_trung_thoi_gian_yhct": self.frame_chuan_doan_xml3_trung_thoi_gian_yhct,
            "chuan_doan_xml3_so_luot_yhct": self.frame_chuan_doan_xml3_so_luot_yhct,
            "chuan_doan_xml3_ngay_giuong_noi_tru": self.frame_chuan_doan_xml3_ngay_giuong_noi_tru,
        }

    def chon_menu(self, menu_name):
        for frame in self.menu_map.values():
            frame.grid_forget()
        self.menu_map[menu_name].grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        self.lbl_trang_thai.configure(text=f"Đang chọn: {menu_name} | v{self.phien_ban}")


if __name__ == "__main__":
    app = AppChinh()
    app.mainloop()


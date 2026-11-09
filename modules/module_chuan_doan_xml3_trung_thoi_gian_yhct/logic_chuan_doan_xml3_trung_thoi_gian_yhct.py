"""Dò các khoảng thời gian DVKT giao nhau theo người thực hiện."""

from bisect import bisect_left
import os

import pandas as pd
from openpyxl.styles import PatternFill

from core.utils import doc_file_thong_minh
from modules.module_chuan_doan_xml3_so_luot_yhct.logic_chuan_doan_xml3_so_luot_yhct import (
    COT_BAT_BUOC,
    COT_TIEN,
    _chuan_hoa_cot,
    _chuan_hoa_gia_tri,
    _chuan_hoa_ma_nhom,
    _chuan_hoa_text,
    _dinh_dang_ngay_hien_thi,
    _doc_phut_dich_vu,
    _tim_dich_vu_cau_hinh,
    xuat_excel_chuan_doan_xml3_so_luot_yhct,
)


LOI_TRUNG_THOI_GIAN = "Trùng thời gian thực hiện DVKT"
CHE_DO_LOAI_TRU_TRUNG = "loai_tru_trung"
CHE_DO_XUAT_TOAN_TAT_CA = "xuat_toan_tat_ca"
COT_KET_QUA = [
    "MA_LK", "STT", "MA_DICH_VU", "TEN_DICH_VU", "NGUOI_THUC_HIEN",
    "MA_BAC_SI", "MA_BA", "NGAY_TH_YL", "PHUT_THUC_HIEN",
    "KET_THUC_DU_KIEN", "LOAI_TRUNG", "TRUNG_VOI_TRUOC",
    "TRUNG_CUNG_GIO", "TRUNG_VOI_SAU", "SO_DVKT_TRUNG",
    *COT_TIEN, "NGAY_KQ", "LOI",
]
COT_CAN_DOC = list(dict.fromkeys([*COT_BAT_BUOC, "MA_LK", "STT", "MA_DICH_VU"]))


def _lay_thoi_gian(val):
    raw = _chuan_hoa_gia_tri(val)
    digits = "".join(ch for ch in raw if ch.isdigit())
    if len(digits) < 12:
        return pd.NaT
    return pd.to_datetime(digits[:12], format="%Y%m%d%H%M", errors="coerce")


def _dinh_dang_thoi_gian(val):
    return "" if pd.isna(val) else val.strftime("%d/%m/%Y %H:%M")


def _khoa_dich_vu(ma, ten):
    return (_chuan_hoa_text(ma), _chuan_hoa_text(ten))


def phan_tich_file_xml3(file_xml3, callback_status=None):
    """Đọc một lần và trả về dữ liệu cùng danh mục mã–tên MA_NHOM 18."""
    if callback_status:
        callback_status("Đang phân tích file XML3...", "blue")
    df = doc_file_thong_minh(file_xml3)
    if df is None or df.empty:
        raise ValueError("File XML3 không có dữ liệu.")
    df = _chuan_hoa_cot(df)
    thieu = [cot for cot in COT_CAN_DOC if cot not in df.columns]
    if thieu:
        raise ValueError(f"Thiếu cột bắt buộc: {', '.join(thieu)}")

    ma_nhom = df["MA_NHOM"].map(_chuan_hoa_ma_nhom)
    nhom_18 = df.loc[ma_nhom.eq("18"), ["MA_DICH_VU", "TEN_DICH_VU"]].copy()
    if nhom_18.empty:
        raise ValueError("Không tìm thấy dịch vụ MA_NHOM = 18.")
    nhom_18 = nhom_18.fillna("")
    for cot in nhom_18.columns:
        nhom_18[cot] = nhom_18[cot].map(_chuan_hoa_gia_tri)
    cau_hinh_mac_dinh = _doc_phut_dich_vu(None)
    danh_muc = []
    for (ma, ten), so_luot in nhom_18.value_counts(sort=False).items():
        nhom = _tim_dich_vu_cau_hinh(ten, cau_hinh_mac_dinh)
        danh_muc.append({
            "ma": ma, "ten": ten, "so_luot": int(so_luot),
            "phut_mac_dinh": cau_hinh_mac_dinh[nhom]["phut"] if nhom else None,
        })
    danh_muc.sort(key=lambda item: (item["phut_mac_dinh"] is None, item["ten"].casefold(), item["ma"]))
    if callback_status:
        callback_status(f"Đã phân tích: {len(danh_muc)} mã–tên dịch vụ nhóm 18.", "green")
    return df, danh_muc


def _kiem_tra_cau_hinh(danh_muc, phut_theo_dich_vu, dich_vu_chon):
    phut_theo_dich_vu = phut_theo_dich_vu or {}
    tat_ca = {}
    chon = set(dich_vu_chon) if dich_vu_chon is not None else None
    for item in danh_muc:
        khoa = _khoa_dich_vu(item["ma"], item["ten"])
        if chon is not None and khoa in chon and khoa not in phut_theo_dich_vu:
            raise ValueError(f"Chưa nhập số phút cho dịch vụ được chọn: {item['ten']} ({item['ma']}).")
        raw = phut_theo_dich_vu.get(khoa, item["phut_mac_dinh"])
        if raw is None or str(raw).strip() == "":
            if chon is not None and khoa in chon:
                raise ValueError(f"Chưa nhập số phút cho dịch vụ được chọn: {item['ten']} ({item['ma']}).")
            continue
        try:
            phut = int(str(raw).strip())
        except (ValueError, TypeError):
            raise ValueError(f"Số phút không hợp lệ: {item['ten']} ({item['ma']}).") from None
        if phut <= 0:
            raise ValueError(f"Số phút phải lớn hơn 0: {item['ten']} ({item['ma']}).")
        tat_ca[khoa] = phut
    if chon is None:
        chon = set(tat_ca)
    if not chon:
        raise ValueError("Hãy chọn ít nhất một dịch vụ để chuẩn đoán.")
    if not chon.issubset(tat_ca):
        raise ValueError("Dịch vụ được chọn phải có thời lượng hợp lệ.")
    return tat_ca, chon


def chuan_doan_xml3_trung_thoi_gian_yhct(
    file_xml3, phut_theo_dich_vu=None, callback_status=None,
    dich_vu_chon=None, df_nguon=None, danh_muc=None,
    che_do=CHE_DO_XUAT_TOAN_TAT_CA,
):
    if che_do not in {CHE_DO_LOAI_TRU_TRUNG, CHE_DO_XUAT_TOAN_TAT_CA}:
        raise ValueError(f"Chế độ chuẩn đoán không hợp lệ: {che_do}")
    if df_nguon is None or danh_muc is None:
        df_nguon, danh_muc = phan_tich_file_xml3(file_xml3, callback_status)
    tat_ca_phut, chon = _kiem_tra_cau_hinh(danh_muc, phut_theo_dich_vu, dich_vu_chon)
    if callback_status:
        callback_status("Đang kiểm tra dịch vụ trước, cùng giờ và sau...", "blue")

    df = df_nguon[COT_CAN_DOC].copy()
    for cot in COT_CAN_DOC:
        df[cot] = df[cot].map(_chuan_hoa_gia_tri)
    df["_ID"] = range(len(df))
    df["_KHOA"] = [_khoa_dich_vu(ma, ten) for ma, ten in zip(df["MA_DICH_VU"], df["TEN_DICH_VU"])]
    df["_PHUT"] = df["_KHOA"].map(tat_ca_phut)
    df["_NGUOI"] = df["NGUOI_THUC_HIEN"].map(_chuan_hoa_text)
    df["_BAT_DAU"] = df["NGAY_TH_YL"].map(_lay_thoi_gian)
    df["_KET_THUC"] = df["_BAT_DAU"] + pd.to_timedelta(df["_PHUT"], unit="m")
    nhom_18 = df["MA_NHOM"].map(_chuan_hoa_ma_nhom).eq("18")
    mask_xet = nhom_18 & df["_KHOA"].isin(chon) & df["_NGUOI"].ne("") & df["_BAT_DAU"].notna()
    ket_qua = df.loc[mask_xet].copy()
    for cot in ["TRUNG_VOI_TRUOC", "TRUNG_CUNG_GIO", "TRUNG_VOI_SAU", "LOAI_TRUNG", "LOI"]:
        ket_qua[cot] = ""
    ket_qua["SO_DVKT_TRUNG"] = 0
    ket_qua["PHUT_THUC_HIEN"] = ket_qua["_PHUT"].astype(int)
    ket_qua["KET_THUC_DU_KIEN"] = ket_qua["_KET_THUC"].map(_dinh_dang_thoi_gian)
    def mo_ta(dong):
        return f"{dong['TEN_DICH_VU']} [{dong['MA_DICH_VU']}] ({_dinh_dang_thoi_gian(dong['_BAT_DAU'])})"

    for _, nhom in df.loc[df["_NGUOI"].ne("") & df["_BAT_DAU"].notna()].groupby("_NGUOI", sort=False):
        nhom = nhom.sort_values(["_BAT_DAU", "_ID"], kind="stable")
        thoi_diem = list(nhom["_BAT_DAU"])
        chi_so = list(nhom.index)
        # Chỉ DVKT có thời lượng đã biết mới được khẳng định là còn đang làm trước lúc bắt đầu.
        phut_max = nhom["_PHUT"].max()
        if pd.isna(phut_max):
            phut_max = 0
        for idx, dong in nhom.loc[nhom.index.isin(ket_qua.index)].iterrows():
            bat_dau = dong["_BAT_DAU"]
            ket_thuc = dong["_KET_THUC"]
            truoc, cung_gio, sau = [], [], []
            bat_dau_tim = bisect_left(thoi_diem, bat_dau - pd.Timedelta(minutes=phut_max))
            ket_thuc_tim = bisect_left(thoi_diem, ket_thuc)
            for vi_tri in range(bat_dau_tim, ket_thuc_tim):
                khac = nhom.loc[chi_so[vi_tri]]
                if khac["_ID"] == dong["_ID"]:
                    continue
                if khac["_BAT_DAU"] < bat_dau:
                    if pd.notna(khac["_KET_THUC"]) and khac["_KET_THUC"] > bat_dau:
                        truoc.append(mo_ta(khac))
                elif khac["_BAT_DAU"] == bat_dau:
                    cung_gio.append(mo_ta(khac))
                else:
                    sau.append(mo_ta(khac))
            ket_qua.at[idx, "TRUNG_VOI_TRUOC"] = "; ".join(truoc)
            ket_qua.at[idx, "TRUNG_CUNG_GIO"] = "; ".join(cung_gio)
            ket_qua.at[idx, "TRUNG_VOI_SAU"] = "; ".join(sau)
            loai = [ten for ten, ds in (("Trước", truoc), ("Cùng giờ", cung_gio), ("Sau", sau)) if ds]
            ket_qua.at[idx, "LOAI_TRUNG"] = ", ".join(loai)
            ket_qua.at[idx, "SO_DVKT_TRUNG"] = len(truoc) + len(cung_gio) + len(sau)
            if loai:
                ket_qua.at[idx, "LOI"] = f"{LOI_TRUNG_THOI_GIAN}: {', '.join(loai)}"

    # Mỗi dịch vụ được chọn có giao thời gian đều được xét, kể cả hai phía của một cặp.
    ket_qua = ket_qua.loc[ket_qua["SO_DVKT_TRUNG"].gt(0)].copy()
    if che_do == CHE_DO_LOAI_TRU_TRUNG:
        chi_so_giu = []
        for _, nhom_dich_vu in ket_qua.groupby(["_NGUOI", "_KHOA"], sort=False):
            nhom_dich_vu = nhom_dich_vu.sort_values(["_BAT_DAU", "_ID"], kind="stable")
            ket_thuc_da_giu = pd.NaT
            for idx, dong in nhom_dich_vu.iterrows():
                if pd.isna(ket_thuc_da_giu) or dong["_BAT_DAU"] >= ket_thuc_da_giu:
                    chi_so_giu.append(idx)
                    ket_thuc_da_giu = dong["_KET_THUC"]
        ket_qua = ket_qua.loc[chi_so_giu].copy()

    for cot in COT_TIEN:
        ket_qua[cot] = pd.to_numeric(ket_qua[cot], errors="coerce")
    for cot in ("NGAY_KQ", "NGAY_TH_YL"):
        ket_qua[cot] = ket_qua[cot].map(_dinh_dang_ngay_hien_thi)
    ket_qua = ket_qua.sort_values(
        ["_NGUOI", "_BAT_DAU", "_ID"], kind="stable",
    ).loc[:, COT_KET_QUA].reset_index(drop=True)
    if callback_status:
        callback_status(
            f"Chuẩn đoán xong: {len(ket_qua)} dòng, {int(ket_qua['SO_DVKT_TRUNG'].gt(0).sum())} dòng trùng.",
            "green",
        )
    return ket_qua


def xuat_excel_chuan_doan_xml3_trung_thoi_gian_yhct(df_ket_qua, file_xuat):
    if df_ket_qua is None or df_ket_qua.empty:
        raise ValueError("Không có dữ liệu để xuất file.")
    base, extension = os.path.splitext(file_xuat)
    counter = 1
    while os.path.exists(file_xuat):
        file_xuat = f"{base} ({counter}){extension}"
        counter += 1
    tong_hop = df_ket_qua.copy()
    tong_hop.insert(0, "NGAY", tong_hop["NGAY_TH_YL"].astype(str).str.split().str[0])
    tong_hop = tong_hop.groupby(
        ["NGUOI_THUC_HIEN", "MA_BAC_SI", "NGAY", "MA_DICH_VU", "TEN_DICH_VU"],
        as_index=False, sort=False, dropna=False,
    ).agg(
        SO_LUOT_TRUNG=("MA_LK", "size"),
        THANH_TIEN_BH=("THANH_TIEN_BH", "sum"),
        T_BHTT=("T_BHTT", "sum"),
    )
    with pd.ExcelWriter(file_xuat, engine="openpyxl") as writer:
        df_ket_qua.to_excel(writer, sheet_name="TrungThoiGian_YHCT", index=False)
        tong_hop.to_excel(writer, sheet_name="TongHopTheoNgay", index=False)
        ws = writer.sheets["TrungThoiGian_YHCT"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        mau_loi = PatternFill(fill_type="solid", fgColor="FFF59D")
        headers = [cell.value for cell in ws[1]]
        cot_so_trung = headers.index("SO_DVKT_TRUNG") if "SO_DVKT_TRUNG" in headers else None
        for row in ws.iter_rows(min_row=2):
            if cot_so_trung is not None and (row[cot_so_trung].value or 0) > 0:
                for cell in row:
                    cell.fill = mau_loi
        for col in ws.columns:
            ten = str(col[0].value)
            if ten in {"TEN_DICH_VU", "TRUNG_VOI_TRUOC", "TRUNG_CUNG_GIO", "TRUNG_VOI_SAU"}:
                width = 55
            elif ten in {"NGAY_TH_YL", "KET_THUC_DU_KIEN", "NGAY_KQ"}:
                width = 20
            else:
                width = 22
            ws.column_dimensions[col[0].column_letter].width = width
        ws_tong_hop = writer.sheets["TongHopTheoNgay"]
        ws_tong_hop.freeze_panes = "A2"
        ws_tong_hop.auto_filter.ref = ws_tong_hop.dimensions
        for col in ws_tong_hop.columns:
            ws_tong_hop.column_dimensions[col[0].column_letter].width = (
                55 if col[0].value == "TEN_DICH_VU" else 22
            )
    return file_xuat

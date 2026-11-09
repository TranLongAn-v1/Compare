import os

import pandas as pd

from core.utils import doc_file_thong_minh


COT_LOC_TRUNG = ["MA_BA", "HO_TEN", "MA_THE"]
COT_LOC_TRUNG_THEO_NGAY = COT_LOC_TRUNG + ["_NGAY_TH_YL_NGAY"]
COT_HIEN_THI = ["MA_BA", "HO_TEN", "SO_PHIEU", "NGAY_TH_YL", "TEN_DICH_VU"]
COT_SO_PHIEU_THAY_THE = ["SO_PHIEU", "SOPHIEU"]


def _chuan_hoa_cot(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    if "SO_PHIEU" not in df.columns and "SOPHIEU" in df.columns:
        df["SO_PHIEU"] = df["SOPHIEU"]
    return df


def _chuan_hoa_ma_nhom(val):
    if pd.isna(val):
        return None
    raw = str(val).strip()
    if not raw:
        return None
    raw = raw.replace(",", ".")
    try:
        return int(float(raw))
    except Exception:
        return None


def _chuan_hoa_text_loc(val):
    return str(val).casefold().replace(" - ", "-").replace(" -", "-").replace("- ", "-")


def _lay_ngay_th_yl(val):
    if pd.isna(val):
        return ""
    raw = str(val).strip()
    if not raw:
        return ""
    if raw.endswith(".0"):
        raw = raw[:-2]
    digits = "".join(ch for ch in raw if ch.isdigit())
    if len(digits) < 8:
        return ""
    return digits[:8]


def _chuan_hoa_ds_ma_nhom(ma_nhom_can_lay):
    if isinstance(ma_nhom_can_lay, (list, tuple, set)):
        return {_chuan_hoa_ma_nhom(v) for v in ma_nhom_can_lay}
    return {_chuan_hoa_ma_nhom(ma_nhom_can_lay)}


def _tao_mask_ten_dich_vu(ten_dich_vu, cum_tu_loc):
    ds_cum_tu = cum_tu_loc if isinstance(cum_tu_loc, (list, tuple, set)) else [cum_tu_loc]
    mask = False
    for cum_tu in ds_cum_tu:
        mask = mask | ten_dich_vu.str.contains(_chuan_hoa_text_loc(cum_tu), regex=False)
    return mask


def tong_hop_benh_nhan_xml3(
    path_file,
    ma_nhom_can_lay=13,
    ten_dich_vu_chua=None,
    ten_dich_vu_chua_theo_ma_nhom=None,
    callback_status=None,
):
    if callback_status:
        callback_status("Đang đọc file dữ liệu...", "blue")

    df = doc_file_thong_minh(path_file)
    if df is None or df.empty:
        raise ValueError("File không có dữ liệu.")

    df = _chuan_hoa_cot(df)
    thieu_cot = [c for c in COT_LOC_TRUNG + ["NGAY_TH_YL", "MA_NHOM", "TEN_DICH_VU"] if c not in df.columns]
    if not any(c in df.columns for c in COT_SO_PHIEU_THAY_THE):
        thieu_cot.append("SO_PHIEU/SOPHIEU")
    if thieu_cot:
        raise ValueError(f"Thiếu cột bắt buộc: {', '.join(thieu_cot)}")

    if callback_status:
        callback_status(f"Đang lọc dữ liệu theo MA_NHOM = {ma_nhom_can_lay}...", "blue")

    ma_nhom_series = df["MA_NHOM"].apply(_chuan_hoa_ma_nhom)
    ds_ma_nhom_can_lay = _chuan_hoa_ds_ma_nhom(ma_nhom_can_lay)
    df_loc = df.loc[ma_nhom_series.isin(ds_ma_nhom_can_lay)].copy()
    ma_nhom_loc = ma_nhom_series.loc[df_loc.index]
    ten_dich_vu = df_loc["TEN_DICH_VU"].fillna("").astype(str).apply(_chuan_hoa_text_loc)

    if ten_dich_vu_chua_theo_ma_nhom:
        if callback_status:
            callback_status("Đang lọc TEN_DICH_VU theo từng MA_NHOM...", "blue")
        mask_theo_nhom = pd.Series(True, index=df_loc.index)
        for ma_nhom, cum_tu_loc in ten_dich_vu_chua_theo_ma_nhom.items():
            ma_nhom_chuan = _chuan_hoa_ma_nhom(ma_nhom)
            mask_nhom = ma_nhom_loc == ma_nhom_chuan
            if mask_nhom.any():
                mask_theo_nhom.loc[mask_nhom] = _tao_mask_ten_dich_vu(ten_dich_vu.loc[mask_nhom], cum_tu_loc)
        df_loc = df_loc.loc[mask_theo_nhom].copy()
        ten_dich_vu = df_loc["TEN_DICH_VU"].fillna("").astype(str).apply(_chuan_hoa_text_loc)

    if ten_dich_vu_chua:
        cum_tu_loc = ten_dich_vu_chua if isinstance(ten_dich_vu_chua, (list, tuple, set)) else [ten_dich_vu_chua]
        if callback_status:
            callback_status(f"Đang lọc TEN_DICH_VU có cụm từ {', '.join(cum_tu_loc)}...", "blue")
        df_loc = df_loc.loc[_tao_mask_ten_dich_vu(ten_dich_vu, cum_tu_loc)].copy()

    if callback_status:
        callback_status("Đang tổng hợp dữ liệu bệnh nhân tới khám...", "blue")

    df_ket_qua = df_loc[COT_LOC_TRUNG + COT_HIEN_THI].copy()
    df_ket_qua = df_ket_qua.loc[:, ~df_ket_qua.columns.duplicated()]
    for cot in df_ket_qua.columns:
        df_ket_qua[cot] = df_ket_qua[cot].fillna("").astype(str).str.strip()
    df_ket_qua["_NGAY_TH_YL_NGAY"] = df_ket_qua["NGAY_TH_YL"].apply(_lay_ngay_th_yl)

    if callback_status:
        callback_status(f"Tổng hợp xong: {len(df_ket_qua)} dòng.", "green")

    return df_ket_qua.reset_index(drop=True)


def loc_trung_benh_nhan(df_nguon):
    if df_nguon is None or df_nguon.empty:
        raise ValueError("Không có dữ liệu để lọc.")

    df = df_nguon.copy()
    thieu_cot = [c for c in COT_LOC_TRUNG + ["NGAY_TH_YL"] if c not in df.columns]
    if thieu_cot:
        raise ValueError(f"Thiếu cột bắt buộc: {', '.join(thieu_cot)}")

    for cot in COT_LOC_TRUNG + ["NGAY_TH_YL"]:
        df[cot] = df[cot].fillna("").astype(str).str.strip()
    df["_NGAY_TH_YL_NGAY"] = df["NGAY_TH_YL"].apply(_lay_ngay_th_yl)

    df_loc = df.drop_duplicates(subset=COT_LOC_TRUNG_THEO_NGAY, keep="first").reset_index(drop=True)
    return df_loc


def xuat_excel_tong_hop_benh_nhan(df_ket_qua, file_xuat):
    if df_ket_qua is None or df_ket_qua.empty:
        raise ValueError("Không có dữ liệu để xuất file.")

    df_ket_qua = df_ket_qua.drop(columns=["_NGAY_TH_YL_NGAY"], errors="ignore")

    base, extension = os.path.splitext(file_xuat)
    counter = 1
    while os.path.exists(file_xuat):
        file_xuat = f"{base} ({counter}){extension}"
        counter += 1

    with pd.ExcelWriter(file_xuat, engine="openpyxl") as writer:
        df_ket_qua.to_excel(writer, sheet_name="TongHop_BenhNhan", index=False)
        ws = writer.sheets["TongHop_BenhNhan"]

        for col in ws.columns:
            ten_cot = str(col[0].value)
            if ten_cot in {"HO_TEN", "TEN_DICH_VU"}:
                width = 34 if ten_cot == "HO_TEN" else 46
            elif ten_cot in {"MA_BA", "MA_THE", "SO_PHIEU", "NGAY_TH_YL"}:
                width = 22
            else:
                width = 18
            ws.column_dimensions[col[0].column_letter].width = width

    return file_xuat

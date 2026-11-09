import os

import pandas as pd

from core.utils import doc_file_thong_minh


COT_BAT_BUOC = [
    "MA_DICH_VU",
    "TEN_DICH_VU",
    "MA_MAY",
    "Khớp MAMAY3176",
]


def _chuan_hoa_text(series):
    return series.fillna("").astype(str).str.strip()


def doc_file_doi_chieu_ma_may(path_file):
    ten_file_khong_ext = os.path.splitext(os.path.basename(path_file))[0]
    if not ten_file_khong_ext.lower().startswith("doi_chieu_ma_may"):
        raise ValueError("File không hợp lệ")

    df = doc_file_thong_minh(path_file)
    if df is None or df.empty:
        raise ValueError("File không hợp lệ")

    cot_thuc_te = {str(c).strip() for c in df.columns.tolist()}
    thieu_cot = [c for c in COT_BAT_BUOC if c not in cot_thuc_te]
    if thieu_cot:
        raise ValueError("File không hợp lệ")

    return df


def du_doan_ma_may(df_nguon):
    if df_nguon is None or df_nguon.empty:
        raise ValueError("Không có dữ liệu để dự đoán.")

    df = df_nguon.copy()
    for cot in ["MA_DICH_VU", "TEN_DICH_VU", "MA_MAY", "Khớp MAMAY3176"]:
        if cot not in df.columns:
            raise ValueError(f"Thiếu cột bắt buộc: {cot}")

    df["MA_DICH_VU"] = _chuan_hoa_text(df["MA_DICH_VU"])
    df["TEN_DICH_VU"] = _chuan_hoa_text(df["TEN_DICH_VU"])
    df["MA_MAY"] = _chuan_hoa_text(df["MA_MAY"])
    df["Khớp MAMAY3176"] = _chuan_hoa_text(df["Khớp MAMAY3176"])

    mask_tham_chieu = (df["Khớp MAMAY3176"].str.lower() == "có") & (df["MA_MAY"] != "")
    nguon_tham_chieu = df.loc[mask_tham_chieu, ["MA_DICH_VU", "TEN_DICH_VU", "MA_MAY"]].copy()

    map_unique = {}
    map_all = {}
    for (ma_dv, ten_dv), group in nguon_tham_chieu.groupby(["MA_DICH_VU", "TEN_DICH_VU"], dropna=False):
        ma_may_unique = sorted({str(v).strip() for v in group["MA_MAY"].tolist() if str(v).strip()})
        map_all[(ma_dv, ten_dv)] = " | ".join(ma_may_unique)
        if len(ma_may_unique) == 1:
            map_unique[(ma_dv, ten_dv)] = ma_may_unique[0]

    def tim_tham_chieu(row):
        key = (row["MA_DICH_VU"], row["TEN_DICH_VU"])
        return map_all.get(key, "")

    df["THAM_CHIEU_MA_MAY"] = df.apply(tim_tham_chieu, axis=1)

    mask_ma_may_trong = df["MA_MAY"] == ""

    def gia_tri_dien(row):
        key = (row["MA_DICH_VU"], row["TEN_DICH_VU"])
        return map_unique.get(key, "")

    df.loc[mask_ma_may_trong, "MA_MAY"] = df.loc[mask_ma_may_trong].apply(gia_tri_dien, axis=1)

    return df.reset_index(drop=True)


def xuat_excel_ma_may_xu_ly(df_ket_qua, file_xuat):
    if df_ket_qua is None or df_ket_qua.empty:
        raise ValueError("Không có dữ liệu để xuất file.")

    file_xuat_goc = file_xuat
    base, extension = os.path.splitext(file_xuat_goc)
    counter = 1
    while os.path.exists(file_xuat):
        file_xuat = f"{base} ({counter}){extension}"
        counter += 1

    with pd.ExcelWriter(file_xuat, engine="openpyxl") as writer:
        df_ket_qua.to_excel(writer, sheet_name="XuLy_MaMay", index=False)
        ws = writer.sheets["XuLy_MaMay"]

        for col in ws.columns:
            ten_cot = str(col[0].value)
            if ten_cot == "TEN_DICH_VU":
                width = 42
            elif ten_cot in {"MA_DICH_VU", "MA_MAY", "THAM_CHIEU_MA_MAY"}:
                width = 22
            elif ten_cot in {"Khớp MAMAY3176", "Khớp MAMAY", "Không tồn tại trong DM"}:
                width = 20
            else:
                width = 18
            ws.column_dimensions[col[0].column_letter].width = width

    return file_xuat

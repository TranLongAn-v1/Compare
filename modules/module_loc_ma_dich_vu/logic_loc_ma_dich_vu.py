import os

import pandas as pd

from core.utils import doc_file_thong_minh


COT_BAT_BUOC = ["MA_DICH_VU", "TEN_DICH_VU", "MA_MAY", "MA_BA", "HO_TEN"]


def doc_va_kiem_tra_file(path_file):
    ten_file_khong_ext = os.path.splitext(os.path.basename(path_file))[0]
    if not ten_file_khong_ext.lower().startswith("tong_hop_xml3"):
        raise ValueError("File không hợp lệ")

    df = doc_file_thong_minh(path_file)
    if df is None or df.empty:
        raise ValueError("File không hợp lệ")

    cot_thuc_te = [str(c).strip() for c in df.columns.tolist()]
    if set(cot_thuc_te) != set(COT_BAT_BUOC):
        raise ValueError("File không hợp lệ")

    return df


def loc_trung_lap_ma_dich_vu(df_nguon):
    if df_nguon is None or df_nguon.empty:
        raise ValueError("Không có dữ liệu để lọc.")

    df = df_nguon.copy()
    for cot in ["MA_DICH_VU", "TEN_DICH_VU"]:
        if cot not in df.columns:
            raise ValueError(f"Thiếu cột bắt buộc: {cot}")
        df[cot] = df[cot].fillna("").astype(str).str.strip()

    df_loc = df.drop_duplicates(subset=["MA_DICH_VU", "TEN_DICH_VU"], keep="first").reset_index(drop=True)
    return df_loc


def xuat_excel_loc(df_ket_qua, file_xuat):
    if df_ket_qua is None or df_ket_qua.empty:
        raise ValueError("Không có dữ liệu để xuất file.")

    file_xuat_goc = file_xuat
    base, extension = os.path.splitext(file_xuat_goc)
    counter = 1
    while os.path.exists(file_xuat):
        file_xuat = f"{base} ({counter}){extension}"
        counter += 1

    with pd.ExcelWriter(file_xuat, engine="openpyxl") as writer:
        df_ket_qua.to_excel(writer, sheet_name="Loc_Ma_Dich_Vu", index=False)
        ws = writer.sheets["Loc_Ma_Dich_Vu"]

        for col in ws.columns:
            ten_cot = str(col[0].value)
            if ten_cot == "TEN_DICH_VU":
                width = 42
            elif ten_cot in {"MA_DICH_VU", "MA_MAY"}:
                width = 20
            else:
                width = 18
            ws.column_dimensions[col[0].column_letter].width = width

    return file_xuat

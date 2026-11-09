# modules/module_c79a/logic_c79a.py
import os
import re
import pandas as pd
from openpyxl.styles import PatternFill
from core.utils import doc_file_thong_minh, lam_sach_chuoi

COL_NAMES_MAP = {
    "(1)": "Tổng cộng (1)",
    "(2)": "Xét nghiệm",
    "(3)": "CĐHA TDCN",
    "(4)": "Thuốc (4)",
    "(5)": "Máu",
    "(6)": "TTPT",
    "(7)": "VTYT (7)",
    "(8)": "DV KT",
    "(9)": "Thuốc tỷ lệ (9)",
    "(10)": "VTYT tỷ lệ (10)",
    "(11)": "Tiền khám",
    "(12)": "Tiền giường",
    "(13)": "Vận chuyển",
    "(14)": "Người bệnh chi trả",
    "(15)": "BHXH thanh toán (15)",
    "(16)": "Ngoài định suất (16)"
}

def safe_float(val):
    try:
        if pd.isna(val) or val == "":
            return 0.0
        clean_val = str(val).replace(",", "").strip()
        return float(clean_val)
    except Exception:
        return 0.0

def is_valid_stt(val):
    if pd.isna(val):
        return False
    val_str = str(val).strip()
    if not val_str:
        return False
    if val_str.isdigit():
        return True
    try:
        f = float(val_str)
        return f.is_integer()
    except ValueError:
        return False

def clean_name(name_val):
    if pd.isna(name_val):
        return ""
    return re.sub(r"\s+", " ", str(name_val).strip()).upper()

def clean_card(card_val):
    if pd.isna(card_val):
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(card_val).strip().upper())

def clean_ma_dk(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

def chuan_hoa_ngay(date_val):
    if pd.isna(date_val):
        return ""
    s = str(date_val).strip()
    if not s:
        return ""
    if s.endswith('.0'):
        s = s[:-2]
    parts = s.split()
    date_part = parts[0] if parts else ""
    for sep in ['/', '-']:
        subparts = date_part.split(sep)
        if len(subparts) == 3:
            d, m, y = subparts[0], subparts[1], subparts[2]
            if len(d) == 1: d = "0" + d
            if len(m) == 1: m = "0" + m
            return f"{d}/{m}/{y}"
    return s

def extract_data_and_financials(file_path):
    df = doc_file_thong_minh(file_path)
    idx_marker = None
    row_vals = None
    for idx in range(len(df)):
        r_vals = [str(x).strip() for x in df.iloc[idx]]
        if '(1)' in r_vals and '(2)' in r_vals:
            idx_marker = idx
            row_vals = r_vals
            break
    if idx_marker is None:
        raise ValueError(f"Không tìm thấy dòng chỉ mục (1)...(16) trong file {os.path.basename(file_path)}")
    
    col_map = {}
    for i, val in enumerate(row_vals):
        val_str = str(val).strip()
        if val_str and val_str != 'nan':
            col_map[val_str] = i
            
    stt_col = col_map.get('A')
    name_col = col_map.get('B')
    card_col = col_map.get('E')
    madk_col = col_map.get('F')
    if madk_col is None:
        madk_col = col_map.get('G')
    date_col = col_map.get('I')
    
    if stt_col is None or name_col is None or card_col is None or date_col is None:
        raise KeyError(f"Thiếu một số cột thiết yếu (STT, Họ tên, Mã thẻ BHYT, Ngày) trong file {os.path.basename(file_path)}")
        
    financial_indices = {}
    for k in range(1, 17):
        col_idx = col_map.get(f'({k})')
        if col_idx is None:
            raise KeyError(f"Thiếu cột tài chính ({k}) trong file {os.path.basename(file_path)}")
        financial_indices[k] = col_idx
        
    data_rows = []
    for idx in range(idx_marker + 1, len(df)):
        row = df.iloc[idx]
        stt_val = row.iloc[stt_col]
        name_val = row.iloc[name_col]
        if is_valid_stt(stt_val) and pd.notna(name_val) and str(name_val).strip() != "":
            row_dict = {
                'Họ và tên': clean_name(name_val),
                'Mã thẻ BHYT': clean_card(row.iloc[card_col]),
                'Ngày': chuan_hoa_ngay(row.iloc[date_col]),
                'Mã ĐK': clean_ma_dk(row.iloc[madk_col]) if madk_col is not None else "",
                'STT_goc': int(float(str(stt_val).strip())),
            }
            for k in range(1, 17):
                row_dict[f'val_{k}'] = safe_float(row.iloc[financial_indices[k]])
            data_rows.append(row_dict)
            
    return pd.DataFrame(data_rows)

def tao_du_lieu_doi_chieu_c79a(file_std, file_por, tu_ngay, den_ngay, callback_status):
    try:
        callback_status("Đang đọc dữ liệu từ file chuẩn...", "blue")
        df_std = extract_data_and_financials(file_std)
        
        callback_status("Đang đọc dữ liệu từ file cổng giám định...", "blue")
        df_por = extract_data_and_financials(file_por)
        
        if df_std.empty:
            return False, "File chuẩn không chứa dòng dữ liệu bệnh nhân hợp lệ nào."
        if df_por.empty:
            return False, "File cổng giám định không chứa dòng dữ liệu bệnh nhân hợp lệ nào."
            
        callback_status("Đang chuẩn hóa ngày tháng...", "blue")
        df_std["Ngay_DT"] = pd.to_datetime(df_std["Ngày"], format="%d/%m/%Y", errors="coerce")
        df_por["Ngay_DT"] = pd.to_datetime(df_por["Ngày"], format="%d/%m/%Y", errors="coerce")
        
        callback_status("Đang lọc dữ liệu theo khoảng thời gian...", "blue")
        tu_ngay_dt = pd.to_datetime(tu_ngay)
        den_ngay_dt = pd.to_datetime(den_ngay) + pd.Timedelta(days=1, seconds=-1)
        
        df_std = df_std.loc[(df_std["Ngay_DT"] >= tu_ngay_dt) & (df_std["Ngay_DT"] <= den_ngay_dt)].copy()
        df_por = df_por.loc[(df_por["Ngay_DT"] >= tu_ngay_dt) & (df_por["Ngay_DT"] <= den_ngay_dt)].copy()
        
        if df_std.empty:
            return False, "Không có dữ liệu bệnh nhân trong file chuẩn phù hợp với bộ lọc thời gian."
            
        callback_status("Đang ghép nối và đối chiếu dữ liệu...", "blue")
        df_std["LAN_KHAM"] = df_std.groupby(["Họ và tên", "Mã thẻ BHYT", "Ngày"]).cumcount()
        df_por["LAN_KHAM"] = df_por.groupby(["Họ và tên", "Mã thẻ BHYT", "Ngày"]).cumcount()
        
        merged = pd.merge(
            df_std,
            df_por,
            on=["Họ và tên", "Mã thẻ BHYT", "Ngày", "LAN_KHAM"],
            how="outer",
            suffixes=("_std", "_por")
        )
        
        rows = []
        for _, row in merged.iterrows():
            ton_tai_std = pd.notna(row["STT_goc_std"])
            ton_tai_por = pd.notna(row["STT_goc_por"])
            
            ho_ten = row["Họ và tên"]
            ma_the = row["Mã thẻ BHYT"]
            ngay = row["Ngày"]
            
            color = "none"
            ghi_chu = "Trùng khớp"
            
            ma_dk = ""
            if ton_tai_std and pd.notna(row["Mã ĐK_std"]):
                ma_dk = str(row["Mã ĐK_std"]).strip()
            if not ma_dk and ton_tai_por and pd.notna(row["Mã ĐK_por"]):
                ma_dk = str(row["Mã ĐK_por"]).strip()

            # Khởi tạo dictionary để lưu giá trị chi tiết phục vụ cho việc xuất file excel sau này
            detail_dict = {
                "Họ và tên": ho_ten,
                "Mã thẻ BHYT": ma_the,
                "Mã ĐK": ma_dk,
                "Ngày vào": ngay,
                "Ngày (Chuẩn)": ngay if ton_tai_std else "",
                "Ngày (Cổng GD)": ngay if ton_tai_por else "",
            }
            
            for k in range(1, 17):
                v_std = row[f"val_{k}_std"] if ton_tai_std else 0.0
                v_por = row[f"val_{k}_por"] if ton_tai_por else 0.0
                detail_dict[f"val_{k}_std"] = v_std
                detail_dict[f"val_{k}_por"] = v_por
                detail_dict[f"val_{k}_diff"] = v_std - v_por
            
            if ton_tai_std and not ton_tai_por:
                ghi_chu = "Bệnh nhân này có trong file chuẩn có nhưng file trên cổng giám định không có"
                color = "blue"
            elif ton_tai_por and not ton_tai_std:
                ghi_chu = "Bệnh nhân có trên file cổng giám định nhưng file chuẩn thì không có"
                color = "yellow"
            else:
                diffs = []
                for k in range(1, 17):
                    v_std = row[f"val_{k}_std"]
                    v_por = row[f"val_{k}_por"]
                    # So sánh tài chính với sai số tuyệt đối lớn hơn 1.0 (hoặc có thể dùng 0.01)
                    if abs(v_std - v_por) > 1.0:
                        col_friendly_name = COL_NAMES_MAP[f"({k})"]
                        diffs.append(col_friendly_name)
                if diffs:
                    ghi_chu = "Lệch: " + ", ".join(diffs)
                    color = "purple"
            
            detail_dict["Ghi chú"] = ghi_chu
            detail_dict["color"] = color
            rows.append(detail_dict)
            
        df_result = pd.DataFrame(rows)
        
        # Sắp xếp các lỗi lên đầu theo yêu cầu
        cat_order = [
            "Bệnh nhân có trên file cổng giám định nhưng file chuẩn thì không có",
            "Bệnh nhân này có trong file chuẩn có nhưng file trên cổng giám định không có",
            "purple", # Phân biệt lệch tài chính thông qua color hoặc ghi chú chứa 'Lệch:'
            "Trùng khớp"
        ]
        
        def assign_sort_priority(row_item):
            gc = row_item["Ghi chú"]
            col = row_item["color"]
            if gc == "Bệnh nhân có trên file cổng giám định nhưng file chuẩn thì không có":
                return 0
            elif gc == "Bệnh nhân này có trong file chuẩn có nhưng file trên cổng giám định không có":
                return 1
            elif col == "purple":
                return 2
            else:
                return 3
                
        df_result["sort_priority"] = df_result.apply(assign_sort_priority, axis=1)
        df_result = df_result.sort_values(by="sort_priority").reset_index(drop=True)
        df_result = df_result.drop(columns=["sort_priority"])
        
        # Thêm cột STT
        df_result.insert(0, "STT", range(1, len(df_result) + 1))
        
        callback_status("Hoàn thành đối chiếu.", "green")
        return True, df_result
    except Exception as e:
        callback_status("Lỗi xảy ra trong quá trình đối chiếu.", "red")
        return False, str(e)

def xuat_excel_c79a(df_full_result, file_xuat):
    file_xuat_goc = file_xuat
    base, extension = os.path.splitext(file_xuat_goc)
    counter = 1
    while os.path.exists(file_xuat):
        file_xuat = f"{base} ({counter}){extension}"
        counter += 1
        
    # Chuẩn bị dữ liệu ghi vào Excel
    # Chúng ta sẽ cấu trúc bảng như sau:
    # STT, Họ và tên, Mã thẻ BHYT, Ngày (Chuẩn), Ngày (Cổng GD), Ghi chú
    # Kế tiếp là các nhóm cột tài chính, mỗi nhóm gồm: [Tên Cột] (Chuẩn), [Tên Cột] (Cổng GD), [Tên Cột] (Chênh lệch)
    cols_basic = ["STT", "Họ và tên", "Mã thẻ BHYT", "Mã ĐK", "Ngày (Chuẩn)", "Ngày (Cổng GD)", "Ghi chú"]
    
    excel_rows = []
    for _, row in df_full_result.iterrows():
        excel_row = {c: row[c] for c in cols_basic}
        for k in range(1, 17):
            friendly = COL_NAMES_MAP[f"({k})"]
            excel_row[f"{friendly} (Chuẩn)"] = row[f"val_{k}_std"]
            excel_row[f"{friendly} (Cổng GD)"] = row[f"val_{k}_por"]
            excel_row[f"{friendly} (Lệch)"] = row[f"val_{k}_diff"]
        excel_rows.append(excel_row)
        
    df_excel = pd.DataFrame(excel_rows)
    
    # Ghi ra file excel sử dụng openpyxl
    with pd.ExcelWriter(file_xuat, engine="openpyxl") as writer:
        df_excel.to_excel(writer, sheet_name="Đối chiếu C79a", index=False)
        worksheet = writer.sheets["Đối chiếu C79a"]
        
        colors = {
            "blue": PatternFill(start_color="B3D9FF", end_color="B3D9FF", fill_type="solid"),  # xanh da trời
            "yellow": PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid"),  # vàng
            "purple": PatternFill(start_color="E6CCFF", end_color="E6CCFF", fill_type="solid"),  # tím
        }
        
        # Tô màu các dòng dựa trên cột color gốc
        for row_idx, mau_to in enumerate(df_full_result["color"]):
            excel_row_num = row_idx + 2  # dòng 1 là tiêu đề
            fill = colors.get(mau_to, None)
            if fill:
                for col_idx in range(1, len(df_excel.columns) + 1):
                    worksheet.cell(row=excel_row_num, column=col_idx).fill = fill
                    
        # Điều chỉnh độ rộng cột cơ bản
        for col in worksheet.columns:
            header_val = col[0].value
            if header_val == "STT":
                worksheet.column_dimensions[col[0].column_letter].width = 8
            elif header_val in ["Họ và tên", "Ngày (Chuẩn)", "Ngày (Cổng GD)"]:
                worksheet.column_dimensions[col[0].column_letter].width = 20
            elif header_val == "Mã thẻ BHYT":
                worksheet.column_dimensions[col[0].column_letter].width = 18
            elif header_val == "Mã ĐK":
                worksheet.column_dimensions[col[0].column_letter].width = 12
            elif header_val == "Ghi chú":
                worksheet.column_dimensions[col[0].column_letter].width = 45
            else:
                worksheet.column_dimensions[col[0].column_letter].width = 16
                
            # Định dạng số cho các cột tài chính
            if header_val not in cols_basic:
                for cell in col[1:]:
                    cell.number_format = "#,##0.00"
                    
    return file_xuat

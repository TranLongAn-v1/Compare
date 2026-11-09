import pandas as pd
from openpyxl.styles import PatternFill
from core.utils import chuyen_doi_ngay_7980, lam_sach_chuoi, doc_file_thong_minh
import os

def xu_ly_doi_chieu_4210(file_vp_chuoi, file_7980, file_xuat, tu_ngay, den_ngay, callback_status):
    try:
        # --- 1. XỬ LÝ TÊN FILE TRÙNG ---
        file_xuat_goc = file_xuat
        base, extension = os.path.splitext(file_xuat_goc)
        counter = 1
        while os.path.exists(file_xuat):
            file_xuat = f"{base} ({counter}){extension}"
            counter += 1

        callback_status("Đang bóc tách dữ liệu từ file...", "blue")
        
        # --- 2. NÂNG CẤP LOGIC: TỰ ĐỘNG GỘP NHIỀU FILE/THƯ MỤC VIỆN PHÍ ---
        danh_sach_duong_dan = [f.strip() for f in file_vp_chuoi.split(';') if f.strip()]
        danh_sach_df_vp = []
        
        for duong_dan in danh_sach_duong_dan:
            if os.path.isfile(duong_dan):
                df_tam = doc_file_thong_minh(duong_dan)
                if df_tam is not None and not df_tam.empty:
                    danh_sach_df_vp.append(df_tam)
            elif os.path.isdir(duong_dan):
                for root, dirs, files in os.walk(duong_dan):
                    for file_name in files:
                        if file_name.lower().endswith(('.xlsx', '.xls', '.csv')):
                            file_path = os.path.join(root, file_name)
                            df_tam = doc_file_thong_minh(file_path)
                            if df_tam is not None and not df_tam.empty:
                                danh_sach_df_vp.append(df_tam)
        
        if not danh_sach_df_vp:
            raise ValueError("Không tìm thấy dữ liệu Viện phí hợp lệ. Vui lòng kiểm tra lại file/thư mục đầu vào.")
            
        df_vp = pd.concat(danh_sach_df_vp, ignore_index=True)
        
        # --- 3. ĐỌC FILE 79/80 VÀ CHUẨN HÓA DỮ LIỆU ---
        df_7980 = doc_file_thong_minh(file_7980)
        
        callback_status("Đang chuẩn hóa dữ liệu...", "blue")
        df_vp.columns = df_vp.columns.str.strip()
        df_7980.columns = df_7980.columns.str.strip()

        df_vp['Mã bệnh nhân'] = df_vp['Mã bệnh nhân'].apply(lam_sach_chuoi)
        df_vp['Họ tên'] = df_vp['Họ tên'].apply(lam_sach_chuoi)
        
        df_vp['Vào viện chuẩn'] = df_vp['Vào viện'].apply(lambda x: str(x).strip()[:16] if pd.notna(x) and str(x).strip() != 'nan' else "")
        df_vp['Ra viện chuẩn'] = df_vp['Ra viện'].apply(lambda x: str(x).strip()[:16] if pd.notna(x) and str(x).strip() != 'nan' else "")

        df_7980['MA_BN'] = df_7980['MA_BN'].apply(lam_sach_chuoi)
        df_7980['HO_TEN'] = df_7980['HO_TEN'].apply(lam_sach_chuoi)
        df_7980['NGAY_VAO_CHUAN'] = df_7980['NGAY_VAO'].apply(chuyen_doi_ngay_7980)
        df_7980['NGAY_RA_CHUAN'] = df_7980['NGAY_RA'].apply(chuyen_doi_ngay_7980)

        df_vp_valid = df_vp[df_vp['Mã bệnh nhân'] != ""].copy()
        df_7980_valid = df_7980[df_7980['MA_BN'] != ""].copy()
        
        # --- 4. LỌC THEO THỜI GIAN & ĐỐI CHIẾU DỮ LIỆU ---
        callback_status("Đang lọc theo khoảng thời gian...", "blue")
        df_vp_valid['Ngay_Ra_DT'] = pd.to_datetime(df_vp_valid['Ra viện chuẩn'], format='%d/%m/%Y %H:%M', errors='coerce')
        df_7980_valid['Ngay_Ra_DT'] = pd.to_datetime(df_7980_valid['NGAY_RA_CHUAN'], format='%d/%m/%Y %H:%M', errors='coerce')
        
        tu_ngay_dt = pd.to_datetime(tu_ngay)
        den_ngay_dt = pd.to_datetime(den_ngay) + pd.Timedelta(days=1, seconds=-1)
        
        mask_vp = (df_vp_valid['Ngay_Ra_DT'] >= tu_ngay_dt) & (df_vp_valid['Ngay_Ra_DT'] <= den_ngay_dt)
        mask_7980 = (df_7980_valid['Ngay_Ra_DT'] >= tu_ngay_dt) & (df_7980_valid['Ngay_Ra_DT'] <= den_ngay_dt)
        
        df_vp_valid = df_vp_valid.loc[mask_vp].copy()
        df_7980_valid = df_7980_valid.loc[mask_7980].copy()

        callback_status("Đang đối chiếu dữ liệu...", "blue")
        df_vp_valid['Sort_Date'] = pd.to_datetime(df_vp_valid['Vào viện chuẩn'], format='%d/%m/%Y %H:%M', errors='coerce')
        df_7980_valid['Sort_Date'] = pd.to_datetime(df_7980_valid['NGAY_VAO_CHUAN'], format='%d/%m/%Y %H:%M', errors='coerce')
        
        # --- [GHI CHÚ SỬA ĐỔI]: TẠO Ổ KHÓA GHÉP THEO NGÀY ---
        df_vp_valid['Ngay_Vao_Date_Str'] = df_vp_valid['Sort_Date'].dt.strftime('%d/%m/%Y').fillna('Unk')
        df_7980_valid['Ngay_Vao_Date_Str'] = df_7980_valid['Sort_Date'].dt.strftime('%d/%m/%Y').fillna('Unk')

        df_vp_valid = df_vp_valid.sort_values(by=['Mã bệnh nhân', 'Sort_Date'])
        df_7980_valid = df_7980_valid.sort_values(by=['MA_BN', 'Sort_Date'])
        
        # Đếm số lần khám CỦA TỪNG BỆNH NHÂN TRONG CÙNG 1 NGÀY VÀO
        df_vp_valid['LAN_KHAM'] = df_vp_valid.groupby(['Mã bệnh nhân', 'Ngay_Vao_Date_Str']).cumcount()
        df_7980_valid['LAN_KHAM'] = df_7980_valid.groupby(['MA_BN', 'Ngay_Vao_Date_Str']).cumcount()
        
        # Ép dữ liệu phải khớp cả Mã Bệnh Nhân, Ngày Vào, và số lần khám trong ngày
        merged = pd.merge(
            df_vp_valid, df_7980_valid, 
            left_on=['Mã bệnh nhân', 'Ngay_Vao_Date_Str', 'LAN_KHAM'], 
            right_on=['MA_BN', 'Ngay_Vao_Date_Str', 'LAN_KHAM'], 
            how='outer'
        )
        
        danh_sach_tong_hop = []
        for index, row in merged.iterrows():
            ton_tai_vp = pd.notna(row['Mã bệnh nhân'])
            ton_tai_7980 = pd.notna(row['MA_BN'])
            ma_bn_hien_thi = lam_sach_chuoi(row['Mã bệnh nhân']) if ton_tai_vp else lam_sach_chuoi(row['MA_BN'])
            
            phan_loai = ""
            mau_to = "none"
            if ton_tai_vp and not ton_tai_7980:
                phan_loai, mau_to = "Có trong Viện phí nhưng không có trong 79_80", "yellow"
            elif ton_tai_7980 and not ton_tai_vp:
                phan_loai, mau_to = "Có trong 79_80 nhưng không có trong Viện phí", "blue"
            else:
                if (lam_sach_chuoi(row['Vào viện chuẩn']) != lam_sach_chuoi(row['NGAY_VAO_CHUAN'])) or (lam_sach_chuoi(row['Ra viện chuẩn']) != lam_sach_chuoi(row['NGAY_RA_CHUAN'])):
                    phan_loai = "Sai lệch Ngày vào/ra"
                else:
                    phan_loai = "Trùng khớp"
                    
            danh_sach_tong_hop.append({
                'MÃ BỆNH NHÂN': ma_bn_hien_thi,
                'TÊN BỆNH NHÂN': row['Họ tên'] if ton_tai_vp else row['HO_TEN'],
                'SỐ LƯỢNG': 1,
                'NGÀY VÀO (VP)': lam_sach_chuoi(row['Vào viện chuẩn']), 'NGÀY VÀO (79_80)': lam_sach_chuoi(row['NGAY_VAO_CHUAN']),
                'NGÀY RA (VP)': lam_sach_chuoi(row['Ra viện chuẩn']), 'NGÀY RA (79_80)': lam_sach_chuoi(row['NGAY_RA_CHUAN']),
                'PHÂN LOẠI': phan_loai, 'color': mau_to 
            })

        df_tong_hop = pd.DataFrame(danh_sach_tong_hop)
        
        thu_tu_uu_tien = ["Có trong Viện phí nhưng không có trong 79_80", "Có trong 79_80 nhưng không có trong Viện phí", "Sai lệch Ngày vào/ra", "Trùng khớp"]
        df_tong_hop['PHÂN LOẠI'] = pd.Categorical(df_tong_hop['PHÂN LOẠI'], categories=thu_tu_uu_tien, ordered=True)
        df_tong_hop = df_tong_hop.sort_values('PHÂN LOẠI').reset_index(drop=True)
        df_tong_hop.insert(0, 'STT', range(1, len(df_tong_hop) + 1))

        # --- 5. XUẤT FILE EXCEL VÀ FORMAT ---
        callback_status("Đang xuất file Excel...", "blue")

        with pd.ExcelWriter(file_xuat, engine='openpyxl') as writer:
            df_xuat_excel = df_tong_hop.drop(columns=['color', 'Ngay_Vao_Date_Str'], errors='ignore') # Thêm errors='ignore' để phòng trường hợp cột này lọt ra ngoài
            df_xuat_excel.to_excel(writer, sheet_name='VP_vs_7980', index=False)
            
            worksheet = writer.sheets['VP_vs_7980']
            blue_fill = PatternFill(start_color='99CCFF', end_color='99CCFF', fill_type='solid') 
            yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
            
            for row_idx, mau_to in enumerate(df_tong_hop['color']):
                excel_row = row_idx + 2 
                fill = blue_fill if mau_to == 'blue' else (yellow_fill if mau_to == 'yellow' else None)
                if fill:
                    for col_idx in range(1, len(df_xuat_excel.columns) + 1):
                        worksheet.cell(row=excel_row, column=col_idx).fill = fill
            
            for col in worksheet.columns:
                ten_cot = str(col[0].value) if col[0].value else ""
                
                if ten_cot in ['STT', 'SỐ LƯỢNG']:
                    chieu_rong = 10
                elif ten_cot in ['MÃ BỆNH NHÂN']:
                    chieu_rong = 15
                elif ten_cot in ['NGÀY VÀO (VP)', 'NGÀY VÀO (79_80)', 'NGÀY RA (VP)', 'NGÀY RA (79_80)']:
                    chieu_rong = 18
                elif ten_cot == 'TÊN BỆNH NHÂN':
                    chieu_rong = 25
                elif ten_cot == 'PHÂN LOẠI':
                    chieu_rong = 45
                else:
                    chieu_rong = 15
                    
                worksheet.column_dimensions[col[0].column_letter].width = chieu_rong
                
        callback_status(f"Đã xuất file thành công: {file_xuat}", "green")
        
        # --- 6. TỰ ĐỘNG MỞ FILE VÀ TRẢ VỀ ĐƯỜNG DẪN MỚI ---
        try:
            os.startfile(file_xuat)
        except Exception as e:
            pass # Nếu máy tính không có phần mềm đọc Excel thì bỏ qua lỗi này
            
        return True, file_xuat # Trả về đường dẫn thực tế (đã gắn số nếu trùng)

    except Exception as e:
        callback_status(f"Lỗi: {str(e)}", "red")
        return False, f"Đã xảy ra lỗi: {str(e)}"
import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
import os

# --- CẤU HÌNH DANH SÁCH MÃ & TÊN CÔNG TY (BẢN V4: MEMORY & AI FULL) ---
THONG_TIN_CO_PHIEU = {
    # =================================================================
    # 💾 NHÓM 1: MEMORY & STORAGE (BỘ NHỚ & LƯU TRỮ) - "Kho lương thực AI"
    # =================================================================
    "8299.TWO": {"Ten": "Phison (Electronics)", "Nganh": "Memory - Controller"}, # Trùm Controller SSD
    "2408.TW": {"Ten": "Nanya Technology", "Nganh": "Memory - DRAM"}, # Sản xuất DRAM lớn nhất ĐL
    "2344.TW": {"Ten": "Winbond Elec", "Nganh": "Memory - Flash/DRAM"},
    "2337.TW": {"Ten": "Macronix (MXIC)", "Nganh": "Memory - NOR Flash"}, # Chuyên chip nhớ cho xe hơi/Nintendo
    "3260.TWO": {"Ten": "ADATA", "Nganh": "Memory - Module"}, # Bán RAM/SSD (Thương mại)
    "2451.TW": {"Ten": "Transcend Info", "Nganh": "Memory - Module"},
    "4967.TW": {"Ten": "TeamGroup", "Nganh": "Memory - Module"},
    "8150.TW": {"Ten": "ChipMOS", "Nganh": "Memory - Packaging"}, # Đóng gói chip nhớ
    "6239.TW": {"Ten": "PTI (Powertech)", "Nganh": "Memory - Packaging"}, # Đóng gói chip nhớ (Đối tác Micron)

    # =================================================================
    # 🏭 NHÓM 2: FOUNDRY (SẢN XUẤT CHIP)
    # =================================================================
    "2330.TW": {"Ten": "TSMC", "Nganh": "Foundry - Logic"},
    "2303.TW": {"Ten": "UMC", "Nganh": "Foundry - Logic"},
    "6770.TW": {"Ten": "PSMC (Powerchip)", "Nganh": "Foundry - Memory"}, # Chuyên gia công chip nhớ
    "5347.TWO": {"Ten": "VIS (Vanguard)", "Nganh": "Foundry - 8inch"},

    # =================================================================
    # 🧠 NHÓM 3: IC DESIGN & IP (THIẾT KẾ & BẢN QUYỀN)
    # =================================================================
    "2454.TW": {"Ten": "MediaTek", "Nganh": "IC Design - Mobile/AI"},
    "3034.TW": {"Ten": "Novatek", "Nganh": "IC Design - Display"},
    "2379.TW": {"Ten": "Realtek", "Nganh": "IC Design - Network"},
    "5269.TW": {"Ten": "ASMedia", "Nganh": "IC Design - High Speed"}, # Controller USB/PCIe
    "3443.TW": {"Ten": "GUC (Global Unichip)", "Nganh": "Design Service (AI)"},
    "3661.TW": {"Ten": "Alchip", "Nganh": "Design Service (AI)"},
    "3035.TW": {"Ten": "Faraday Tech", "Nganh": "Design Service"},
    "8096.TWO": {"Ten": "CoAsia", "Nganh": "Design Service"}, # Đối tác Samsung
    "3529.TWO": {"Ten": "eMemory", "Nganh": "IP Core"},
    "6533.TW": {"Ten": "Andes Tech", "Nganh": "IP Core (RISC-V)"},

    # =================================================================
    # 📡 NHÓM 4: COMPOUND SEMI (BÁN DẪN HỢP CHẤT) - 5G/QUANG HỌC
    # =================================================================
    "2455.TW": {"Ten": "Visual Photonics (VPEC)", "Nganh": "Compound Semi"}, # GaAs wafers
    "3105.TWO": {"Ten": "Win Semi", "Nganh": "Compound Semi"},
    "8086.TWO": {"Ten": "AWSC", "Nganh": "Compound Semi"},
    "3707.TW": {"Ten": "Epistar (Ennostar)", "Nganh": "Compound/LED"},

    # =================================================================
    # 📦 NHÓM 5: OSAT & EQUIPMENT (HẬU CẦN)
    # =================================================================
    "3711.TW": {"Ten": "ASE Tech", "Nganh": "OSAT (Packaging)"},
    "2449.TW": {"Ten": "KYEC", "Nganh": "OSAT (Testing)"},
    "6488.TW": {"Ten": "GlobalWafers", "Nganh": "Wafer (Material)"},
    "5483.TWO": {"Ten": "Sino-American", "Nganh": "Wafer (Material)"},
    "3680.TW": {"Ten": "Gudeng", "Nganh": "Equipment (EUV Pod)"},

    # =================================================================
    # 🖥️ NHÓM 6: AI SERVER & PC (HẠ TẦNG PHẦN CỨNG)
    # =================================================================
    "2317.TW": {"Ten": "Foxconn", "Nganh": "AI Server/OEM"},
    "2382.TW": {"Ten": "Quanta", "Nganh": "AI Server/OEM"},
    "3231.TW": {"Ten": "Wistron", "Nganh": "AI Server/OEM"},
    "2356.TW": {"Ten": "Inventec", "Nganh": "AI Server/OEM"},
    "2376.TW": {"Ten": "Gigabyte", "Nganh": "AI Server/Brand"},
    "2357.TW": {"Ten": "Asus", "Nganh": "PC/Brand"},
    "2301.TW": {"Ten": "Lite-On", "Nganh": "Power Supply"},
    "2308.TW": {"Ten": "Delta Elec", "Nganh": "Power Supply"},

    # =================================================================
    # 📺 NHÓM 7: DISPLAY & COMPONENTS (MÀN HÌNH & LINH KIỆN)
    # =================================================================
    "2409.TW": {"Ten": "AUO", "Nganh": "Display Panel"}, # Bán nhà máy cho Micron
    "3481.TW": {"Ten": "Innolux", "Nganh": "Display Panel"},
    "3008.TW": {"Ten": "Largan", "Nganh": "Optics (Lens)"},
    "3037.TW": {"Ten": "Unimicron", "Nganh": "PCB (ABF)"},
    "2327.TW": {"Ten": "Yageo", "Nganh": "Passive Comp"},

    # =================================================================
    # 🏦 NHÓM 8: TÀI CHÍNH & VẬN TẢI (TRỤ CỘT)
    # =================================================================
    "2881.TW": {"Ten": "Fubon Fin", "Nganh": "Financial"},
    "2882.TW": {"Ten": "Cathay Fin", "Nganh": "Financial"},
    "2603.TW": {"Ten": "Evergreen", "Nganh": "Shipping"},
    "2002.TW": {"Ten": "China Steel", "Nganh": "Steel"}
}

DANH_SACH_MA = list(THONG_TIN_CO_PHIEU.keys())
WINDOW_LONG = 21 

print(f"🚀 Đang tải dữ liệu {len(DANH_SACH_MA)} mã chứng khoán Đài Loan (Bản V4 - Memory Full)...")

try:
    # Tải dữ liệu hàng loạt (Bulk Download)
    data = yf.download(DANH_SACH_MA, period="3mo", group_by='ticker', auto_adjust=True, threads=True)
except Exception as e:
    print("❌ Lỗi kết nối:", e)
    exit()

ket_qua = []
print("⏳ Đang phân tích chi tiết từng mã...")

for ma in DANH_SACH_MA:
    try:
        # Xử lý MultiIndex
        if len(DANH_SACH_MA) == 1: df = data
        else:
            if ma not in data.columns.levels[0]: continue
            df = data[ma].dropna()

        if len(df) < WINDOW_LONG + 5: continue

        # --- TÍNH TOÁN ---
        gia_hien_tai = df['Close'].iloc[-1]
        vol_hien_tai = df['Volume'].iloc[-1]
        
        info = THONG_TIN_CO_PHIEU.get(ma, {"Ten": "Unknown", "Nganh": "Other"})
        
        # Tab 1: Tín hiệu ngày
        gia_hom_qua = df['Close'].iloc[-2]
        pct_doi_ngay = (gia_hien_tai - gia_hom_qua) / gia_hom_qua * 100
        sma_20 = df['Close'].rolling(window=20).mean().iloc[-1]
        vol_tb_20 = df['Volume'].rolling(window=20).mean().iloc[-1]
        
        tin_hieu_ngay = "Yếu"
        if gia_hien_tai > sma_20:
            if vol_hien_tai > vol_tb_20: tin_hieu_ngay = "Bùng nổ (Breakout)"
            else: tin_hieu_ngay = "Tích lũy (Up)"
        
        # Tab 2: Xu hướng dòng tiền
        gia_21_ngay_truoc = df['Close'].iloc[-(WINDOW_LONG)]
        pct_tang_1_thang = ((gia_hien_tai - gia_21_ngay_truoc) / gia_21_ngay_truoc) * 100
        
        vol_tb_5 = df['Volume'].rolling(window=5).mean().iloc[-1]
        suc_manh_dong_tien = (vol_tb_5 / vol_tb_20) if vol_tb_20 > 0 else 0
        gtgd_ty_twd = (gia_hien_tai * vol_tb_20) / 1_000_000_000 

        ket_qua.append({
            'Mã': ma.replace(".TW", "").replace(".TWO", ""), 
            'Tên Công Ty': info["Ten"],
            'Ngành': info["Nganh"],
            'Giá': round(gia_hien_tai, 1),
            '%_Ngày': round(pct_doi_ngay, 2),
            '%_Vol_vs_TB': round((vol_hien_tai/vol_tb_20)*100, 0) if vol_tb_20 > 0 else 0,
            'Tín_Hiệu_Ngày': tin_hieu_ngay,
            '%_Tăng_1_Tháng': round(pct_tang_1_thang, 2),
            'Sức_Mạnh_Dòng_Tiền': round(suc_manh_dong_tien, 2),
            'GTGD_TB_Tỷ': round(gtgd_ty_twd, 3)
        })
    except Exception as e:
        continue

# --- XUẤT FILE ---
if ket_qua:
    df_full = pd.DataFrame(ket_qua)
    
    # Sheet 1
    df_tab1 = df_full[['Mã', 'Tên Công Ty', 'Ngành', 'Giá', '%_Ngày', '%_Vol_vs_TB', 'Tín_Hiệu_Ngày']].sort_values(by='%_Vol_vs_TB', ascending=False)
    
    # Sheet 2
    df_tab2 = df_full.sort_values(by=['Ngành', 'Sức_Mạnh_Dòng_Tiền'], ascending=[True, False])
    
    # Sheet 3: Chấm điểm ngành
    df_tab3 = df_full.groupby('Ngành').agg({
        '%_Tăng_1_Tháng': 'mean',
        'Sức_Mạnh_Dòng_Tiền': 'mean',
        'GTGD_TB_Tỷ': 'sum',
        'Mã': 'count'
    }).reset_index()

    max_money = df_tab3['GTGD_TB_Tỷ'].max() or 1
    max_price = df_tab3['%_Tăng_1_Tháng'].max() or 1
    # Trọng số tiền 70%
    df_tab3['Điểm (0-100)'] = (
    (df_tab3['%_Tăng_1_Tháng'] / max_price * 30) + 
    (df_tab3['GTGD_TB_Tỷ'] / max_money * 70)
    ).round(1)
    df_tab3 = df_tab3.sort_values(by='Điểm (0-100)', ascending=False)

    file_name = "Taiwan_Market_Data_Latest.xlsx"
    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        df_tab1.to_excel(writer, sheet_name='1_Tin_Hieu_Hom_Nay', index=False)
        df_tab2.to_excel(writer, sheet_name='2_Xu_Huong_21_Ngay', index=False)
        df_tab3.to_excel(writer, sheet_name='3_Song_Nganh', index=False)
        
    print(f"\n✅ Đã xuất báo cáo: {file_name}")
    print("👉 Đã tách riêng nhóm Memory & Storage theo yêu cầu!")
else:
    print("❌ Không có dữ liệu.")
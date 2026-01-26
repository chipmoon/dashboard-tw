import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
import os

# --- 1. CẤU HÌNH DANH SÁCH MÃ MỞ RỘNG (40+ MÃ ĐẦU NGÀNH) ---
THONG_TIN_CO_PHIEU = {
    # 💾 NHÓM 1: MEMORY & STORAGE
    "8299.TWO": {"Ten": "Phison (Electronics)", "Ten_CN": "群聯", "Nganh": "Memory - Controller"},
    "2408.TW": {"Ten": "Nanya Technology", "Ten_CN": "南亞科", "Nganh": "Memory - DRAM"},
    "2344.TW": {"Ten": "Winbond Elec", "Ten_CN": "華邦電", "Nganh": "Memory - Flash/DRAM"},
    "2337.TW": {"Ten": "Macronix (MXIC)", "Ten_CN": "旺宏", "Nganh": "Memory - NOR Flash"},
    "3260.TWO": {"Ten": "ADATA", "Ten_CN": "威剛", "Nganh": "Memory - Module"},
    "2451.TW": {"Ten": "Transcend Info", "Ten_CN": "創見", "Nganh": "Memory - Module"},
    "4967.TW": {"Ten": "TeamGroup", "Ten_CN": "十銓", "Nganh": "Memory - Module"},
    "8150.TW": {"Ten": "ChipMOS", "Ten_CN": "南茂", "Nganh": "Memory - Packaging"},
    "6239.TW": {"Ten": "PTI (Powertech)", "Ten_CN": "力成", "Nganh": "Memory - Packaging"},

    # 🏭 NHÓM 2: FOUNDRY (SẢN XUẤT CHIP)
    "2330.TW": {"Ten": "TSMC", "Ten_CN": "台積電", "Nganh": "Foundry - Logic"},
    "2303.TW": {"Ten": "UMC", "Ten_CN": "聯電", "Nganh": "Foundry - Logic"},
    "6770.TW": {"Ten": "PSMC (Powerchip)", "Ten_CN": "力積電", "Nganh": "Foundry - Memory"},
    "5347.TWO": {"Ten": "VIS (Vanguard)", "Ten_CN": "世界先進", "Nganh": "Foundry - 8inch"},

    # 🧠 NHÓM 3: IC DESIGN & IP
    "2454.TW": {"Ten": "MediaTek", "Ten_CN": "聯發科", "Nganh": "IC Design - Mobile/AI"},
    "3034.TW": {"Ten": "Novatek", "Ten_CN": "聯詠", "Nganh": "IC Design - Display"},
    "2379.TW": {"Ten": "Realtek", "Ten_CN": "瑞昱", "Nganh": "IC Design - Network"},
    "5269.TW": {"Ten": "ASMedia", "Ten_CN": "祥碩", "Nganh": "IC Design - High Speed"},
    "3443.TW": {"Ten": "GUC (Global Unichip)", "Ten_CN": "創意", "Nganh": "Design Service (AI)"},
    "3661.TW": {"Ten": "Alchip", "Ten_CN": "世芯-KY", "Nganh": "Design Service (AI)"},
    "3035.TW": {"Ten": "Faraday Tech", "Ten_CN": "智原", "Nganh": "Design Service"},
    "8096.TWO": {"Ten": "CoAsia", "Ten_CN": "擎亞", "Nganh": "Design Service"},
    "3529.TWO": {"Ten": "eMemory", "Ten_CN": "力旺", "Nganh": "IP Core"},
    "6533.TW": {"Ten": "Andes Tech", "Ten_CN": "晶心科", "Nganh": "IP Core (RISC-V)"},

    # 📡 NHÓM 4: COMPOUND SEMI (5G/QUANG HỌC)
    "2455.TW": {"Ten": "Visual Photonics (VPEC)", "Ten_CN": "全新", "Nganh": "Compound Semi"},
    "3105.TWO": {"Ten": "Win Semi", "Ten_CN": "穩懋", "Nganh": "Compound Semi"},
    "8086.TWO": {"Ten": "AWSC", "Ten_CN": "宏捷科", "Nganh": "Compound Semi"},
    "3707.TW": {"Ten": "Epistar (Ennostar)", "Ten_CN": "富采", "Nganh": "Compound/LED"},

    # 📦 NHÓM 5: OSAT & EQUIPMENT (HẬU CẦN)
    "3711.TW": {"Ten": "ASE Tech", "Ten_CN": "日月光投控", "Nganh": "OSAT (Packaging)"},
    "2449.TW": {"Ten": "KYEC", "Ten_CN": "京元電子", "Nganh": "OSAT (Testing)"},
    "6488.TW": {"Ten": "GlobalWafers", "Ten_CN": "環球晶", "Nganh": "Wafer (Material)"},
    "5483.TWO": {"Ten": "Sino-American", "Ten_CN": "中美晶", "Nganh": "Wafer (Material)"},
    "3680.TW": {"Ten": "Gudeng", "Ten_CN": "家登", "Nganh": "Equipment (EUV Pod)"},

    # 🤖 NHÓM BỔ SUNG: AI SERVER & OEM
    "2317.TW": {"Ten": "Foxconn", "Ten_CN": "鴻海", "Nganh": "AI Server/OEM"},
    "3231.TW": {"Ten": "Wistron", "Ten_CN": "緯創", "Nganh": "AI Server/OEM"},
    "2382.TW": {"Ten": "Quanta", "Ten_CN": "廣達", "Nganh": "AI Server/OEM"},
    "2356.TW": {"Ten": "Inventec", "Ten_CN": "英業達", "Nganh": "AI Server/OEM"},
    "2301.TW": {"Ten": "Lite-On", "Ten_CN": "光寶科", "Nganh": "Power Supply"},
    "2308.TW": {"Ten": "Delta Elec", "Ten_CN": "台達電", "Nganh": "Power Supply"}
}

# --- 2. CẤU HÌNH MY FAVORITE (NHẬP MÃ BẠN SỞ HỮU TẠI ĐÂY) ---
MY_FAVORITES = ["2330", "2317", "2454", "3260", "8299"]

def get_quick_action(row):
    if row['%_Ngày'] > 1.8 and row['%_Vol_vs_TB'] > 150: return "🚀 MUA ĐUỔI"
    if row['Sức_Mạnh_Dòng_Tiền'] > 2.0: return "💰 TIỀN VÀO MẠNH"
    if row['%_Tăng_1_Tháng'] > 20 and row['%_Ngày'] < -1.5: return "⚠️ CHỐT LỜI BỚT"
    if row['%_Ngày'] < -3 and row['%_Vol_vs_TB'] > 130: return "❌ THOÁT HÀNG"
    return "👀 THEO DÕI"

# --- 3. QUÉT DỮ LIỆU ---
ket_qua = []
today = datetime.now()
start_date = today - timedelta(days=60)

for ticker, info in THONG_TIN_CO_PHIEU.items():
    try:
        data = yf.download(ticker, start=start_date, end=today, progress=False)
        if data.empty or len(data) < 22: continue
        
        gia_ht = data['Close'].iloc[-1]
        gia_truoc = data['Close'].iloc[-2]
        pct_ngay = ((gia_ht - gia_truoc) / gia_truoc) * 100
        vol_ht = data['Volume'].iloc[-1]
        vol_tb = data['Volume'].rolling(window=20).mean().iloc[-1]
        pct_vol = (vol_ht / vol_tb) * 100
        pct_1m = ((gia_ht - data['Close'].iloc[-21]) / data['Close'].iloc[-21]) * 100
        money_flow = (pct_vol / 100) * (1 + abs(pct_ngay) / 100)
        
        ket_qua.append({
            "Mã": ticker.split('.')[0],
            "Tên Công Ty (CN)": info.get('Ten_CN', info['Ten']),
            "Tên Công Ty (EN)": info['Ten'],
            "Ngành": info['Nganh'],
            "Giá": round(float(gia_ht), 2),
            "%_Ngày": round(float(pct_ngay), 2),
            "%_Vol_vs_TB": round(float(pct_vol), 0),
            "%_Tăng_1_Tháng": round(float(pct_1m), 2),
            "Sức_Mạnh_Dòng_Tiền": round(float(money_flow), 2),
            "Tín_Hiệu_Ngày": "Breakout" if (pct_ngay > 1 and pct_vol > 120) else "Tích lũy" if pct_ngay > 0 else "Yếu",
            "GTGD_TB_Tỷ": round((vol_tb * gia_ht) / 1e9, 3)
        })
    except: continue

# --- 4. XUẤT FILE 4 TABS ---
if ket_qua:
    df_full = pd.DataFrame(ket_qua)
    file_name = "Taiwan_Market_Data_Latest.xlsx"
    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        df_full[['Mã', 'Tên Công Ty (CN)', 'Giá', '%_Ngày', '%_Vol_vs_TB', 'Tín_Hiệu_Ngày']].to_excel(writer, sheet_name='1_Tin_Hieu_Hom_Nay', index=False)
        df_full[['Mã', 'Tên Công Ty (CN)', 'Ngành', '%_Tăng_1_Tháng', 'Sức_Mạnh_Dòng_Tiền']].to_excel(writer, sheet_name='2_Xu_Huong_21_Ngay', index=False)
        df_full.groupby('Ngành').agg({'%_Tăng_1_Tháng': 'mean', 'Sức_Mạnh_Dòng_Tiền': 'mean', 'GTGD_TB_Tỷ': 'sum', 'Mã': 'count'}).reset_index().to_excel(writer, sheet_name='3_Song_Nganh', index=False)
        
        df_fav = df_full[df_full['Mã'].isin(MY_FAVORITES)].copy()
        df_fav['QUICK_ACTION'] = df_fav.apply(get_quick_action, axis=1)
        df_fav[['Mã', 'Tên Công Ty (CN)', 'Giá', '%_Ngày', 'QUICK_ACTION']].to_excel(writer, sheet_name='4_My_Favorite', index=False)
    print(f"✅ Success! Saved {len(df_full)} stocks.")
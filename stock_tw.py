import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
import os
import logging
from pathlib import Path
import sys
import io

# --- SETUP LOGGING with UTF-8 encoding for Windows ---
# Fix encoding for Windows console
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('stock_tw_debug.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# --- 1. CẤU HÌNH DANH SÁCH MÃ MỞ RỘNG (40+ MÃ ĐẦU NGÀNH) ---
THONG_TIN_CO_PHIEU = {
    # 💾 NHÓM 1: MEMORY & STORAGE (BỘ NHỚ)
    "8299.TWO": {"Ten": "Phison (Electronics)", "Ten_CN": "群聯", "Nganh": "Memory - Controller"},
    "2408.TW": {"Ten": "Nanya Technology", "Ten_CN": "南亞科", "Nganh": "Memory - DRAM"},
    "2344.TW": {"Ten": "Winbond Elec", "Ten_CN": "華邦電", "Nganh": "Memory - Flash/DRAM"},
    "2337.TW": {"Ten": "Macronix (MXIC)", "Ten_CN": "旺宏", "Nganh": "Memory - NOR Flash"},
    "3260.TWO": {"Ten": "ADATA", "Ten_CN": "威剛", "Nganh": "Memory - Module"},
    "2451.TW": {"Ten": "Transcend Info", "Ten_CN": "創見", "Nganh": "Memory - Module"},
    "4967.TW": {"Ten": "TeamGroup", "Ten_CN": "十銓", "Nganh": "Memory - Module"},
    "8150.TW": {"Ten": "ChipMOS", "Ten_CN": "南茂", "Nganh": "Memory - Packaging"},
    "6239.TW": {"Ten": "PTI (Powertech)", "Ten_CN": "力成", "Nganh": "Memory - Packaging"},

    # 🏭 NHÓM 2: FOUNDRY & WAFERS (SẢN XUẤT CHIP & VẬT LIỆU)
    "2330.TW": {"Ten": "TSMC", "Ten_CN": "台積電", "Nganh": "Foundry - Logic"},
    "2303.TW": {"Ten": "UMC", "Ten_CN": "聯電", "Nganh": "Foundry - Logic"},
    "6770.TW": {"Ten": "PSMC (Powerchip)", "Ten_CN": "力積電", "Nganh": "Foundry - Memory"},
    "5347.TWO": {"Ten": "VIS (Vanguard)", "Ten_CN": "世界先進", "Nganh": "Foundry - 8inch"},
    "6488.TWO": {"Ten": "GlobalWafers", "Ten_CN": "環球晶", "Nganh": "Wafer - Material"},
    "5483.TWO": {"Ten": "Sino-American", "Ten_CN": "中美晶", "Nganh": "Wafer - Material"},

    # 🧠 NHÓM 3: IC DESIGN, IP & EQUIPMENT
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
    "3680.TW": {"Ten": "Gudeng", "Ten_CN": "家登", "Nganh": "Equipment (EUV Pod)"},
    "6133.TWO": {"Ten": "Gimhwak", "Ten_CN": "金橋", "Nganh": "Electronics"},
    "6173.TWO": {"Ten": "Shinmore", "Ten_CN": "信昌電", "Nganh": "Electronic Components"},
    # 📡 NHÓM 4: COMPOUND SEMI & OSAT (HẬU CẦN & 5G)
    "2455.TW": {"Ten": "Visual Photonics", "Ten_CN": "全新", "Nganh": "Compound Semi"},
    "3105.TWO": {"Ten": "Win Semi", "Ten_CN": "穩懋", "Nganh": "Compound Semi"},
    "8086.TWO": {"Ten": "AWSC", "Ten_CN": "宏捷科", "Nganh": "Compound Semi"},
    "3714.TW": {"Ten": "Ennostar Inc", "Ten_CN": "富采", "Nganh": "Compound/LED"},
    "3711.TW": {"Ten": "ASE Tech", "Ten_CN": "日月光投控", "Nganh": "OSAT (Packaging)"},
    "2449.TW": {"Ten": "KYEC", "Ten_CN": "京元電子", "Nganh": "OSAT (Testing)"},

    # 🤖 NHÓM 5: AI SERVER, OEM & POWER SUPPLY
    "2317.TW": {"Ten": "Foxconn", "Ten_CN": "鴻海", "Nganh": "AI Server/OEM"},
    "3231.TW": {"Ten": "Wistron", "Ten_CN": "緯創", "Nganh": "AI Server/OEM"},
    "2382.TW": {"Ten": "Quanta", "Ten_CN": "廣達", "Nganh": "AI Server/OEM"},
    "2356.TW": {"Ten": "Inventec", "Ten_CN": "英業達", "Nganh": "AI Server/OEM"},
    "2301.TW": {"Ten": "Lite-On", "Ten_CN": "光寶科", "Nganh": "Power Supply"},
    "2308.TW": {"Ten": "Delta Elec", "Ten_CN": "台達電", "Nganh": "Power Supply"},

    # 🚢 NHÓM 6: SHIPPING & LOGISTICS (VẬN TẢI BIỂN)
    "2603.TW": {"Ten": "Evergreen Marine", "Ten_CN": "長榮", "Nganh": "Shipping"},
    "2609.TW": {"Ten": "Yang Ming", "Ten_CN": "陽明", "Nganh": "Shipping"},
    "2615.TW": {"Ten": "Wan Hai Lines", "Ten_CN": "萬海", "Nganh": "Shipping"},
    "2618.TW": {"Ten": "EVA Air", "Ten_CN": "長榮航", "Nganh": "Airline"},
    "2610.TW": {"Ten": "China Airlines", "Ten_CN": "華航", "Nganh": "Airline"},

    # 💰 NHÓM 7: FINANCIALS (TÀI CHÍNH - TRỤ CỘT VN-INDEX ĐÀI LOAN)
    "2881.TW": {"Ten": "Fubon Financial", "Ten_CN": "富邦金", "Nganh": "Financial"},
    "2882.TW": {"Ten": "Cathay Financial", "Ten_CN": "國泰金", "Nganh": "Financial"},
    "2891.TW": {"Ten": "CTBC Financial", "Ten_CN": "中信金", "Nganh": "Financial"},
    "5880.TW": {"Ten": "TCB Financial", "Ten_CN": "合庫金", "Nganh": "Financial"},
    "2886.TW": {"Ten": "Mega Financial", "Ten_CN": "兆豐金", "Nganh": "Financial"},

    # 🏗️ NHÓM 8: TRADITIONAL INDUSTRY (NHỰA, THÉP, Ô TÔ)
    "1301.TW": {"Ten": "Formosa Plastics", "Ten_CN": "台塑", "Nganh": "Plastics"},
    "2002.TW": {"Ten": "China Steel", "Ten_CN": "中鋼", "Nganh": "Steel"},
    "2201.TW": {"Ten": "Yulon Motor", "Ten_CN": "裕隆", "Nganh": "Automobile"},
    "1526.TWO": {"Ten": "Kien Cheng", "Ten_CN": "建錩", "Nganh": "Industrial"},
}

# --- 2. CẤU HÌNH MY FAVORITE (NHẬP MÃ BẠN SỞ HỮU TẠI ĐÂY) ---
MY_FAVORITES = ["2454", "2317", "2455", "8299", "8096", "1526", "6133", "6173"]

# Validate that all favorites exist in THONG_TIN_CO_PHIEU
logger.info(f"🎯 MY_FAVORITES configured: {MY_FAVORITES}")
for fav_code in MY_FAVORITES:
    ticker_variants = [f"{fav_code}.TW", f"{fav_code}.TWO"]
    found_ticker = None
    for ticker in THONG_TIN_CO_PHIEU.keys():
        if ticker.replace(".TWO", "").replace(".TW", "") == fav_code:
            found_ticker = ticker
            company_name = THONG_TIN_CO_PHIEU[ticker]["Ten"]
            sector = THONG_TIN_CO_PHIEU[ticker]["Nganh"]
            logger.info(f"  ✓ {fav_code} → {ticker} ({company_name}, {sector})")
            break
    if not found_ticker:
        logger.error(f"  ✗ {fav_code}: NOT FOUND in THONG_TIN_CO_PHIEU dictionary!")

def get_quick_action(row):
    """🤖 AI Trading Signal Generator"""
    if row['%_Ngày'] > 1.8 and row['%_Vol_vs_TB'] > 150: return "🚀 MUA ĐUỔI"
    if row['Sức_Mạnh_Dòng_Tiền'] > 2.0: return "💰 TIỀN VÀO MẠNH"
    if row['%_Tăng_1_Tháng'] > 20 and row['%_Ngày'] < -1.5: return "⚠️ CHỐT LỜI BỚT"
    if row['%_Ngày'] < -3 and row['%_Vol_vs_TB'] > 130: return "❌ THOÁT HÀNG"
    return "👀 THEO DÕI"

def validate_price_data(data, ticker, min_rows=22, is_favorite=False):
    """✅ Validate data integrity before processing"""
    if data is None or data.empty:
        logger.warning(f"⚠️ {ticker}: Empty data returned")
        return False
    # For favorites, accept 10+ days; for others, require 22 days
    min_required = 10 if is_favorite else min_rows
    if len(data) < min_required:
        logger.warning(f"⚠️ {ticker}: Insufficient data ({len(data)}/{min_required})")
        return False
    return True

def safe_convert_to_float(value, default=0.0):
    """🔄 Safe conversion of Series/scalar to float"""
    try:
        if hasattr(value, 'item'):
            return float(value.item())
        elif isinstance(value, (int, float)):
            return float(value)
        else:
            return float(value)
    except (TypeError, ValueError, AttributeError):
        return default

def calculate_professional_indicators(df):
    """📊 Calculate professional-grade technical indicators"""
    indicators = {}
    close = df['Close']
    volume = df['Volume']
    
    try:
        # 1. RSI (Relative Strength Index) - Momentum (0-100 scale)
        delta = close.diff()
        gain = (delta.where(delta > 0, 0)).rolling(window=14).mean()
        loss = (-delta.where(delta < 0, 0)).rolling(window=14).mean()
        rs = gain / loss
        indicators['RSI'] = safe_convert_to_float((100 - (100 / (1 + rs))).iloc[-1], 50)
        
        # 2. MACD (Moving Average Convergence Divergence) - Trend
        ema_12 = close.ewm(span=12, adjust=False).mean()
        ema_26 = close.ewm(span=26, adjust=False).mean()
        indicators['MACD'] = safe_convert_to_float((ema_12 - ema_26).iloc[-1], 0)
        indicators['MACD_Signal'] = safe_convert_to_float((ema_12 - ema_26).ewm(span=9, adjust=False).mean().iloc[-1], 0)
        
        # 3. Bollinger Bands - Volatility (0-100 position)
        sma_20 = close.rolling(window=20).mean()
        std_20 = close.rolling(window=20).std()
        bb_upper = (sma_20 + 2 * std_20).iloc[-1]
        bb_lower = (sma_20 - 2 * std_20).iloc[-1]
        bb_middle = sma_20.iloc[-1]
        current_price = close.iloc[-1]
        bb_range = bb_upper - bb_lower
        indicators['BB_Position'] = ((current_price - bb_lower) / bb_range * 100) if bb_range > 0 else 50
        
        # 4. Stochastic Oscillator - Overbought/Oversold (0-100)
        low_14 = close.rolling(window=14).min()
        high_14 = close.rolling(window=14).max()
        k_percent = ((close - low_14) / (high_14 - low_14) * 100).iloc[-1]
        indicators['Stochastic'] = safe_convert_to_float(k_percent, 50)
        
        # 5. ATR (Average True Range) - Volatility in %
        high = df['High']
        low = df['Low']
        tr1 = high - low
        tr2 = abs(high - close.shift())
        tr3 = abs(low - close.shift())
        tr = pd.concat([tr1, tr2, tr3], axis=1).max(axis=1)
        atr_value = tr.rolling(window=14).mean().iloc[-1]
        indicators['ATR_Percent'] = (atr_value / current_price * 100) if current_price > 0 else 0
        
        # 6. Volume Trend - Increasing/Decreasing
        vol_sma = volume.rolling(window=20).mean().iloc[-1]
        current_vol = volume.iloc[-1]
        indicators['Vol_Trend'] = (current_vol / vol_sma - 1) * 100 if vol_sma > 0 else 0
        
    except Exception as e:
        logger.debug(f"Indicator calculation warning: {str(e)}")
        indicators = {'RSI': 50, 'MACD': 0, 'MACD_Signal': 0, 'BB_Position': 50, 'Stochastic': 50, 'ATR_Percent': 0, 'Vol_Trend': 0}
    
    return indicators

# --- 3. QUÉT DỮ LIỆU (BULK DOWNLOAD - More Reliable) ---
logger.info("🚀 STARTING TAIWAN STOCK ANALYSIS")
logger.info(f"📊 Total stocks to analyze: {len(THONG_TIN_CO_PHIEU)}")

DANH_SACH_MA = list(THONG_TIN_CO_PHIEU.keys())
today = datetime.now()
start_date = today - timedelta(days=60)

try:
    # Bulk download all tickers at once (more reliable)
    logger.info("📥 Bulk downloading all stocks...")
    data = yf.download(DANH_SACH_MA, start=start_date, end=today, progress=False, group_by='ticker', auto_adjust=True, threads=True)
    logger.info(f"✅ Downloaded {len(DANH_SACH_MA)} stocks")
except Exception as e:
    logger.error(f"❌ Download failed: {str(e)}")
    print(f"❌ Error downloading data: {str(e)}")
    exit()

ket_qua = []
success_count = 0
error_count = 0

for ticker in DANH_SACH_MA:
    try:
        logger.debug(f"📥 Processing {ticker}...")
        
        # Handle MultiIndex structure from bulk download
        if len(DANH_SACH_MA) == 1:
            df = data
        else:
            if ticker not in data.columns.levels[0]:
                logger.warning(f"⚠️ {ticker}: Not found in data")
                error_count += 1
                continue
            df = data[ticker].dropna()
        
        # Validate data (10 days for favorites, 22 for others)
        is_fav = ticker.replace('.TWO', '').replace('.TW', '') in MY_FAVORITES
        if not validate_price_data(df, ticker, min_rows=22, is_favorite=is_fav):
            error_count += 1
            continue
        
        # Get stock info
        info = THONG_TIN_CO_PHIEU.get(ticker, {"Ten": "Unknown", "Ten_CN": "", "Nganh": "Other"})
        
        # --- CALCULATE PROFESSIONAL INDICATORS ---
        pro_indicators = calculate_professional_indicators(df)
        
        # --- TÍNH TOÁN ---
        gia_hien_tai = df['Close'].iloc[-1]
        gia_hom_qua = df['Close'].iloc[-2]
        vol_hien_tai = df['Volume'].iloc[-1]
        
        pct_doi_ngay = (gia_hien_tai - gia_hom_qua) / gia_hom_qua * 100
        sma_20 = df['Close'].rolling(window=20).mean().iloc[-1]
        vol_tb_20 = df['Volume'].rolling(window=20).mean().iloc[-1]
        
        # 21-day trend
        gia_21_ngay_truoc = df['Close'].iloc[-21]
        pct_tang_1_thang = ((gia_hien_tai - gia_21_ngay_truoc) / gia_21_ngay_truoc) * 100
        
        vol_tb_5 = df['Volume'].rolling(window=5).mean().iloc[-1]
        suc_manh_dong_tien = (vol_tb_5 / vol_tb_20) if vol_tb_20 > 0 else 0
        gtgd_ty_twd = (gia_hien_tai * vol_tb_20) / 1_000_000_000 
        
        # Safe conversions
        gia_hien_tai_val = safe_convert_to_float(gia_hien_tai)
        pct_doi_ngay_val = safe_convert_to_float(pct_doi_ngay)
        pct_vol_val = safe_convert_to_float((vol_hien_tai / vol_tb_20 * 100) if vol_tb_20 > 0 else 0)
        pct_tang_1_thang_val = safe_convert_to_float(pct_tang_1_thang)
        suc_manh_dong_tien_val = safe_convert_to_float(suc_manh_dong_tien)
        gtgd_ty_twd_val = safe_convert_to_float(gtgd_ty_twd)
        
        # Signal determination
        tin_hieu_ngay = "Yếu"
        if gia_hien_tai > sma_20:
            if vol_hien_tai > vol_tb_20: 
                tin_hieu_ngay = "Bùng nổ (Breakout)"
            else: 
                tin_hieu_ngay = "Tích lũy (Up)"
        
        mã_code = ticker.replace(".TWO", "").replace(".TW", "")
        is_favorite = mã_code in MY_FAVORITES
        favorite_marker = "⭐" if is_favorite else "  "
        
        ket_qua.append({
            "Mã": mã_code,
            "Tên Công Ty": info['Ten'],
            "Tên Công Ty (CN)": info.get('Ten_CN', info['Ten']),
            "Ngành": info['Nganh'],
            "Giá": round(gia_hien_tai_val, 2),
            "%_Ngày": round(pct_doi_ngay_val, 2),
            "%_Vol_vs_TB": round(pct_vol_val, 0),
            "%_Tăng_1_Tháng": round(pct_tang_1_thang_val, 2),
            "Sức_Mạnh_Dòng_Tiền": round(suc_manh_dong_tien_val, 2),
            "Tín_Hiệu_Ngày": tin_hieu_ngay,
            "GTGD_TB_Tỷ": round(gtgd_ty_twd_val, 3),
            # Professional Indicators for Favorites
            "RSI": round(pro_indicators['RSI'], 2),
            "MACD": round(pro_indicators['MACD'], 4),
            "BB_Position": round(pro_indicators['BB_Position'], 1),
            "Stochastic": round(pro_indicators['Stochastic'], 1),
            "ATR%": round(pro_indicators['ATR_Percent'], 2),
            "Vol_Trend": round(pro_indicators['Vol_Trend'], 1),
        })
        success_count += 1
        logger.debug(f"✅ {favorite_marker} {ticker} ({mã_code}): {info['Ten']} - Price: {gia_hien_tai_val:.2f} TWD")
        
    except Exception as e:
        error_count += 1
        logger.error(f"❌ Error processing {ticker}: {str(e)}")
        continue

logger.info(f"✅ Data collection completed: {success_count} success, {error_count} errors")

# --- DIAGNOSE FAVORITE STOCKS ---
logger.info("📍 Checking favorite stocks collection...")
collected_mã = set(item["Mã"] for item in ket_qua)
collected_favorites = [fav for fav in MY_FAVORITES if fav in collected_mã]
missing_favorites = [fav for fav in MY_FAVORITES if fav not in collected_mã]

logger.info(f"📊 Favorites collected in ket_qua: {len(collected_favorites)}/{len(MY_FAVORITES)}")
for fav in collected_favorites:
    # Find details from ket_qua
    fav_data = next((item for item in ket_qua if item["Mã"] == fav), None)
    if fav_data:
        logger.info(f"  ✓ {fav}: {fav_data.get('Tên Công Ty', 'N/A')} - Price: {fav_data.get('Giá', 'N/A')} TWD")

# --- ENHANCED ROOT CAUSE ANALYSIS ---
logger.info("\n📊 ENHANCED ROOT CAUSE ANALYSIS:")
logger.info(f"Total in ket_qua: {len(ket_qua)} stocks")
logger.info(f"Collected Mã codes: {sorted(collected_mã)}")
logger.info(f"MY_FAVORITES: {MY_FAVORITES}")

if missing_favorites:
    logger.warning(f"\n⚠️ MISSING FROM ket_qua: {missing_favorites}")
    for fav in missing_favorites:
        # Find the full ticker code
        full_ticker = None
        for ticker in DANH_SACH_MA:
            if ticker.replace(".TWO", "").replace(".TW", "") == fav:
                full_ticker = ticker
                break
        company_info = THONG_TIN_CO_PHIEU.get(full_ticker, {})
        company_name = company_info.get("Ten", "Unknown")
        sector = company_info.get("Nganh", "Unknown")
        
        # Check if it's in the full dataframe BEFORE filtering
        logger.warning(f"\n  Stock: {fav} ({full_ticker}) - {company_name} [{sector}]")
        
        # This will be checked after df_full is created
        logger.warning(f"    → Check Sheet 1 & 2 for presence (will verify below)")
else:
    logger.info(f"✅ All {len(MY_FAVORITES)} favorite stocks collected successfully!")

# --- 4. XUẤT FILE 4 TABS ---
if ket_qua:
    logger.info(f"📊 Creating Excel report with {len(ket_qua)} stocks...")
    df_full = pd.DataFrame(ket_qua)
    
    # Sort by different criteria for each sheet
    df_tab1 = df_full.sort_values(by='%_Vol_vs_TB', ascending=False)
    df_tab2 = df_full.sort_values(by=['%_Tăng_1_Tháng'], ascending=False)
    
    # --- ENHANCED: Check if "missing" favorites are actually in df_full ---
    logger.info("\n" + "="*70)
    logger.info("🔍 CHECKING IF 'MISSING' FAVORITES ARE IN df_full (SHEETS 1 & 2):")
    logger.info("="*70)
    
    collected_mã_full = set(df_full['Mã'].values)
    logger.info(f"\nTotal in df_full: {len(df_full)} stocks")
    logger.info(f"Mã codes in df_full: {sorted(collected_mã_full)}")
    
    # Re-check favorites against df_full
    truly_missing_from_df = []
    present_in_sheets_but_missing_from_favorites = []
    
    for fav in MY_FAVORITES:
        if fav in collected_mã_full:
            fav_row = df_full[df_full['Mã'] == fav].iloc[0]
            logger.info(f"  ✅ {fav}: {fav_row['Tên Công Ty']} - Price: {fav_row['Giá']} TWD")
            logger.info(f"     → FOUND in Sheet 1 & 2 (Sheets show all 40 stocks)")
        else:
            truly_missing_from_df.append(fav)
            logger.warning(f"  ❌ {fav}: NOT in df_full (not even in Sheets 1 & 2)")
    
    # Check if any are in df_full but somehow filtered from Sheet 4
    if not truly_missing_from_df:
        logger.info(f"\n✅ ALL {len(MY_FAVORITES)} FAVORITES ARE IN df_full!")
        logger.info("   → If Sheet 4 is empty, there's a FILTERING ISSUE, not data issue")
    else:
        logger.warning(f"\n⚠️ {len(truly_missing_from_df)} truly missing from df_full:")
        for fav in truly_missing_from_df:
            logger.warning(f"   - {fav}")
    
    logger.info("="*70 + "\n")
    file_name = "Taiwan_Market_Data_Latest.xlsx"
    
    try:
        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            # Sheet 1: Daily Signals (sorted by volume strength)
            df_tab1[['Mã', 'Tên Công Ty (CN)', 'Tên Công Ty', 'Giá', '%_Ngày', '%_Vol_vs_TB', 'Tín_Hiệu_Ngày', 'GTGD_TB_Tỷ']].to_excel(
                writer, sheet_name='1_Tin_Hieu_Hom_Nay', index=False
            )
            logger.debug("✅ Sheet 1 created: 1_Tin_Hieu_Hom_Nay")
            
            # Sheet 2: 21-day Trend (sorted by 1-month gain)
            df_tab2[['Mã', 'Tên Công Ty (CN)', 'Tên Công Ty', 'Ngành', '%_Tăng_1_Tháng', 'Sức_Mạnh_Dòng_Tiền', 'GTGD_TB_Tỷ']].to_excel(
                writer, sheet_name='2_Xu_Huong_21_Ngay', index=False
            )
            logger.debug("✅ Sheet 2 created: 2_Xu_Huong_21_Ngay")
            
            # Sheet 3: Sector Analysis
            df_sector = df_full.groupby('Ngành').agg({
                '%_Tăng_1_Tháng': 'mean', 
                'Sức_Mạnh_Dòng_Tiền': 'mean', 
                'GTGD_TB_Tỷ': 'sum', 
                'Mã': 'count'
            }).reset_index()
            df_sector.columns = ['Ngành', 'Avg_%_1Tháng', 'Avg_Sức_Mạnh', 'GTGD_TB_Tỷ', 'Số_Mã']
            df_sector = df_sector.sort_values(by='Avg_%_1Tháng', ascending=False)
            df_sector.to_excel(
                writer, sheet_name='3_Song_Nganh', index=False
            )
            logger.debug("✅ Sheet 3 created: 3_Song_Nganh")
            
            # Sheet 4: My Favorite Stocks with Trading Signals
            logger.info("\n" + "="*70)
            logger.info("🎯 SHEET 4 - MY FAVORITE STOCKS FILTERING ANALYSIS:")
            logger.info("="*70)
            
            logger.info(f"MY_FAVORITES: {MY_FAVORITES}")
            logger.info(f"df_full['Mã'] unique values: {sorted(df_full['Mã'].unique())}")
            
            # Debug: Check each favorite
            favorites_in_df = []
            favorites_not_in_df = []
            for fav in MY_FAVORITES:
                matches = df_full[df_full['Mã'] == fav]
                if len(matches) > 0:
                    favorites_in_df.append(fav)
                    logger.info(f"  ✅ {fav}: Found in df_full ({len(matches)} row(s))")
                else:
                    favorites_not_in_df.append(fav)
                    logger.warning(f"  ❌ {fav}: NOT found in df_full")
            
            df_fav = df_full[df_full['Mã'].isin(MY_FAVORITES)].copy()
            logger.info(f"\nFiltering result: {len(df_fav)} rows selected from {len(df_full)} total")
            logger.info(f"Selected favorites: {sorted(df_fav['Mã'].values)}")
            
            if not df_fav.empty:
                df_fav['QUICK_ACTION'] = df_fav.apply(get_quick_action, axis=1)
                # Professional columns for favorites: Key indicators ranked by importance
                fav_columns = [
                    'Mã', 'Tên Công Ty (CN)', 'Tên Công Ty',
                    'Giá', '%_Ngày', '%_Tăng_1_Tháng',
                    'RSI', 'MACD', 'BB_Position', 'Stochastic',
                    'ATR%', 'Vol_Trend', 'Sức_Mạnh_Dòng_Tiền',
                    'QUICK_ACTION'
                ]
                df_fav[fav_columns].to_excel(
                    writer, sheet_name='4_My_Favorite', index=False
                )
                fav_count = len(df_fav)
                logger.info(f"\n✅ Sheet 4 created: 4_My_Favorite ({fav_count}/{len(MY_FAVORITES)} favorites)")
                logger.info(f"   Columns: Basic Info + 6 Professional Indicators")
                if fav_count < len(MY_FAVORITES):
                    missing = [fav for fav in MY_FAVORITES if fav not in df_fav['Mã'].values]
                    logger.warning(f"⚠️ Missing in Sheet 4: {missing}")
                    logger.info(f"   → Now accepting 10+ days for favorites (was 22)")
                logger.info("="*70 + "\n")
            else:
                logger.warning("⚠️ No favorite stocks found in data")
                logger.info("="*70 + "\n")
        
        logger.info(f"✅✅✅ SUCCESS! File saved: {file_name}")
        logger.info(f"📈 Report contains {len(df_full)} stocks across {len(df_full['Ngành'].unique())} sectors")
        print(f"\n{'='*60}")
        print(f"✅✅✅ SUCCESS! Saved {len(df_full)} stocks to {file_name}")
        print(f"📊 Sectors analyzed: {len(df_full['Ngành'].unique())}")
        print(f"⭐ Favorites tracked: {len(df_fav)}")
        print(f"{'='*60}\n")
        
    except Exception as e:
        logger.error(f"❌ FAILED to write Excel file: {str(e)}")
        print(f"\n❌ ERROR saving file: {str(e)}")
        print("Check 'stock_tw_debug.log' for details")
else:
    logger.error("❌ NO DATA COLLECTED - Empty result list")
    print("\n" + "="*60)
    print("❌ NO DATA COLLECTED")
    print("Possible causes:")
    print("  • Network connection issue")
    print("  • Yahoo Finance API rate limit")
    print("  • All stocks failed validation")
    print("Check 'stock_tw_debug.log' for details")
    print("="*60 + "\n")

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

# --- 1. EXTENDED STOCK LIST CONFIGURATION (40+ LEADING STOCKS) ---
STOCK_INFO = {
    # 💾 GROUP 1: MEMORY & STORAGE
    "8299.TWO": {"Name": "Phison (Electronics)", "Name_CN": "群聯", "Sector": "Memory - Controller"},
    "2408.TW": {"Name": "Nanya Technology", "Name_CN": "南亞科", "Sector": "Memory - DRAM"},
    "2344.TW": {"Name": "Winbond Elec", "Name_CN": "華邦電", "Sector": "Memory - Flash/DRAM"},
    "2337.TW": {"Name": "Macronix (MXIC)", "Name_CN": "旺宏", "Sector": "Memory - NOR Flash"},
    "3260.TWO": {"Name": "ADATA", "Name_CN": "威剛", "Sector": "Memory - Module"},
    "2451.TW": {"Name": "Transcend Info", "Name_CN": "創見", "Sector": "Memory - Module"},
    "4967.TW": {"Name": "TeamGroup", "Name_CN": "十銓", "Sector": "Memory - Module"},
    "8150.TW": {"Name": "ChipMOS", "Name_CN": "南茂", "Sector": "Memory - Packaging"},
    "6239.TW": {"Name": "PTI (Powertech)", "Name_CN": "力成", "Sector": "Memory - Packaging"},

    # 🏭 GROUP 2: FOUNDRY & WAFERS (CHIP PRODUCTION & MATERIALS)
    "2330.TW": {"Name": "TSMC", "Name_CN": "台積電", "Sector": "Foundry - Logic"},
    "2303.TW": {"Name": "UMC", "Name_CN": "聯電", "Sector": "Foundry - Logic"},
    "6770.TW": {"Name": "PSMC (Powerchip)", "Name_CN": "力積電", "Sector": "Foundry - Memory"},
    "5347.TWO": {"Name": "VIS (Vanguard)", "Name_CN": "世界先進", "Sector": "Foundry - 8inch"},
    "6488.TWO": {"Name": "GlobalWafers", "Name_CN": "環球晶", "Sector": "Wafer - Material"},
    "5483.TWO": {"Name": "Sino-American", "Name_CN": "中美晶", "Sector": "Wafer - Material"},

    # 🧠 GROUP 3: IC DESIGN, IP & EQUIPMENT
    "2454.TW": {"Name": "MediaTek", "Name_CN": "聯發科", "Sector": "IC Design - Mobile/AI"},
    "3034.TW": {"Name": "Novatek", "Name_CN": "聯詠", "Sector": "IC Design - Display"},
    "2379.TW": {"Name": "Realtek", "Name_CN": "瑞昱", "Sector": "IC Design - Network"},
    "5269.TW": {"Name": "ASMedia", "Name_CN": "祥碩", "Sector": "IC Design - High Speed"},
    "3443.TW": {"Name": "GUC (Global Unichip)", "Name_CN": "創意", "Sector": "Design Service (AI)"},
    "3661.TW": {"Name": "Alchip", "Name_CN": "世芯-KY", "Sector": "Design Service (AI)"},
    "3035.TW": {"Name": "Faraday Tech", "Name_CN": "智原", "Sector": "Design Service"},
    "8096.TWO": {"Name": "CoAsia", "Name_CN": "擎亞", "Sector": "Design Service"},
    "3529.TWO": {"Name": "eMemory", "Name_CN": "力旺", "Sector": "IP Core"},
    "6533.TW": {"Name": "Andes Tech", "Name_CN": "晶心科", "Sector": "IP Core (RISC-V)"},
    "3680.TW": {"Name": "Gudeng", "Name_CN": "家登", "Sector": "Equipment (EUV Pod)"},
    "6133.TWO": {"Name": "Gimhwak", "Name_CN": "金橋", "Sector": "Electronics"},
    "6173.TWO": {"Name": "Shinmore", "Name_CN": "信昌電", "Sector": "Electronic Components"},
    # 📡 GROUP 4: COMPOUND SEMI & OSAT (BACKEND & 5G)
    "2455.TW": {"Name": "Visual Photonics", "Name_CN": "全新", "Sector": "Compound Semi"},
    "3105.TWO": {"Name": "Win Semi", "Name_CN": "穩懋", "Sector": "Compound Semi"},
    "8086.TWO": {"Name": "AWSC", "Name_CN": "宏捷科", "Sector": "Compound Semi"},
    "3714.TW": {"Name": "Ennostar Inc", "Name_CN": "富采", "Sector": "Compound/LED"},
    "3711.TW": {"Name": "ASE Tech", "Name_CN": "日月光投控", "Sector": "OSAT (Packaging)"},
    "2449.TW": {"Name": "KYEC", "Name_CN": "京元電子", "Sector": "OSAT (Testing)"},

    # 🤖 GROUP 5: AI SERVER, OEM & POWER SUPPLY
    "2317.TW": {"Name": "Foxconn", "Name_CN": "鴻海", "Sector": "AI Server/OEM"},
    "3231.TW": {"Name": "Wistron", "Name_CN": "緯創", "Sector": "AI Server/OEM"},
    "2382.TW": {"Name": "Quanta", "Name_CN": "廣達", "Sector": "AI Server/OEM"},
    "2356.TW": {"Name": "Inventec", "Name_CN": "英業達", "Sector": "AI Server/OEM"},
    "2301.TW": {"Name": "Lite-On", "Name_CN": "光寶科", "Sector": "Power Supply"},
    "2308.TW": {"Name": "Delta Elec", "Name_CN": "台達電", "Sector": "Power Supply"},

    # 🚢 GROUP 6: SHIPPING & LOGISTICS (MARITIME TRANSPORT)
    "2603.TW": {"Name": "Evergreen Marine", "Name_CN": "長榮", "Sector": "Shipping"},
    "2609.TW": {"Name": "Yang Ming", "Name_CN": "陽明", "Sector": "Shipping"},
    "2615.TW": {"Name": "Wan Hai Lines", "Name_CN": "萬海", "Sector": "Shipping"},
    "2618.TW": {"Name": "EVA Air", "Name_CN": "長榮航", "Sector": "Airline"},
    "2610.TW": {"Name": "China Airlines", "Name_CN": "華航", "Sector": "Airline"},

    # 💰 GROUP 7: FINANCIALS (FINANCIAL PILLARS)
    "2881.TW": {"Name": "Fubon Financial", "Name_CN": "富邦金", "Sector": "Financial"},
    "2882.TW": {"Name": "Cathay Financial", "Name_CN": "國泰金", "Sector": "Financial"},
    "2891.TW": {"Name": "CTBC Financial", "Name_CN": "中信金", "Sector": "Financial"},
    "5880.TW": {"Name": "TCB Financial", "Name_CN": "合庫金", "Sector": "Financial"},
    "2886.TW": {"Name": "Mega Financial", "Name_CN": "兆豐金", "Sector": "Financial"},

    # 🏗️ GROUP 8: TRADITIONAL INDUSTRY (PLASTICS, STEEL, AUTO)
    "1301.TW": {"Name": "Formosa Plastics", "Name_CN": "台塑", "Sector": "Plastics"},
    "2002.TW": {"Name": "China Steel", "Name_CN": "中鋼", "Sector": "Steel"},
    "2201.TW": {"Name": "Yulon Motor", "Name_CN": "裕隆", "Sector": "Automobile"},
    "1526.TWO": {"Name": "Kien Cheng", "Name_CN": "建錩", "Sector": "Industrial"},
}

# --- 2. MY FAVORITE CONFIGURATION (ENTER YOUR PORTFOLIO CODES HERE) ---
MY_FAVORITES = ["2454", "2317", "2455", "8299", "8096", "1526", "6133", "6173"]

# Validate that all favorites exist in STOCK_INFO
logger.info(f"🎯 MY_FAVORITES configured: {MY_FAVORITES}")
for fav_code in MY_FAVORITES:
    ticker_variants = [f"{fav_code}.TW", f"{fav_code}.TWO"]
    found_ticker = None
    for ticker in STOCK_INFO.keys():
        if ticker.replace(".TWO", "").replace(".TW", "") == fav_code:
            found_ticker = ticker
            company_name = STOCK_INFO[ticker]["Name"]
            sector = STOCK_INFO[ticker]["Sector"]
            logger.info(f"  ✓ {fav_code} → {ticker} ({company_name}, {sector})")
            break
    if not found_ticker:
        logger.error(f"  ✗ {fav_code}: NOT FOUND in STOCK_INFO dictionary!")

def get_quick_action(row):
    """🤖 AI Trading Signal Generator"""
    if row['%_Ngày'] > 1.8 and row['%_Vol_vs_TB'] > 150: return "🚀 BUY STRONG"
    if row['Sức_Mạnh_Dòng_Tiền'] > 2.0: return "💰 STRONG INFLOW"
    if row['%_Tăng_1_Tháng'] > 20 and row['%_Ngày'] < -1.5: return "⚠️ TAKE PROFIT"
    if row['%_Ngày'] < -3 and row['%_Vol_vs_TB'] > 130: return "❌ EXIT"
    return "👀 WATCH"

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

# --- 3. DATA SCANNING (BULK DOWNLOAD - More Reliable) ---
logger.info("🚀 STARTING TAIWAN STOCK ANALYSIS")
logger.info(f"📊 Total stocks to analyze: {len(STOCK_INFO)}")

TICKER_LIST = list(STOCK_INFO.keys())
today = datetime.now()
start_date = today - timedelta(days=60)

try:
    # Bulk download all tickers at once (more reliable)
    logger.info("📥 Bulk downloading all stocks...")
    data = yf.download(TICKER_LIST, start=start_date, end=today, progress=False, group_by='ticker', auto_adjust=True, threads=True)
    logger.info(f"✅ Downloaded {len(TICKER_LIST)} stocks")
except Exception as e:
    logger.error(f"❌ Download failed: {str(e)}")
    print(f"❌ Error downloading data: {str(e)}")
    exit()

results = []
success_count = 0
error_count = 0

for ticker in TICKER_LIST:
    try:
        logger.debug(f"📥 Processing {ticker}...")
        
        # Handle MultiIndex structure from bulk download
        if len(TICKER_LIST) == 1:
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
        info = STOCK_INFO.get(ticker, {"Name": "Unknown", "Name_CN": "", "Sector": "Other"})
        
        # --- CALCULATE PROFESSIONAL INDICATORS ---
        pro_indicators = calculate_professional_indicators(df)
        
        # --- CALCULATIONS ---
        current_price_val = df['Close'].iloc[-1]
        prev_price = df['Close'].iloc[-2]
        current_vol = df['Volume'].iloc[-1]
        
        pct_change_day = (current_price_val - prev_price) / prev_price * 100
        sma_20 = df['Close'].rolling(window=20).mean().iloc[-1]
        vol_tb_20 = df['Volume'].rolling(window=20).mean().iloc[-1]
        
        # 21-day trend
        price_21d_ago = df['Close'].iloc[-21]
        pct_gain_1m = ((current_price_val - price_21d_ago) / price_21d_ago) * 100
        
        vol_tb_5 = df['Volume'].rolling(window=5).mean().iloc[-1]
        money_flow_strength = (vol_tb_5 / vol_tb_20) if vol_tb_20 > 0 else 0
        avg_trading_val_b = (current_price_val * vol_tb_20) / 1_000_000_000 
        
        # Safe conversions
        current_price_val_val = safe_convert_to_float(current_price_val)
        pct_change_day_val = safe_convert_to_float(pct_change_day)
        pct_vol_val = safe_convert_to_float((current_vol / vol_tb_20 * 100) if vol_tb_20 > 0 else 0)
        pct_gain_1m_val = safe_convert_to_float(pct_gain_1m)
        money_flow_strength_val = safe_convert_to_float(money_flow_strength)
        avg_trading_val_b_val = safe_convert_to_float(avg_trading_val_b)
        
        # Signal determination
        signal_day = "Weak"
        if current_price_val > sma_20:
            if current_vol > vol_tb_20: 
                signal_day = "Breakout"
            else: 
                signal_day = "Accumulation (Up)"
        
        stock_code = ticker.replace(".TWO", "").replace(".TW", "")
        is_favorite = stock_code in MY_FAVORITES
        favorite_marker = "⭐" if is_favorite else "  "
        
        results.append({
            "Code": stock_code,
            "Name": info['Name'],
            "Name_CN": info.get('Name_CN', info['Name']),
            "Sector": info['Sector'],
            "Price": round(current_price_val_val, 2),
            "Pct_Day": round(pct_change_day_val, 2),
            "Vol_vs_Avg": round(pct_vol_val, 0),
            "Pct_1Month": round(pct_gain_1m_val, 2),
            "Money_Flow_Strength": round(money_flow_strength_val, 2),
            "Signal": signal_day,
            "Avg_Trading_Value_B": round(avg_trading_val_b_val, 3),
            # Professional Indicators for Favorites
            "RSI": round(pro_indicators['RSI'], 2),
            "MACD": round(pro_indicators['MACD'], 4),
            "BB_Position": round(pro_indicators['BB_Position'], 1),
            "Stochastic": round(pro_indicators['Stochastic'], 1),
            "ATR_Pct": round(pro_indicators['ATR_Percent'], 2),
            "Vol_Trend": round(pro_indicators['Vol_Trend'], 1),
        })
        success_count += 1
        logger.debug(f"✅ {favorite_marker} {ticker} ({stock_code}): {info['Name']} - Price: {current_price_val_val:.2f} TWD")
        
    except Exception as e:
        error_count += 1
        logger.error(f"❌ Error processing {ticker}: {str(e)}")
        continue

logger.info(f"✅ Data collection completed: {success_count} success, {error_count} errors")

# --- DIAGNOSE FAVORITE STOCKS ---
logger.info("📍 Checking favorite stocks collection...")
collected_codes = set(item["Code"] for item in results)
collected_favorites = [fav for fav in MY_FAVORITES if fav in collected_codes]
missing_favorites = [fav for fav in MY_FAVORITES if fav not in collected_codes]

logger.info(f"📊 Favorites collected in results: {len(collected_favorites)}/{len(MY_FAVORITES)}")
for fav in collected_favorites:
    # Find details from results
    favorite_data = next((item for item in results if item["Code"] == fav), None)
    if favorite_data:
        logger.info(f"  ✓ {fav}: {favorite_data.get('Name', 'N/A')} - Price: {favorite_data.get('Price', 'N/A')} TWD")

# --- ENHANCED ROOT CAUSE ANALYSIS ---
logger.info("\n📊 ENHANCED ROOT CAUSE ANALYSIS:")
logger.info(f"Total in results: {len(results)} stocks")
logger.info(f"Collected Code codes: {sorted(collected_codes)}")
logger.info(f"MY_FAVORITES: {MY_FAVORITES}")

if missing_favorites:
    logger.warning(f"\n⚠️ MISSING FROM results: {missing_favorites}")
    for fav in missing_favorites:
        # Find the full ticker code
        full_ticker = None
        for ticker in TICKER_LIST:
            if ticker.replace(".TWO", "").replace(".TW", "") == fav:
                full_ticker = ticker
                break
        company_info = STOCK_INFO.get(full_ticker, {})
        company_name = company_info.get("Name", "Unknown")
        sector = company_info.get("Sector", "Unknown")
        
        # Check if it's in the full dataframe BEFORE filtering
        logger.warning(f"\n  Stock: {fav} ({full_ticker}) - {company_name} [{sector}]")
        
        # This will be checked after df_full is created
        logger.warning(f"    → Check Sheet 1 & 2 for presence (will verify below)")
else:
    logger.info(f"✅ All {len(MY_FAVORITES)} favorite stocks collected successfully!")

# --- 4. EXPORT FILE WITH 4 TABS ---
if results:
    logger.info(f"📊 Creating Excel report with {len(results)} stocks...")
    df_full = pd.DataFrame(results)
    
    # Rename columns to match Vietnamese names used throughout the code
    df_full = df_full.rename(columns={
        'Code': 'Mã',
        'Name': 'Tên Công Ty',
        'Name_CN': 'Tên Công Ty (CN)',
        'Price': 'Giá',
        'Pct_Day': '%_Ngày',
        'Vol_vs_Avg': '%_Vol_vs_TB',
        'Pct_1Month': '%_Tăng_1_Tháng',
        'Signal': 'Tín_Hiệu_Ngày',
        'Avg_Trading_Value_B': 'GTGD_TB_Tỷ',
        'Money_Flow_Strength': 'Sức_Mạnh_Dòng_Tiền',
        'ATR_Pct': 'ATR%'
    })
    
    # Sort by different criteria for each sheet
    df_tab1 = df_full.sort_values(by='%_Vol_vs_TB', ascending=False)
    df_tab2 = df_full.sort_values(by=['%_Tăng_1_Tháng'], ascending=False)
    
    # --- ENHANCED: Check if "missing" favorites are actually in df_full ---
    logger.info("\n" + "="*70)
    logger.info("🔍 CHECKING IF 'MISSING' FAVORITES ARE IN df_full (SHEETS 1 & 2):")
    logger.info("="*70)
    
    collected_codes_full = set(df_full['Mã'].values)
    logger.info(f"\nTotal in df_full: {len(df_full)} stocks")
    logger.info(f"Mã codes in df_full: {sorted(collected_codes_full)}")
    
    # Re-check favorites against df_full
    truly_missing_from_df = []
    present_in_sheets_but_missing_from_favorites = []
    
    for fav in MY_FAVORITES:
        if fav in collected_codes_full:
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
                writer, sheet_name='1_Daily_Signals', index=False
            )
            logger.debug("✅ Sheet 1 created: 1_Tin_Hieu_Hom_Nay")
            
            # Sheet 2: 21-day Trend (sorted by 1-month gain)
            df_tab2[['Mã', 'Tên Công Ty (CN)', 'Tên Công Ty', 'Sector', '%_Tăng_1_Tháng', 'Sức_Mạnh_Dòng_Tiền', 'GTGD_TB_Tỷ']].to_excel(
                writer, sheet_name='2_21Day_Trend', index=False
            )
            logger.debug("✅ Sheet 2 created: 2_Xu_Huong_21_Ngay")
            
            # Sheet 3: Sector Analysis
            df_sector = df_full.groupby('Sector').agg({
                '%_Tăng_1_Tháng': 'mean', 
                'Sức_Mạnh_Dòng_Tiền': 'mean', 
                'GTGD_TB_Tỷ': 'sum', 
                'Mã': 'count'
            }).reset_index()
            df_sector.columns = ['Sector', 'Avg_Pct_1M', 'Avg_Money_Flow', 'GTGD_TB_Tỷ', 'Stock_Count']
            df_sector = df_sector.sort_values(by='Avg_Pct_1M', ascending=False)
            df_sector.to_excel(
                writer, sheet_name='3_Industry_Analysis', index=False
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
                    writer, sheet_name='4_My_Favorites', index=False
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
        logger.info(f"📈 Report contains {len(df_full)} stocks across {len(df_full['Sector'].unique())} sectors")
        print(f"\n{'='*60}")
        print(f"✅✅✅ SUCCESS! Saved {len(df_full)} stocks to {file_name}")
        print(f"📊 Sectors analyzed: {len(df_full['Sector'].unique())}")
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

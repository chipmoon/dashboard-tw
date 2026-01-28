import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

# ============================================================================
# 1. PAGE CONFIGURATION
# ============================================================================
st.set_page_config(layout="wide", page_title="Taiwan Market Dashboard 🇹🇼", page_icon="💰")

st.title("💰 TAIWAN MARKET DASHBOARD - SMART MONEY FLOW (40+ Stocks)")
st.markdown("### Professional Industry Analysis | Technical Indicators | Comprehensive Coverage")

# ============================================================================
# 2. LOAD DATA
# ============================================================================
current_folder = os.path.dirname(os.path.abspath(__file__))
target_file = os.path.join(current_folder, "Taiwan_Market_Data_Latest.xlsx")

if not os.path.exists(target_file):
    st.error(f"❌ Data file not found: {target_file}")
    st.info("⚠️ Please run **stock_tw_fixed.py** first to generate the data file.")
    st.stop()

@st.cache_data
def load_data():
    """Load all sheets from Excel file"""
    df_daily = pd.read_excel(target_file, sheet_name='1_Daily_Signals')
    df_trend = pd.read_excel(target_file, sheet_name='2_21Day_Trend')
    df_sector = pd.read_excel(target_file, sheet_name='3_Industry_Analysis')
    try:
        df_favorite = pd.read_excel(target_file, sheet_name='4_My_Favorites')
    except:
        df_favorite = pd.DataFrame()
    # --- Normalize column names: map Vietnamese Excel columns to English names used in the dashboard ---
    viet_to_eng = {
        'Mã': 'Code',
        'Tên Công Ty': 'Name',
        'Tên Công Ty (CN)': 'Name_CN',
        'Giá': 'Price',
        '%_Ngày': 'Pct_Day',
        '%_Vol_vs_TB': 'Vol_vs_Avg',
        '%_Tăng_1_Tháng': 'Pct_1Month',
        'Sức_Mạnh_Dòng_Tiền': 'Money_Flow_Strength',
        'GTGD_TB_Tỷ': 'Avg_Trading_Value_B',
        'ATR%': 'ATR_Pct',
        'Tín_Hiệu_Ngày': 'Signal',
        'Sector': 'Sector'
    }

    def _rename_if_needed(df):
        if df is None or df.empty:
            return df
        rename_map = {k: v for k, v in viet_to_eng.items() if k in df.columns}
        if rename_map:
            return df.rename(columns=rename_map)
        return df

    df_daily = _rename_if_needed(df_daily)
    df_trend = _rename_if_needed(df_trend)
    df_favorite = _rename_if_needed(df_favorite)

    # df_sector uses different column names in the generator - ensure fallback keys exist
    if df_sector is not None and not df_sector.empty:
        # Map trading-value column and keep Avg_Pct_1M as-is if present
        sector_rename = {}
        if 'GTGD_TB_Tỷ' in df_sector.columns:
            sector_rename['GTGD_TB_Tỷ'] = 'Total_Trading_Value_B'
        if 'Avg_Pct_1M' in df_sector.columns:
            sector_rename['Avg_Pct_1M'] = 'Avg_Pct_1M'
        if sector_rename:
            df_sector = df_sector.rename(columns=sector_rename)

    return df_daily, df_trend, df_sector, df_favorite

try:
    df_daily, df_trend, df_sector, df_favorite = load_data()
except Exception as e:
    st.error(f"❌ Error loading Excel file: {str(e)}")
    st.stop()

# ============================================================================
# 3. DEBUG INFO & DOWNLOAD
# ============================================================================
with st.expander("🔍 DEBUG: Data Validation", expanded=False):
    st.write(f"✅ Sheet 1 (Daily Signals): {len(df_daily)} stocks, {len(df_daily.columns)} columns")
    st.write(f"✅ Sheet 2 (21-Day Trend): {len(df_trend)} stocks, {len(df_trend.columns)} columns")
    st.write(f"✅ Sheet 3 (Industry): {len(df_sector)} industries, {len(df_sector.columns)} columns")
    if not df_favorite.empty:
        st.write(f"✅ Sheet 4 (Favorites): {len(df_favorite)} stocks, {len(df_favorite.columns)} columns")

    st.write("\n📋 **Sheet 1 Columns:**", list(df_daily.columns))
    st.write("📋 **Sheet 2 Columns:**", list(df_trend.columns))
    st.write("📋 **Sheet 3 Columns:**", list(df_sector.columns))
    if not df_favorite.empty:
        st.write("📋 **Sheet 4 Columns:**", list(df_favorite.columns))

    # Check if Industry column exists
    if 'Industry' in df_trend.columns:
        st.success("✅ 'Industry' column found in Sheet 2 - Treemap will work!")
    else:
        st.warning("⚠️ 'Industry' column NOT found in Sheet 2 - will use fallback method")

with st.expander("📥 DOWNLOAD DATA", expanded=False):
    col_dl1, col_dl2 = st.columns([1, 4])
    with col_dl1:
        with open(target_file, "rb") as f:
            st.download_button(
                label="📥 Download Excel",
                data=f,
                file_name="Taiwan_Market_Analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    with col_dl2:
        st.info("📊 Excel file includes 4 sheets: Daily Signals, 21-Day Trend, Industry Analysis, and My Favorites (with 6 technical indicators).")

# ============================================================================
# 4. MY FAVORITES - COMPREHENSIVE VISUALIZATION
# ============================================================================
if not df_favorite.empty:
    st.divider()
    st.header("⭐ MY FAVORITE STOCKS (8 Portfolio Picks)")

    col_fav1, col_fav2 = st.columns([1, 1])

    # --- PIE CHART: Daily Performance ---
    with col_fav1:
        st.subheader("📊 Daily Performance Distribution")
        df_fav_perf = df_favorite.copy()
        df_fav_perf['Status'] = df_fav_perf['Pct_Day'].apply(
            lambda x: 'Strong Gain (>2%)' if x > 2 else
                     ('Mild Gain (0-2%)' if x > 0 else
                     ('Mild Loss (0 to -2%)' if x > -2 else 'Strong Loss (<-2%)'))
        )

        status_counts = df_fav_perf['Status'].value_counts()
        colors_daily = {
            'Strong Gain (>2%)': '#00CC66',
            'Mild Gain (0-2%)': '#90EE90',
            'Mild Loss (0 to -2%)': '#FFB366',
            'Strong Loss (<-2%)': '#FF4444'
        }

        fig_pie_daily = go.Figure(data=[go.Pie(
            labels=status_counts.index,
            values=status_counts.values,
            hole=0.4,
            marker=dict(colors=[colors_daily.get(x, '#CCCCCC') for x in status_counts.index]),
            textinfo='label+percent+value',
            textposition='outside',
            hovertemplate='%{label}<br>Count: %{value}<br>Percentage: %{percent}<extra></extra>'
        )])

        fig_pie_daily.update_layout(
            title=f"Today's Movement ({len(df_favorite)} stocks)",
            height=400,
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5)
        )
        st.plotly_chart(fig_pie_daily, use_container_width=True)

    # --- PIE CHART: Monthly Trend ---
    with col_fav2:
        st.subheader("📈 Monthly Trend (21 Days)")
        df_fav_trend = df_favorite.copy()
        df_fav_trend['Trend'] = df_fav_trend['Pct_1Month'].apply(
            lambda x: 'Strong Gain (>10%)' if x > 10 else
                     ('Moderate Gain (5-10%)' if x > 5 else
                     ('Mild Gain (0-5%)' if x > 0 else
                     ('Mild Loss (0 to -5%)' if x > -5 else
                     ('Moderate Loss (-5 to -10%)' if x > -10 else 'Strong Loss (<-10%)'))))
        )

        trend_counts = df_fav_trend['Trend'].value_counts()
        colors_trend = {
            'Strong Gain (>10%)': '#006600',
            'Moderate Gain (5-10%)': '#00AA00',
            'Mild Gain (0-5%)': '#90EE90',
            'Mild Loss (0 to -5%)': '#FFD700',
            'Moderate Loss (-5 to -10%)': '#FF8C00',
            'Strong Loss (<-10%)': '#CC0000'
        }

        fig_pie_trend = go.Figure(data=[go.Pie(
            labels=trend_counts.index,
            values=trend_counts.values,
            hole=0.4,
            marker=dict(colors=[colors_trend.get(x, '#CCCCCC') for x in trend_counts.index]),
            textinfo='label+percent+value',
            textposition='outside',
            hovertemplate='%{label}<br>Count: %{value}<br>Percentage: %{percent}<extra></extra>'
        )])

        fig_pie_trend.update_layout(
            title=f"1-Month Performance ({len(df_favorite)} stocks)",
            height=400,
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5)
        )
        st.plotly_chart(fig_pie_trend, use_container_width=True)

    # --- DETAILED TABLE WITH TECHNICAL INDICATORS ---
    st.subheader("📋 Portfolio Details with Technical Indicators")
    df_display = df_favorite.copy()
    df_display['Icon'] = df_display['Pct_Day'].apply(
        lambda x: '🚀' if x > 3 else ('📈' if x > 0 else ('📉' if x > -3 else '⚠️'))
    )

    display_cols = ['Icon', 'Code', 'Name_CN', 'Name', 'Price', 'Pct_Day', 'Pct_1Month',
                   'RSI', 'MACD', 'BB_Position', 'Stochastic', 'ATR_Pct', 'Vol_Trend',
                   'Money_Flow_Strength', 'QUICK_ACTION']
    available_cols = [col for col in display_cols if col in df_display.columns]

    st.dataframe(
        df_display[available_cols].sort_values('Pct_Day', ascending=False),
        hide_index=True,
        use_container_width=True,
        height=350
    )

    # --- SUMMARY STATISTICS ---
    col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)

    with col_stat1:
        avg_daily = df_favorite['Pct_Day'].mean()
        st.metric("📊 Avg Daily Change", f"{avg_daily:.2f}%",
                 delta=f"{avg_daily:.2f}%", delta_color="normal")

    with col_stat2:
        avg_monthly = df_favorite['Pct_1Month'].mean()
        st.metric("📈 Avg Monthly Change", f"{avg_monthly:.2f}%",
                 delta=f"{avg_monthly:.2f}%", delta_color="normal")

    with col_stat3:
        positive_count = len(df_favorite[df_favorite['Pct_Day'] > 0])
        st.metric("✅ Gaining Today", f"{positive_count}/{len(df_favorite)}",
                 delta=f"{positive_count/len(df_favorite)*100:.0f}%")

    with col_stat4:
        strong_stocks = len(df_favorite[df_favorite['Pct_1Month'] > 10])
        st.metric("🔥 Strong Gainers (1M)", f"{strong_stocks}/{len(df_favorite)}",
                 delta=f"{strong_stocks/len(df_favorite)*100:.0f}%")

# ============================================================================
# 5. CURRENCY CONVERSION SETTINGS
# ============================================================================
st.divider()
col_opt, _ = st.columns([2, 3])
with col_opt:
    currency_mode = st.radio(
        "💱 Liquidity Display Mode:",
        ("Original (Billion TWD)", "Million USD ($)", "Thousand Billion VND (₫)"),
        horizontal=True
    )

def convert_val(val):
    if currency_mode == "Million USD ($)":
        return val * 1000 * 0.031
    elif currency_mode == "Thousand Billion VND (₫)":
        return val * 770 / 1000
    return val

unit_label = "Billion TWD"
if "USD" in currency_mode:
    unit_label = "Million USD"
if "VND" in currency_mode:
    unit_label = "Thousand Billion VND"

# ============================================================================
# 6. HIERARCHICAL TREEMAP - INDUSTRY → STOCKS (WITH ERROR HANDLING)
# ============================================================================
st.subheader(f"1. HIERARCHICAL MONEY FLOW MAP (Industry → 40+ Stocks)")

df_treemap = df_trend.copy()
df_treemap['Liquidity'] = df_treemap['Avg_Trading_Value_B'].apply(convert_val)

# COMPREHENSIVE ERROR HANDLING: Check if Industry column exists
if 'Industry' not in df_treemap.columns:
    st.warning("⚠️ 'Industry' column not found in data. Using 'Sector' as fallback...")

    # FALLBACK: Map Sector to Industry categories
    def map_sector_to_industry(sector):
        """Map Sector to Industry categories"""
        if pd.isna(sector):
            return "Others"
        sector = str(sector)

        if "AI Server" in sector or "Power Supply" in sector or "Design Service (AI)" in sector:
            return "AI Infrastructure & Server"
        elif "IC Design" in sector or "IP Core" in sector:
            return "Semiconductor Design (Upstream)"
        elif "Foundry" in sector or "Wafer" in sector:
            return "Semiconductor Manufacturing (Midstream)"
        elif "Memory" in sector or "OSAT" in sector or "Packaging" in sector:
            return "Packaging & Memory (Downstream)"
        elif "Compound" in sector or "LED" in sector:
            return "Compound Semiconductor"
        elif "Shipping" in sector or "Airline" in sector:
            return "Transportation & Logistics"
        elif "Financial" in sector:
            return "Financial & Banking"
        elif "Equipment" in sector or "Electronic" in sector or "Electronics" in sector:
            return "Equipment & Electronic Components"
        elif sector in ["Plastics", "Steel", "Automobile", "Industrial"]:
            return "Traditional Industry"
        else:
            return "Others"

    df_treemap['Industry'] = df_treemap['Sector'].apply(map_sector_to_industry)
    st.info("✅ Industry categories created from Sector data")

df_treemap['Industry_Bold'] = df_treemap['Industry'].apply(lambda x: f"<b>{x}</b>")

try:
    # Combine stock code with Chinese name for clearer labels inside treemap boxes
    if 'Name_CN' in df_treemap.columns:
        df_treemap['Code_Label'] = df_treemap['Code'].astype(str) + ' - ' + df_treemap['Name_CN'].fillna('').astype(str)
    else:
        df_treemap['Code_Label'] = df_treemap['Code'].astype(str)

    fig_hier = px.treemap(
        df_treemap,
        path=['Industry_Bold', 'Code_Label'],
        values='Liquidity',
        color='Pct_1Month',
        color_continuous_scale='RdYlGn',
        color_continuous_midpoint=0,
        hover_data={
            'Code': True,
            'Name': True,
            'Name_CN': True,
            'Pct_1Month': ':.2f',
            'Liquidity': ':.2f',
            'Industry': True,
            'Industry_Bold': False
        },
        labels={
            'Liquidity': f'Liquidity ({unit_label})',
            'Pct_1Month': '% Change (1 Month)',
            'Industry_Bold': 'Industry'
        }
    )

    fig_hier.update_traces(
        textposition='middle center',
        textfont=dict(size=11),
        marker=dict(
            line=dict(width=2, color='white'),
            pad=dict(t=20, l=5, r=5, b=5)
        )
    )

    fig_hier.update_layout(
        height=800,
        title=f"Size = Liquidity ({unit_label}) | Color = % Change (1 Month)<br><sub>Click on industry (bold) to zoom in → Click 'All' to reset</sub>",
        font=dict(size=11),
        margin=dict(l=10, r=10, t=80, b=10)
    )

    st.plotly_chart(fig_hier, use_container_width=True)
    st.info("💡 **How to use:** Click on an industry box (**bold text**) to see individual stocks. Click 'All' at the top to return to overview.")

except Exception as e:
    st.error(f"❌ Error creating hierarchical treemap: {str(e)}")
    st.write("Falling back to simple sector-level treemap...")

    # FALLBACK 2: Simple sector-level treemap
    df_sector_plot = df_sector.copy()
    df_sector_plot['Liquidity_Display'] = df_sector_plot['Total_Trading_Value_B'].apply(convert_val)
    df_sector_plot['Value_Display'] = pd.to_numeric(df_sector_plot['Liquidity_Display'], errors='coerce').fillna(1)
    df_sector_plot['Color_Value'] = pd.to_numeric(df_sector_plot['Avg_Pct_1M'], errors='coerce').fillna(0)

    fig_map = px.treemap(
        df_sector_plot,
        path=['Industry'],
        values='Value_Display',
        color='Color_Value',
        color_continuous_scale='RdYlGn',
        color_continuous_midpoint=0,
        labels={'Value_Display': f'Liquidity ({unit_label})', 'Color_Value': '% Change (1M)'}
    )
    fig_map.update_layout(height=700, title=f"Industry Overview (Size = Liquidity, Color = Performance)")
    st.plotly_chart(fig_map, use_container_width=True)

# ============================================================================
# 7. SECTOR DETAIL SCATTER & TOP VOLUME
# ============================================================================
col1, col2 = st.columns([3, 2])

with col1:
    st.subheader("2. SECTOR DETAIL ANALYSIS (4-Quadrant Model)")

    # Check which column to use for filtering
    filter_column = 'Industry' if 'Industry' in df_trend.columns else 'Sector'
    selected_sector = st.selectbox(f"🔍 Select {filter_column.lower()} to analyze:", 
                                   sorted(df_trend[filter_column].unique()))

    df_sub = df_trend[df_trend[filter_column] == selected_sector].copy()
    df_sub['Liquidity_Display'] = df_sub['Avg_Trading_Value_B'].apply(convert_val)

    if not df_sub.empty:
        try:
            fig_scatter = px.scatter(
                df_sub,
                x="Money_Flow_Strength",
                y="Pct_1Month",
                size="Liquidity_Display",
                color="Money_Flow_Strength",
                text="Code",
                hover_name="Name",
                labels={"Money_Flow_Strength": "Money Flow (Buying Pressure)", "Pct_1Month": "Price Momentum (%)"},
                color_continuous_scale='Portland'
            )

            fig_scatter.add_vline(x=1.0, line_dash="dash", line_color="gray")
            fig_scatter.add_hline(y=0, line_dash="dash", line_color="gray")
            fig_scatter.update_layout(height=500)
            st.plotly_chart(fig_scatter, use_container_width=True)
            # --- TOP 5 OUTFLOWS: show stocks with lowest Money_Flow_Strength ---
            try:
                df_outflow = df_sub.sort_values(by='Money_Flow_Strength', ascending=True).head(5).copy()
                if 'Name_CN' in df_outflow.columns:
                    df_outflow['Code_Label'] = df_outflow['Code'].astype(str) + ' - ' + df_outflow['Name_CN'].fillna('').astype(str)
                else:
                    df_outflow['Code_Label'] = df_outflow['Code'].astype(str)

                outflow_cols = [c for c in ['Code_Label', 'Name', 'Money_Flow_Strength', 'Pct_1Month', 'Liquidity_Display'] if c in df_outflow.columns]
                st.markdown("**Top 5 Money Outflows (lowest Money Flow Strength)**")
                st.dataframe(df_outflow[outflow_cols].reset_index(drop=True), use_container_width=True, height=220)
            except Exception:
                # Non-fatal: continue if outflow table cannot be created
                pass
        except Exception as e:
            st.error(f"❌ Error creating scatter plot: {e}")
    else:
        st.warning(f"⚠️ No data available for: {selected_sector}")

with col2:
    st.subheader("3. TOP VOLUME SURGES")
    df_vol = df_daily.sort_values(by='Vol_vs_Avg', ascending=False).head(15)
    # Add Chinese name beside Code for easier identification
    df_vol = df_vol.copy()
    if 'Name_CN' in df_vol.columns:
        df_vol['Code_Label'] = df_vol['Code'].astype(str) + ' - ' + df_vol['Name_CN'].fillna('').astype(str)
    else:
        df_vol['Code_Label'] = df_vol['Code'].astype(str)

    st.dataframe(
        df_vol[['Code_Label', 'Name', 'Price', 'Vol_vs_Avg', 'Signal']],
        hide_index=True,
        use_container_width=True,
        height=500
    )

# ============================================================================
# 8. FOOTER
# ============================================================================
st.divider()
st.caption("📊 **Comprehensive Taiwan Stock Analysis** | 40+ Stocks | 9 Professional Industries | 6 Technical Indicators")
st.caption("🔄 Data updates daily via yfinance | Powered by Streamlit + Python")
st.caption("✨ Enhanced with comprehensive error handling and fallback mechanisms")

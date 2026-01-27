import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(layout="wide", page_title="Taiwan Market Dashboard 🇹🇼")
st.title("💰 DASHBOARD DÒNG TIỀN ĐÀI LOAN (SMART MONEY)")

# --- 2. TẢI DỮ LIỆU ---
current_folder = os.path.dirname(os.path.abspath(__file__))
target_file = os.path.join(current_folder, "Taiwan_Market_Data_Latest.xlsx")

if not os.path.exists(target_file):
    st.error(f"❌ Không tìm thấy file dữ liệu: {target_file}")
    st.info("Vui lòng chạy stock_tw.py trước để tạo file dữ liệu.")
    st.stop()

# Đọc các Sheet dữ liệu
@st.cache_data
def load_data():
    df_daily = pd.read_excel(target_file, sheet_name='1_Tin_Hieu_Hom_Nay')
    df_trend = pd.read_excel(target_file, sheet_name='2_Xu_Huong_21_Ngay')
    df_sector = pd.read_excel(target_file, sheet_name='3_Song_Nganh')

    # Load favorite stocks if available
    try:
        df_favorite = pd.read_excel(target_file, sheet_name='4_My_Favorite')
    except:
        df_favorite = pd.DataFrame()

    return df_daily, df_trend, df_sector, df_favorite

try:
    df_daily, df_trend, df_sector, df_favorite = load_data()
except Exception as e:
    st.error(f"❌ Lỗi đọc file Excel: {str(e)}")
    st.stop()

# Debug info
with st.expander("🔍 DEBUG: Kiểm tra dữ liệu", expanded=False):
    st.write(f"✅ Sheet 1 (Daily): {len(df_daily)} hàng, {len(df_daily.columns)} cột")
    st.write(f"✅ Sheet 2 (Trend): {len(df_trend)} hàng, {len(df_trend.columns)} cột")
    st.write(f"✅ Sheet 3 (Sector): {len(df_sector)} hàng, {len(df_sector.columns)} cột")
    if not df_favorite.empty:
        st.write(f"✅ Sheet 4 (Favorites): {len(df_favorite)} hàng, {len(df_favorite.columns)} cột")
        st.write("Columns Sheet 4:", df_favorite.columns.tolist())

# --- 3. NÚT TẢI FILE EXCEL ---
with st.expander("📥 TRÍCH XUẤT DỮ LIỆU", expanded=False):
    col_dl1, col_dl2 = st.columns([1, 4])
    with col_dl1:
        with open(target_file, "rb") as f:
            st.download_button(
                label="📥 Tải Excel về máy",
                data=f,
                file_name="Taiwan_Stock_Data_Analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    with col_dl2:
        st.info("File Excel bao gồm 4 Sheet: Tín hiệu hôm nay, Xu hướng 21 ngày, Dòng tiền sóng ngành và Danh mục yêu thích.")

# --- 4. MY FAVORITE STOCKS VISUALIZATION ---
if not df_favorite.empty:
    st.divider()
    st.header("⭐ DANH MỤC YÊU THÍCH CỦA TÔI")

    col_fav1, col_fav2 = st.columns([1, 1])

    with col_fav1:
        st.subheader("📊 Phân Bố Tăng/Giảm (Hôm Nay)")

        # Classify by daily performance
        df_fav_perf = df_favorite.copy()
        df_fav_perf['Trạng_Thái'] = df_fav_perf['%_Ngày'].apply(
            lambda x: 'Tăng Mạnh (>2%)' if x > 2 else 
                     ('Tăng Nhẹ (0-2%)' if x > 0 else 
                     ('Giảm Nhẹ (0 to -2%)' if x > -2 else 'Giảm Mạnh (<-2%)'))
        )

        # Count by status
        status_counts = df_fav_perf['Trạng_Thái'].value_counts()

        # Create pie chart with custom colors
        colors_daily = {
            'Tăng Mạnh (>2%)': '#00CC66',
            'Tăng Nhẹ (0-2%)': '#90EE90',
            'Giảm Nhẹ (0 to -2%)': '#FFB366',
            'Giảm Mạnh (<-2%)': '#FF4444'
        }

        fig_pie_daily = go.Figure(data=[go.Pie(
            labels=status_counts.index,
            values=status_counts.values,
            hole=0.4,
            marker=dict(colors=[colors_daily.get(x, '#CCCCCC') for x in status_counts.index]),
            textinfo='label+percent+value',
            textposition='outside',
            hovertemplate='%{label}<br>Số lượng: %{value}<br>Tỷ lệ: %{percent}<extra></extra>'
        )])

        fig_pie_daily.update_layout(
            title=f"Biến Động Hôm Nay ({len(df_favorite)} cổ phiếu)",
            height=400,
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5)
        )

        st.plotly_chart(fig_pie_daily, use_container_width=True)

    with col_fav2:
        st.subheader("📈 Xu Hướng 1 Tháng")

        # Classify by monthly performance
        df_fav_trend = df_favorite.copy()
        df_fav_trend['Xu_Hướng'] = df_fav_trend['%_Tăng_1_Tháng'].apply(
            lambda x: 'Tăng Mạnh (>10%)' if x > 10 else 
                     ('Tăng Vừa (5-10%)' if x > 5 else 
                     ('Tăng Nhẹ (0-5%)' if x > 0 else 
                     ('Giảm Nhẹ (0 to -5%)' if x > -5 else 
                     ('Giảm Vừa (-5 to -10%)' if x > -10 else 'Giảm Mạnh (<-10%)'))))
        )

        trend_counts = df_fav_trend['Xu_Hướng'].value_counts()

        colors_trend = {
            'Tăng Mạnh (>10%)': '#006600',
            'Tăng Vừa (5-10%)': '#00AA00',
            'Tăng Nhẹ (0-5%)': '#90EE90',
            'Giảm Nhẹ (0 to -5%)': '#FFD700',
            'Giảm Vừa (-5 to -10%)': '#FF8C00',
            'Giảm Mạnh (<-10%)': '#CC0000'
        }

        fig_pie_trend = go.Figure(data=[go.Pie(
            labels=trend_counts.index,
            values=trend_counts.values,
            hole=0.4,
            marker=dict(colors=[colors_trend.get(x, '#CCCCCC') for x in trend_counts.index]),
            textinfo='label+percent+value',
            textposition='outside',
            hovertemplate='%{label}<br>Số lượng: %{value}<br>Tỷ lệ: %{percent}<extra></extra>'
        )])

        fig_pie_trend.update_layout(
            title=f"Hiệu Suất 1 Tháng ({len(df_favorite)} cổ phiếu)",
            height=400,
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5)
        )

        st.plotly_chart(fig_pie_trend, use_container_width=True)

    # Detailed table with performance metrics
    st.subheader("📋 Chi Tiết Danh Mục")

    # Add performance indicator column
    df_display = df_favorite.copy()
    df_display['Biểu_Tượng'] = df_display['%_Ngày'].apply(
        lambda x: '🚀' if x > 3 else ('📈' if x > 0 else ('📉' if x > -3 else '⚠️'))
    )

    # Select columns for display
    display_cols = ['Biểu_Tượng', 'Mã', 'Tên Công Ty', 'Giá', '%_Ngày', '%_Tăng_1_Tháng', 
                    'RSI', 'MACD', 'Sức_Mạnh_Dòng_Tiền', 'QUICK_ACTION']

    # Filter to available columns
    available_cols = [col for col in display_cols if col in df_display.columns]

    st.dataframe(
        df_display[available_cols].sort_values('%_Ngày', ascending=False),
        hide_index=True,
        use_container_width=True,
        height=350
    )

    # Summary statistics
    col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)

    with col_stat1:
        avg_daily = df_favorite['%_Ngày'].mean()
        st.metric("📊 Trung Bình Ngày", f"{avg_daily:.2f}%", 
                 delta=f"{avg_daily:.2f}%",
                 delta_color="normal")

    with col_stat2:
        avg_monthly = df_favorite['%_Tăng_1_Tháng'].mean()
        st.metric("📈 Trung Bình Tháng", f"{avg_monthly:.2f}%",
                 delta=f"{avg_monthly:.2f}%",
                 delta_color="normal")

    with col_stat3:
        positive_count = len(df_favorite[df_favorite['%_Ngày'] > 0])
        st.metric("✅ Tăng Giá Hôm Nay", f"{positive_count}/{len(df_favorite)}",
                 delta=f"{positive_count/len(df_favorite)*100:.0f}%")

    with col_stat4:
        strong_stocks = len(df_favorite[df_favorite['%_Tăng_1_Tháng'] > 10])
        st.metric("🔥 Tăng Mạnh 1 Tháng", f"{strong_stocks}/{len(df_favorite)}",
                 delta=f"{strong_stocks/len(df_favorite)*100:.0f}%")

# --- 5. CÀI ĐẶT ĐƠN VỊ TIỀN TỆ ---
st.divider()
col_opt, _ = st.columns([2, 3])
with col_opt:
    currency_mode = st.radio(
        "Chế độ hiển thị thanh khoản trên biểu đồ:",
        ("Gốc (Tỷ TWD)", "Triệu USD ($)", "Nghìn Tỷ VNĐ (₫)"),
        horizontal=True
    )

def convert_val(val):
    if currency_mode == "Triệu USD ($)": 
        return val * 1000 * 0.031
    elif currency_mode == "Nghìn Tỷ VNĐ (₫)": 
        return val * 770 / 1000
    return val

unit_label = "Tỷ TWD"
if "USD" in currency_mode: 
    unit_label = "Triệu USD"
if "VNĐ" in currency_mode: 
    unit_label = "Nghìn Tỷ VNĐ"

df_sector['Thanh_Khoan_Hien_Thi'] = df_sector['GTGD_TB_Tỷ'].apply(convert_val)
df_trend['Thanh_Khoan_Hien_Thi'] = df_trend['GTGD_TB_Tỷ'].apply(convert_val)

# --- 6. GIAO DIỆN BIỂU ĐỒ ---
col1, col2 = st.columns([3, 2])

with col1:
    st.subheader(f"1. BẢN ĐỒ DÒNG TIỀN NGÀNH ({unit_label})")

    try:
        df_plot = df_sector.copy()
        df_plot['Value_Display'] = pd.to_numeric(df_plot['Thanh_Khoan_Hien_Thi'], errors='coerce').fillna(1)
        df_plot['Color_Value'] = pd.to_numeric(df_plot['Avg_%_1Tháng'], errors='coerce').fillna(0)

        if df_plot.empty or df_plot['Value_Display'].sum() == 0:
            st.warning("⚠️ Không có dữ liệu hợp lệ để hiển thị treemap")
        else:
            fig_map = px.treemap(
                df_plot,
                path=['Ngành'],
                values='Value_Display',
                color='Color_Value',
                color_continuous_scale='RdYlGn',
                color_continuous_midpoint=0,
                hover_data={'Value_Display': ':.2f', 'Color_Value': ':.2f'}
            )

            fig_map.update_layout(
                height=700,
                title="Độ lớn ô = Thanh khoản | Màu đỏ = Giảm, Xanh = Tăng",
                font=dict(size=10),
                margin=dict(l=5, r=80, t=40, b=5)
            )

            fig_map.update_traces(
                textposition='middle center',
                marker=dict(line=dict(width=1, color='white'))
            )

            st.plotly_chart(fig_map, use_container_width=True)

    except Exception as e:
        st.error(f"❌ Error creating treemap: {str(e)}")

with col2:
    st.subheader("2. TOP ĐỘT BIẾN KHỐI LƯỢNG")
    df_vol = df_daily.sort_values(by='%_Vol_vs_TB', ascending=False).head(12)
    st.dataframe(
        df_vol[['Mã', 'Tên Công Ty', 'Giá', '%_Vol_vs_TB', 'Tín_Hiệu_Ngày']],
        hide_index=True,
        use_container_width=True
    )

# --- 7. CHI TIẾT THEO NGÀNH ---
st.divider()
st.subheader("3. SOI CHI TIẾT THEO NGÀNH (MÔ HÌNH 4 PHẦN TƯ)")
selected_sector = st.selectbox("Chọn ngành bạn muốn soi:", df_sector['Ngành'].unique())

df_sub = df_trend[df_trend['Ngành'] == selected_sector]

if not df_sub.empty:
    try:
        fig_scatter = px.scatter(
            df_sub,
            x="Sức_Mạnh_Dòng_Tiền",
            y="%_Tăng_1_Tháng",
            size="Thanh_Khoan_Hien_Thi",
            color="Sức_Mạnh_Dòng_Tiền",
            text="Mã",
            hover_name="Tên Công Ty",
            labels={"Sức_Mạnh_Dòng_Tiền": "Lực Mua (Money Flow)", "%_Tăng_1_Tháng": "Đà Tăng Giá (%)"},
            color_continuous_scale='Portland'
        )
        fig_scatter.add_vline(x=1.0, line_dash="dash", line_color="gray")
        fig_scatter.add_hline(y=0, line_dash="dash", line_color="gray")
        fig_scatter.update_layout(height=500)
        st.plotly_chart(fig_scatter, use_container_width=True)
    except Exception as e:
        st.error(f"❌ Lỗi vẽ biểu đồ scatter: {e}")
else:
    st.warning(f"⚠️ Không có dữ liệu cho ngành: {selected_sector}")

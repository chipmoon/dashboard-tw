import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os

# ============================================================================
# 1. CẤU HÌNH TRANG
# ============================================================================
st.set_page_config(layout="wide", page_title="Vietnam Market Dashboard 🇻🇳", page_icon="💰")
st.title("💰 DASHBOARD DÒNG TIỀN VIỆT NAM - SMART MONEY FLOW")
st.markdown("### Phân Tích Ngành Chuyên Nghiệp | Chỉ Báo Kỹ Thuật | Bao Phủ Toàn Diện")

# ============================================================================
# 2. TẢI DỮ LIỆU
# ============================================================================
current_folder = os.path.dirname(os.path.abspath(__file__))
target_file = os.path.join(current_folder, "Vietnam_Market_Data_Latest.xlsx")

if not os.path.exists(target_file):
    st.error(f"❌ Không tìm thấy file dữ liệu: {target_file}")
    st.info("⚠️ Vui lòng chạy **stock_vn_final.py** trước để tạo file dữ liệu.")
    st.stop()

@st.cache_data
def load_data():
    """Tải tất cả các sheet từ file Excel"""
    df_daily = pd.read_excel(target_file, sheet_name='1_Tin_Hieu_Hom_Nay')
    df_trend = pd.read_excel(target_file, sheet_name='2_Xu_Huong_21_Ngay')
    df_sector = pd.read_excel(target_file, sheet_name='3_Song_Nganh')
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

# ============================================================================
# 3. DEBUG INFO & DOWNLOAD
# ============================================================================
with st.expander("🔍 DEBUG: Kiểm Tra Dữ Liệu", expanded=False):
    st.write(f"✅ Sheet 1 (Tín Hiệu Hôm Nay): {len(df_daily)} cổ phiếu, {len(df_daily.columns)} cột")
    st.write(f"✅ Sheet 2 (Xu Hướng 21 Ngày): {len(df_trend)} cổ phiếu, {len(df_trend.columns)} cột")
    st.write(f"✅ Sheet 3 (Dòng Ngành): {len(df_sector)} ngành, {len(df_sector.columns)} cột")
    if not df_favorite.empty:
        st.write(f"✅ Sheet 4 (Danh Mục Yêu Thích): {len(df_favorite)} cổ phiếu, {len(df_favorite.columns)} cột")

    st.write("\n📋 **Sheet 1 Columns:**", list(df_daily.columns))
    st.write("📋 **Sheet 2 Columns:**", list(df_trend.columns))
    st.write("📋 **Sheet 3 Columns:**", list(df_sector.columns))
    if not df_favorite.empty:
        st.write("📋 **Sheet 4 Columns:**", list(df_favorite.columns))

with st.expander("📥 TRÍCH XUẤT DỮ LIỆU", expanded=False):
    col_dl1, col_dl2 = st.columns([1, 4])
    with col_dl1:
        with open(target_file, "rb") as f:
            st.download_button(
                label="📥 Tải Excel về máy",
                data=f,
                file_name="Vietnam_Market_Analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    with col_dl2:
        st.info("📊 File Excel bao gồm 4 Sheet: Tín hiệu hôm nay, Xu hướng 21 ngày, Dòng tiền ngành và Danh mục yêu thích (với 6 chỉ báo kỹ thuật).")

# ============================================================================
# 4. DANH MỤC YÊU THÍCH - VISUALIZATION TOÀN DIỆN
# ============================================================================
if not df_favorite.empty:
    st.divider()
    st.header("⭐ DANH MỤC CỔ PHIẾU YÊU THÍCH")

    col_fav1, col_fav2 = st.columns([1, 1])

    # --- BIỂU ĐỒ TRÒN: Hiệu Suất Ngày ---
    with col_fav1:
        st.subheader("📊 Phân Bố Tăng/Giảm (Hôm Nay)")
        df_fav_perf = df_favorite.copy()
        df_fav_perf['Trạng_Thái'] = df_fav_perf['%_Ngày'].apply(
            lambda x: 'Tăng Mạnh (>2%)' if x > 2 else
                     ('Tăng Nhẹ (0-2%)' if x > 0 else
                     ('Giảm Nhẹ (0 to -2%)' if x > -2 else 'Giảm Mạnh (<-2%)'))
        )

        status_counts = df_fav_perf['Trạng_Thái'].value_counts()
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
            hovertemplate='%{label}<br>Số lượng: %{value}<br>Tỷ lệ: %{percent}'
        )])

        fig_pie_daily.update_layout(
            title=f"Biến Động Hôm Nay ({len(df_favorite)} cổ phiếu)",
            height=400,
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5)
        )
        st.plotly_chart(fig_pie_daily, use_container_width=True)

    # --- BIỂU ĐỒ TRÒN: Xu Hướng Tháng ---
    with col_fav2:
        st.subheader("📈 Xu Hướng 1 Tháng")
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
            hovertemplate='%{label}<br>Số lượng: %{value}<br>Tỷ lệ: %{percent}'
        )])

        fig_pie_trend.update_layout(
            title=f"Hiệu Suất 1 Tháng ({len(df_favorite)} cổ phiếu)",
            height=400,
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5)
        )
        st.plotly_chart(fig_pie_trend, use_container_width=True)

    # --- BẢNG CHI TIẾT VỚI CHỈ BÁO KỸ THUẬT ---
    st.subheader("📋 Chi Tiết Danh Mục Với Chỉ Báo Kỹ Thuật")
    df_display = df_favorite.copy()
    df_display['Biểu_Tượng'] = df_display['%_Ngày'].apply(
        lambda x: '🚀' if x > 3 else ('📈' if x > 0 else ('📉' if x > -3 else '⚠️'))
    )

    display_cols = ['Biểu_Tượng', 'Mã', 'Ngành', 'Giá', '%_Ngày', '%_Tăng_1_Tháng',
                    'RSI', 'MACD', 'BB_Position', 'Stochastic', 'ATR%', 'Vol_Trend',
                    'Sức_Mạnh_Dòng_Tiền', 'QUICK_ACTION']
    available_cols = [col for col in display_cols if col in df_display.columns]

    st.dataframe(
        df_display[available_cols].sort_values('%_Ngày', ascending=False),
        hide_index=True,
        use_container_width=True,
        height=350
    )

    # --- THỐNG KÊ TỔNG HỢP ---
    col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)

    with col_stat1:
        avg_daily = df_favorite['%_Ngày'].mean()
        st.metric("📊 Trung Bình Ngày", f"{avg_daily:.2f}%",
                 delta=f"{avg_daily:.2f}%", delta_color="normal")

    with col_stat2:
        avg_monthly = df_favorite['%_Tăng_1_Tháng'].mean()
        st.metric("📈 Trung Bình Tháng", f"{avg_monthly:.2f}%",
                 delta=f"{avg_monthly:.2f}%", delta_color="normal")

    with col_stat3:
        positive_count = len(df_favorite[df_favorite['%_Ngày'] > 0])
        st.metric("✅ Tăng Giá Hôm Nay", f"{positive_count}/{len(df_favorite)}",
                 delta=f"{positive_count/len(df_favorite)*100:.0f}%")

    with col_stat4:
        strong_stocks = len(df_favorite[df_favorite['%_Tăng_1_Tháng'] > 10])
        st.metric("🔥 Tăng Mạnh 1 Tháng", f"{strong_stocks}/{len(df_favorite)}",
                 delta=f"{strong_stocks/len(df_favorite)*100:.0f}%")

# ============================================================================
# 5. CÀI ĐẶT ĐƠN VỊ TIỀN TỆ
# ============================================================================
st.divider()
col_opt, _ = st.columns([2, 3])
with col_opt:
    currency_mode = st.radio(
        "💱 Chế Độ Hiển Thị Thanh Khoản:",
        ("Gốc (Tỷ VND)", "Triệu USD ($)", "Tỷ TWD (台幣)"),
        horizontal=True
    )

def convert_val(val):
    if currency_mode == "Triệu USD ($)":
        return val * 1000 / 25.0  # 1 USD ≈ 25,000 VND
    elif currency_mode == "Tỷ TWD (台幣)":
        return val * 1000 / 770  # 1 TWD ≈ 770 VND
    return val

unit_label = "Tỷ VND"
if "USD" in currency_mode:
    unit_label = "Triệu USD"
if "TWD" in currency_mode:
    unit_label = "Tỷ TWD"

# ============================================================================
# 6. BẢN ĐỒ PHÂN CẤP - NGÀNH → CỔ PHIẾU
# ============================================================================
st.subheader(f"1. BẢN ĐỒ DÒNG TIỀN CHI TIẾT (Ngành → Cổ Phiếu)")

df_treemap = df_trend.copy()
df_treemap['Thanh_Khoan'] = df_treemap['GTGD_TB_Tỷ'].apply(convert_val)
df_treemap['Ngành_Bold'] = df_treemap['Ngành'].apply(lambda x: f"<b>{x}</b>")

try:
    fig_hier = px.treemap(
        df_treemap,
        path=['Ngành_Bold', 'Mã'],
        values='Thanh_Khoan',
        color='%_Tăng_1_Tháng',
        color_continuous_scale='RdYlGn',
        color_continuous_midpoint=0,
        hover_data={
            'Mã': True,
            'Giá': ':.2f',
            '%_Tăng_1_Tháng': ':.2f',
            'Thanh_Khoan': ':.2f',
            'Sàn': True,
            'Ngành': True,
            'Ngành_Bold': False
        },
        labels={
            'Thanh_Khoan': f'Thanh khoản ({unit_label})',
            '%_Tăng_1_Tháng': '% Tăng 1 Tháng',
            'Ngành_Bold': 'Ngành'
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
        title=f"Kích thước = Thanh khoản ({unit_label}) | Màu sắc = % Tăng 1 Tháng<br>Click vào ngành (chữ đậm) để phóng to → Click 'All' để quay lại",
        font=dict(size=11),
        margin=dict(l=10, r=10, t=80, b=10)
    )

    st.plotly_chart(fig_hier, use_container_width=True)
    st.info("💡 **Cách sử dụng:** Click vào ô ngành (**chữ đậm**) để xem chi tiết các cổ phiếu. Click 'All' ở trên để quay lại tổng quan.")

except Exception as e:
    st.error(f"❌ Lỗi tạo biểu đồ phân cấp: {str(e)}")
    st.write("Chuyển sang treemap cơ bản...")

    # FALLBACK: Treemap ngành đơn giản
    if 'Tổng GTGD (Tỷ)' in df_sector.columns:
        df_sector_plot = df_sector.copy()
        df_sector_plot['Thanh_Khoan_Hien_Thi'] = df_sector_plot['Tổng GTGD (Tỷ)'].apply(convert_val)
        df_sector_plot['Value_Display'] = pd.to_numeric(df_sector_plot['Thanh_Khoan_Hien_Thi'], errors='coerce').fillna(1)
        df_sector_plot['Color_Value'] = pd.to_numeric(df_sector_plot['TB % Tăng (1M)'], errors='coerce').fillna(0)

        fig_map = px.treemap(
            df_sector_plot,
            path=['Ngành'],
            values='Value_Display',
            color='Color_Value',
            color_continuous_scale='RdYlGn',
            color_continuous_midpoint=0,
            labels={'Value_Display': f'Thanh khoản ({unit_label})', 'Color_Value': '% Tăng (1T)'}
        )

        fig_map.update_layout(height=700, title=f"Tổng Quan Ngành (Kích thước = Thanh khoản, Màu = Hiệu suất)")
        st.plotly_chart(fig_map, use_container_width=True)

# ============================================================================
# 7. PHÂN TÍCH CHI TIẾT NGÀNH & TOP VOLUME
# ============================================================================
col1, col2 = st.columns([3, 2])

with col1:
    st.subheader("2. SOI CHI TIẾT THEO NGÀNH (MÔ HÌNH 4 PHẦN TƯ)")

    selected_sector = st.selectbox("🔍 Chọn ngành để phân tích:",
                                   sorted(df_trend['Ngành'].unique()))

    df_sub = df_trend[df_trend['Ngành'] == selected_sector].copy()
    df_sub['Thanh_Khoan_Hien_Thi'] = df_sub['GTGD_TB_Tỷ'].apply(convert_val)

    if not df_sub.empty:
        try:
            fig_scatter = px.scatter(
                df_sub,
                x="Sức_Mạnh_Dòng_Tiền",
                y="%_Tăng_1_Tháng",
                size="Thanh_Khoan_Hien_Thi",
                color="Sức_Mạnh_Dòng_Tiền",
                text="Mã",
                labels={"Sức_Mạnh_Dòng_Tiền": "Lực Mua (Dòng Tiền)", "%_Tăng_1_Tháng": "Đà Tăng Giá (%)"},
                color_continuous_scale='Portland'
            )

            fig_scatter.add_vline(x=1.0, line_dash="dash", line_color="gray")
            fig_scatter.add_hline(y=0, line_dash="dash", line_color="gray")
            fig_scatter.update_layout(height=500)
            st.plotly_chart(fig_scatter, use_container_width=True)

            # --- TOP 5 DÒNG TIỀN YẾU ---
            try:
                df_outflow = df_sub.sort_values(by='Sức_Mạnh_Dòng_Tiền', ascending=True).head(5).copy()
                outflow_cols = [c for c in ['Mã', 'Giá', 'Sức_Mạnh_Dòng_Tiền', '%_Tăng_1_Tháng', 'Thanh_Khoan_Hien_Thi'] if c in df_outflow.columns]
                st.markdown("**Top 5 Dòng Tiền Yếu Nhất (Lực bán mạnh)**")
                st.dataframe(df_outflow[outflow_cols].reset_index(drop=True), use_container_width=True, height=220)
            except Exception:
                pass

        except Exception as e:
            st.error(f"❌ Lỗi vẽ biểu đồ scatter: {e}")
    else:
        st.warning(f"⚠️ Không có dữ liệu cho ngành: {selected_sector}")

with col2:
    st.subheader("3. TOP ĐỘT BIẾN KHỐI LƯỢNG")
    df_vol = df_daily.sort_values(by='%_Vol_vs_TB', ascending=False).head(15)

    st.dataframe(
        df_vol[['Mã', 'Giá', '%_Vol_vs_TB', 'Tín_Hiệu_Ngày']],
        hide_index=True,
        use_container_width=True,
        height=500
    )

# ============================================================================
# 8. TOP 10 DÒNG TIỀN MẠNH NHẤT (PHẦN TƯ 1) - TỔNG QUAN THỊ TRƯỜNG
# ============================================================================
st.divider()
st.header("🔥 TOP 10 DÒNG TIỀN MẠNH NHẤT (PHẦN TƯ 1) - TỔNG QUAN THỊ TRƯỜNG")
st.markdown("**Phần Tư 1 = Lực Mua Mạnh + Động Lượng Dương** | Bức Tranh Kinh Tế Liên Ngành")

# Lọc Phần Tư 1: Sức_Mạnh_Dòng_Tiền > 1.0 VÀ %_Tăng_1_Tháng > 0
df_q1 = df_trend[(df_trend['Sức_Mạnh_Dòng_Tiền'] > 1.0) & (df_trend['%_Tăng_1_Tháng'] > 0)].copy()

if len(df_q1) > 0:
    # Sắp xếp theo Sức_Mạnh_Dòng_Tiền và lấy top 10
    df_top10 = df_q1.sort_values(by='Sức_Mạnh_Dòng_Tiền', ascending=False).head(10).copy()

    # Thêm chỉ báo trạng thái
    df_top10['Trạng_Thái_Dòng_Tiền'] = df_top10['Sức_Mạnh_Dòng_Tiền'].apply(
        lambda x: '🔥 RẤT MẠNH' if x > 1.5 else ('💪 MẠNH' if x > 1.2 else '✅ TỐT')
    )

    df_top10['Trạng_Thái_Động_Lượng'] = df_top10['%_Tăng_1_Tháng'].apply(
        lambda x: '🚀 XUẤT SẮC (>15%)' if x > 15 else ('📈 MẠNH (5-15%)' if x > 5 else '✔️ DƯƠNG (0-5%)')
    )

    df_top10['Thanh_Khoan_Hien_Thi'] = df_top10['GTGD_TB_Tỷ'].apply(convert_val)

    # === BẢNG CHỈ SỐ CHÍNH ===
    col_m1, col_m2, col_m3, col_m4 = st.columns(4)

    with col_m1:
        st.metric("📊 Tổng Số CP Phần Tư 1", f"{len(df_q1)}", 
                 help="Cổ phiếu có Dòng Tiền > 1.0 và động lượng dương")

    with col_m2:
        top_flow = df_top10['Sức_Mạnh_Dòng_Tiền'].iloc[0]
        top_code = df_top10['Mã'].iloc[0]
        st.metric("🥇 Dòng Tiền Cao Nhất", f"{top_flow:.2f}", 
                 delta=f"{top_code}", delta_color="off")

    with col_m3:
        avg_flow = df_top10['Sức_Mạnh_Dòng_Tiền'].mean()
        st.metric("💪 TB Dòng Tiền (Top 10)", f"{avg_flow:.2f}",
                 help="Lực mua trung bình của top 10")

    with col_m4:
        avg_perf = df_top10['%_Tăng_1_Tháng'].mean()
        st.metric("📈 TB Hiệu Suất (1T)", f"{avg_perf:.2f}%",
                 delta=f"{avg_perf:.2f}%", delta_color="normal")

    # === BẢNG CHÍNH ===
    st.subheader("📋 Top 10 Cổ Phiếu Theo Sức Mạnh Dòng Tiền")

    display_columns = ['Mã', 'Ngành', 'Sàn', 'Giá', '%_Tăng_1_Tháng', 
                      'Sức_Mạnh_Dòng_Tiền', 'Trạng_Thái_Dòng_Tiền', 'Trạng_Thái_Động_Lượng', 
                      'Thanh_Khoan_Hien_Thi']

    available_display_cols = [col for col in display_columns if col in df_top10.columns]

    df_top10_display = df_top10[available_display_cols].copy()
    df_top10_display = df_top10_display.rename(columns={
        'Thanh_Khoan_Hien_Thi': f'Thanh Khoản ({unit_label})',
        '%_Tăng_1_Tháng': 'Tăng 1T (%)'
    })

    st.dataframe(
        df_top10_display.reset_index(drop=True),
        hide_index=True,
        use_container_width=True,
        height=400
    )

    # === BIỂU ĐỒ TRỰC QUAN ===
    col_v1, col_v2 = st.columns([1, 1])

    with col_v1:
        st.subheader("🏭 Phân Bố Ngành (Tổng Quan Kinh Tế)")
        sector_counts = df_top10['Ngành'].value_counts()

        fig_pie_sector = go.Figure(data=[go.Pie(
            labels=sector_counts.index,
            values=sector_counts.values,
            hole=0.4,
            textinfo='label+value',
            textposition='outside',
            hovertemplate='%{label}<br>Số lượng: %{value}<br>Tỷ lệ: %{percent}'
        )])

        fig_pie_sector.update_layout(
            title="Top 10 Cổ Phiếu Theo Ngành",
            height=400,
            showlegend=True,
            legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.1)
        )
        st.plotly_chart(fig_pie_sector, use_container_width=True)

        st.info("💡 **Sự đa dạng ngành cho thấy độ rộng kinh tế.** Tập trung = Rally theo ngành cụ thể.")

    with col_v2:
        st.subheader("⚖️ Dòng Tiền vs Hiệu Suất")

        # PHIÊN BẢN TỐI ƯU: Trục X rộng rãi và dễ đọc hơn
        try:
            fig_bar = make_subplots(specs=[[{"secondary_y": True}]])

            fig_bar.add_trace(
                go.Bar(
                    name='Sức Mạnh Dòng Tiền',
                    x=df_top10['Mã'],
                    y=df_top10['Sức_Mạnh_Dòng_Tiền'],
                    marker_color='#2E86AB',
                    offsetgroup=0
                ),
                secondary_y=False
            )

            fig_bar.add_trace(
                go.Bar(
                    name='Tăng 1T (%)',
                    x=df_top10['Mã'],
                    y=df_top10['%_Tăng_1_Tháng'],
                    marker_color='#06D6A0',
                    offsetgroup=1
                ),
                secondary_y=True
            )

            # TRỤC X CẢI TIẾN: Nhiều không gian và dễ đọc hơn
            fig_bar.update_xaxes(
                title_text="Mã Cổ Phiếu",
                tickangle=-65,
                tickfont=dict(size=10),
                automargin=True
            )

            fig_bar.update_yaxes(title_text="Sức Mạnh Dòng Tiền", secondary_y=False)
            fig_bar.update_yaxes(title_text="Tăng 1T (%)", secondary_y=True)

            fig_bar.update_layout(
                title_text="So Sánh Sức Mạnh vs Hiệu Suất",
                barmode='group',
                height=500,
                margin=dict(b=120),
                legend=dict(orientation="h", yanchor="top", y=-0.25, xanchor="center", x=0.5)
            )

            st.plotly_chart(fig_bar, use_container_width=True)

        except Exception as e:
            # FALLBACK: Biểu đồ cột nhóm đơn giản với trục đơn
            st.warning(f"⚠️ Sử dụng biểu đồ đơn giản (trục kép không khả dụng): {e}")

            # Chuẩn hóa giá trị về thang 0-100 để so sánh
            df_plot = df_top10.copy()
            df_plot['MF_Chuẩn_Hóa'] = (df_plot['Sức_Mạnh_Dòng_Tiền'] - df_plot['Sức_Mạnh_Dòng_Tiền'].min()) / (df_plot['Sức_Mạnh_Dòng_Tiền'].max() - df_plot['Sức_Mạnh_Dòng_Tiền'].min()) * 100
            df_plot['Perf_Chuẩn_Hóa'] = (df_plot['%_Tăng_1_Tháng'] - df_plot['%_Tăng_1_Tháng'].min()) / (df_plot['%_Tăng_1_Tháng'].max() - df_plot['%_Tăng_1_Tháng'].min()) * 100

            fig_simple = go.Figure()

            fig_simple.add_trace(go.Bar(
                name='Dòng Tiền (Chuẩn hóa)',
                x=df_plot['Mã'],
                y=df_plot['MF_Chuẩn_Hóa'],
                marker_color='#2E86AB'
            ))

            fig_simple.add_trace(go.Bar(
                name='Tăng 1T (Chuẩn hóa)',
                x=df_plot['Mã'],
                y=df_plot['Perf_Chuẩn_Hóa'],
                marker_color='#06D6A0'
            ))

            fig_simple.update_layout(
                title="So Sánh Chuẩn Hóa (thang 0-100)",
                xaxis=dict(title='Mã Cổ Phiếu', tickangle=-65, tickfont=dict(size=10), automargin=True),
                yaxis=dict(title='Giá Trị Chuẩn Hóa (0-100)'),
                barmode='group',
                height=500,
                margin=dict(b=120)
            )

            st.plotly_chart(fig_simple, use_container_width=True)

        st.info("💡 **So sánh lực mua vs động lượng giá.** Dòng tiền cao + tăng cao = Conviction mạnh.")

else:
    st.warning("⚠️ Không tìm thấy cổ phiếu nào ở Phần Tư 1 (Dòng Tiền > 1.0 và Động Lượng Dương)")
    st.info("💡 Điều này cho thấy điều kiện thị trường yếu. Xem xét chiến lược phòng thủ hoặc chờ setup tốt hơn.")

# ============================================================================
# 9. FOOTER
# ============================================================================
st.divider()
st.caption("📊 **Phân Tích Toàn Diện Thị Trường Việt Nam** | 3 Sàn HOSE-HNX-UPCOM | Đa Ngành | 6 Chỉ Báo Kỹ Thuật")
st.caption("🔄 Dữ liệu cập nhật hàng ngày qua vnstock | Powered by Streamlit + Python")
st.caption("✨ Nâng cấp với phân tích Phần Tư 1 cho insight kinh tế toàn thị trường | Tối ưu độ rõ biểu đồ")

import streamlit as st
import pandas as pd
import plotly.express as px
import os

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(layout="wide", page_title="Taiwan Market Dashboard 🇹🇼")
st.title("💰 DASHBOARD DÒNG TIỀN ĐÀI LOAN (SMART MONEY)")

# --- 2. TẢI DỮ LIỆU ---
current_folder = os.path.dirname(os.path.abspath(__file__))
target_file = os.path.join(current_folder, "Taiwan_Market_Data_Latest.xlsx")

if not os.path.exists(target_file):
    st.error(f"❌ Không tìm thấy file dữ liệu: {target_file}")
    st.info("Vui lòng đợi GitHub Actions chạy xong hoặc kiểm tra lại tên file trong kho GitHub.")
    st.stop()

# Đọc các Sheet dữ liệu
@st.cache_data
def load_data():
    df_daily = pd.read_excel(target_file, sheet_name='1_Tin_Hieu_Hom_Nay')
    df_trend = pd.read_excel(target_file, sheet_name='2_Xu_Huong_21_Ngay')
    df_sector = pd.read_excel(target_file, sheet_name='3_Song_Nganh')
    return df_daily, df_trend, df_sector

try:
    df_daily, df_trend, df_sector = load_data()
except Exception as e:
    st.error(f"❌ Lỗi đọc file Excel: {str(e)}")
    st.stop()

# Debug info
with st.expander("🔍 DEBUG: Kiểm tra dữ liệu", expanded=False):
    st.write(f"✅ Sheet 1 (Daily): {len(df_daily)} hàng, {len(df_daily.columns)} cột")
    st.write(f"✅ Sheet 2 (Trend): {len(df_trend)} hàng, {len(df_trend.columns)} cột")
    st.write(f"✅ Sheet 3 (Sector): {len(df_sector)} hàng, {len(df_sector.columns)} cột")
    st.write("Columns Sheet 3:", df_sector.columns.tolist())
    st.dataframe(df_sector.head(), use_container_width=True)

# --- 3. NÚT TẢI FILE EXCEL ---
with st.expander("📥 TRÍCH XUẤT DỮ LIỆU", expanded=True):
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
        st.info("File Excel bao gồm 3 Sheet: Tín hiệu hôm nay, Xu hướng 21 ngày và Dòng tiền sóng ngành.")

# --- 4. CÀI ĐẶT ĐƠN VỊ TIỀN TỆ ---
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

# --- 5. GIAO DIỆN BIỂU ĐỒ ---
col1, col2 = st.columns([3, 2])

with col1:
    st.subheader(f"1. BẢN ĐỒ DÒNG TIỀN NGÀNH ({unit_label})")

    # TREEMAP CREATION - ROBUST VERSION
    try:
        # Prepare data
        df_plot = df_sector.copy()
        df_plot['Value_Display'] = pd.to_numeric(df_plot['Thanh_Khoan_Hien_Thi'], errors='coerce').fillna(1)
        df_plot['Color_Value'] = pd.to_numeric(df_plot['Avg_%_1Tháng'], errors='coerce').fillna(0)

        # Validate data
        if df_plot.empty or df_plot['Value_Display'].sum() == 0:
            st.warning("⚠️ Không có dữ liệu hợp lệ để hiển thị treemap")
        else:
            # Create treemap
            fig_map = px.treemap(
                df_plot,
                path=['Ngành'],
                values='Value_Display',
                color='Color_Value',
                color_continuous_scale='RdYlGn',
                color_continuous_midpoint=0,
                hover_data={'Value_Display': ':.2f', 'Color_Value': ':.2f'}
            )

            # Update layout for better display
            fig_map.update_layout(
                height=700,
                title="Độ lớn ô = Thanh khoản | Màu đỏ = Giảm, Xanh = Tăng",
                font=dict(size=10),
                margin=dict(l=5, r=80, t=40, b=5)
            )

            # Update traces
            fig_map.update_traces(
                textposition='middle center',
                marker=dict(line=dict(width=1, color='white'))
            )

            # Display
            st.plotly_chart(fig_map, use_container_width=True)

            # Debug info
            with st.expander("Debug Info"):
                st.write(f"✅ Treemap rendered with {len(df_plot)} sectors")
                st.write(f"Total value: {df_plot['Value_Display'].sum():.2f}")

    except Exception as e:
        st.error(f"❌ Error creating treemap: {str(e)}")
        import traceback
        with st.expander("Full Error Details"):
            st.code(traceback.format_exc())
            st.write(f"Data shape: {df_sector.shape}")
            st.dataframe(df_sector.head())

with col2:
    st.subheader("2. TOP ĐỘT BIẾN KHỐI LƯỢNG")
    df_vol = df_daily.sort_values(by='%_Vol_vs_TB', ascending=False).head(12)
    st.dataframe(
        df_vol[['Mã', 'Tên Công Ty', 'Giá', '%_Vol_vs_TB', 'Tín_Hiệu_Ngày']],
        hide_index=True,
        use_container_width=True
    )

# --- 6. CHI TIẾT THEO NGÀNH ---
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

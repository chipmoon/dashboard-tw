import streamlit as st
import pandas as pd
import plotly.express as px
import os
import glob

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(layout="wide", page_title="Taiwan Market Dashboard 🇹🇼")
st.title("💰 DASHBOARD DÒNG TIỀN ĐÀI LOAN (SMART MONEY)")

# --- 2. TẢI DỮ LIỆU ---
# Tìm file Excel mới nhất do GitHub Action tạo ra
current_folder = os.path.dirname(os.path.abspath(__file__))
target_file = os.path.join(current_folder, "Taiwan_Market_Data_Latest.xlsx")

if not os.path.exists(target_file):
    st.error(f"❌ Không tìm thấy file {target_file}")
    st.info("Vui lòng đợi GitHub Actions chạy xong hoặc kiểm tra tên file trong kho GitHub.")
    st.stop()

# Đọc các Sheet dữ liệu
try:
    df_daily = pd.read_excel(target_file, sheet_name='1_Tin_Hieu_Hom_Nay')
    df_trend = pd.read_excel(target_file, sheet_name='2_Xu_Huong_21_Ngay')
    df_sector = pd.read_excel(target_file, sheet_name='3_Song_Nganh')
except Exception as e:
    st.error(f"Lỗi khi đọc file Excel: {e}")
    st.stop()

# --- 3. CÀI ĐẶT ĐƠN VỊ TIỀN TỆ (Option) ---
col_opt, _ = st.columns([2, 3])
with col_opt:
    currency_mode = st.radio(
        "Chế độ hiển thị thanh khoản:",
        ("Gốc (Tỷ TWD)", "Triệu USD ($)", "Nghìn Tỷ VNĐ (₫)"),
        horizontal=True
    )

# Hàm chuyển đổi đơn vị
def convert_val(val):
    if currency_mode == "Triệu USD ($)":
        return val * 1000 * 0.031  # 1 TWD ~ 0.031 USD
    elif currency_mode == "Nghìn Tỷ VNĐ (₫)":
        return val * 770 / 1000    # 1 TWD ~ 770 VND
    return val

unit_label = "Tỷ TWD"
if "USD" in currency_mode: unit_label = "Triệu USD"
if "VNĐ" in currency_mode: unit_label = "Nghìn Tỷ VNĐ"

# Áp dụng chuyển đổi
df_sector['Thanh_Khoan_Hien_Thi'] = df_sector['GTGD_TB_Tỷ'].apply(convert_val)
df_trend['Thanh_Khoan_Hien_Thi'] = df_trend['GTGD_TB_Tỷ'].apply(convert_val)

# --- 4. GIAO DIỆN CHÍNH ---
col1, col2 = st.columns([3, 2])

with col1:
    st.subheader(f"1. BẢN ĐỒ DÒNG TIỀN NGÀNH ({unit_label})")
    # SỬA LỖI TẠI ĐÂY: values='Thanh_Khoan_Hien_Thi' thay vì 'Tổng GTGD (Tỷ)'
    fig_map = px.treemap(
        df_sector, 
        path=['Ngành'], 
        values='Thanh_Khoan_Hien_Thi',
        color='%_Tăng_1_Tháng',
        color_continuous_scale='RdYlGn',
        hover_data=['GTGD_TB_Tỷ', 'Mã'],
        title="Độ lớn ô = Thanh khoản | Màu sắc = Hiệu suất giá 1 tháng"
    )
    fig_map.update_layout(margin=dict(t=30, l=10, r=10, b=10))
    st.plotly_chart(fig_map, use_container_width=True)

with col2:
    st.subheader("2. TOP ĐỘT BIẾN KHỐI LƯỢNG")
    # Lấy top mã có Vol tăng mạnh so với trung bình
    df_vol = df_daily.sort_values(by='%_Vol_vs_TB', ascending=False).head(12)
    st.dataframe(
        df_vol[['Mã', 'Tên Công Ty', 'Giá', '%_Vol_vs_TB', 'Tín_Hiệu_Ngày']],
        hide_index=True,
        use_container_width=True,
        column_config={
            "%_Vol_vs_TB": st.column_config.NumberColumn("Vol/TB (%)", format="%d%%"),
            "Giá": st.column_config.NumberColumn("Giá (TWD)", format="%.1f")
        }
    )

# --- 5. CHI TIẾT THEO NGÀNH ---
st.divider()
st.subheader("3. SOI CHI TIẾT TỪNG NGÀNH (MÔ HÌNH 4 PHẦN TƯ)")

selected_sector = st.selectbox("Chọn ngành bạn muốn soi:", df_sector['Ngành'].unique())

df_sub = df_trend[df_trend['Ngành'] == selected_sector]

if not df_sub.empty:
    fig_scatter = px.scatter(
        df_sub,
        x="Sức_Mạnh_Dòng_Tiền",
        y="%_Tăng_1_Tháng",
        size="Thanh_Khoan_Hien_Thi",
        color="Sức_Mạnh_Dòng_Tiền",
        text="Mã",
        hover_name="Tên Công Ty",
        title=f"Vị thế các cổ phiếu ngành: {selected_sector}",
        labels={"Sức_Mạnh_Dòng_Tiền": "Lực Mua (Money Flow)", "%_Tăng_1_Tháng": "Đà Tăng Giá (%)"},
        color_continuous_scale='Portland'
    )
    # Đường kẻ phân tách
    fig_scatter.add_vline(x=1.0, line_dash="dash", line_color="gray")
    fig_scatter.add_hline(y=0, line_dash="dash", line_color="gray")
    st.plotly_chart(fig_scatter, use_container_width=True)
else:
    st.info("Không có dữ liệu chi tiết cho ngành này.")
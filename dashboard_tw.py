import streamlit as st
import pandas as pd
import plotly.express as px
import os
import glob

st.set_page_config(layout="wide", page_title="Taiwan Stock Dashboard 🇹🇼")
st.title("🇹🇼 DASHBOARD CHỨNG KHOÁN ĐÀI LOAN")

# --- TẢI FILE ---
current_folder = os.path.dirname(os.path.abspath(__file__))
pattern = os.path.join(current_folder, "Phan_Tich_Dong_Tien_TW_*.xlsx")
list_files = glob.glob(pattern)

if not list_files:
    st.error("❌ Chưa có file dữ liệu. Hãy chạy scan_tw.py trước.")
    st.stop()

latest_file = max(list_files, key=os.path.getctime)
file_name = os.path.basename(latest_file)

with st.expander(f"✅ Dữ liệu cập nhật: {file_name}", expanded=True):
    with open(latest_file, "rb") as f:
        st.download_button("📥 Tải Excel về máy", f, file_name)

try:
    df_sector = pd.read_excel(latest_file, sheet_name='3_Song_Nganh')
    df_trend = pd.read_excel(latest_file, sheet_name='2_Xu_Huong_21_Ngay')
    df_daily = pd.read_excel(latest_file, sheet_name='1_Tin_Hieu_Hom_Nay')
except:
    st.error("Lỗi đọc file Excel.")
    st.stop()

# --- KHU VỰC 1 ---
st.subheader("1. DÒNG TIỀN NGÀNH (SECTOR HEATMAP)")
fig_map = px.treemap(
    df_sector, 
    path=['Ngành'], 
    values='GTGD_TB_Tỷ',  
    color='%_Tăng_1_Tháng',   
    color_continuous_scale='RdYlGn',
    title="Diện tích = Tiền vào (Tỷ TWD)"
)
st.plotly_chart(fig_map, use_container_width=True)

# --- KHU VỰC 2 & 3 ---
col1, col2 = st.columns([2, 1]) 

with col1:
    st.subheader("2. PHÂN TÍCH CHI TIẾT (Có tên Công Ty)")
    sectors = df_sector['Ngành'].unique().tolist()
    selected_sector = st.selectbox("Chọn Ngành:", sectors)
    
    subset = df_trend[df_trend['Ngành'] == selected_sector]
    
    if not subset.empty:
        # CẬP NHẬT: Thêm 'Tên Công Ty' vào hover_data để rê chuột là thấy tên
        fig_scatter = px.scatter(
            subset,
            x="Sức_Mạnh_Dòng_Tiền",
            y="%_Tăng_1_Tháng",
            size="GTGD_TB_Tỷ",
            color="Sức_Mạnh_Dòng_Tiền",
            text="Mã",
            hover_name="Tên Công Ty", # <--- HIỂN THỊ TÊN KHI RÊ CHUỘT
            title=f"Vị thế các mã ngành {selected_sector}",
            labels={"Sức_Mạnh_Dòng_Tiền": "Lực Mua (Vol/VolTB)", "%_Tăng_1_Tháng": "Tăng giá 1 tháng (%)"},
            color_continuous_scale='Portland'
        )
        fig_scatter.add_vline(x=1.0, line_dash="dash")
        fig_scatter.add_hline(y=0, line_dash="dash")
        st.plotly_chart(fig_scatter, use_container_width=True)
    else:
        st.warning("Không có dữ liệu.")

with col2:
    st.subheader("3. MÃ BÙNG NỔ HÔM NAY")
    top_vol = df_daily.sort_values(by='%_Vol_vs_TB', ascending=False).head(15)
    
    # CẬP NHẬT: Hiển thị cột Tên Công Ty trong bảng
    st.dataframe(
        top_vol[['Mã', 'Tên Công Ty', 'Giá', '%_Vol_vs_TB']],
        hide_index=True,
        use_container_width=True,
        column_config={
            "Tên Công Ty": st.column_config.TextColumn("Công Ty", width="medium"),
            "Giá": st.column_config.NumberColumn("Giá (TWD)", format="%.1f"),
            "%_Vol_vs_TB": st.column_config.ProgressColumn("Đột biến Vol", format="%d%%", min_value=0, max_value=500),
        }
    )
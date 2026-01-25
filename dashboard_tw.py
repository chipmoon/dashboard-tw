import streamlit as st
import pandas as pd
import plotly.express as px
import os
import glob

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(layout="wide", page_title="Market Money Flow Dashboard")
st.title("💰 DASHBOARD DÒNG TIỀN THÔNG MINH (SMART MONEY)")

# --- 2. HÀM TẢI DỮ LIỆU THÔNG MINH ---
# Tự động tìm file Excel nằm CÙNG THƯ MỤC với file code này
current_folder = os.path.dirname(os.path.abspath(__file__))
# Tìm tất cả file bắt đầu bằng 'Phan_Tich' và kết thúc bằng .xlsx
pattern = os.path.join(current_folder, "Taiwan_Market_Data_Latest.xlsx")
list_files = glob.glob(pattern)

if not list_files:
    st.error(f"❌ Không tìm thấy file dữ liệu Excel nào trong thư mục: {current_folder}")
    st.info("👉 Hãy đảm bảo bạn đã upload file Excel (ví dụ: Phan_Tich_Dong_Tien_2026-01-24.xlsx) lên cùng nơi với file dashboard.py")
    st.stop()
else:
    # Lấy file mới nhất dựa trên thời gian tạo
    latest_file = max(list_files, key=os.path.getctime)
    file_name = os.path.basename(latest_file)
    
    # Hiển thị thông báo thành công
    with st.expander(f"✅ Đang sử dụng dữ liệu từ: {file_name}", expanded=True):
        st.write("Dữ liệu được cập nhật lần cuối vào:", pd.to_datetime(os.path.getctime(latest_file), unit='s').strftime('%d/%m/%Y %H:%M'))
        
        # Thêm nút tải file về máy (Tính năng rất tiện lợi khi chia sẻ)
        with open(latest_file, "rb") as f:
            st.download_button(
                label="📥 Tải file Excel gốc",
                data=f,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    try:
        # Đọc dữ liệu (Lưu ý: Tên Sheet phải khớp với file Excel của bạn)
        # Nếu code báo lỗi "Worksheet not found", hãy mở Excel ra xem tên tab là gì và sửa lại ở dưới
        df_sector = pd.read_excel(latest_file, sheet_name='3_Song_Nganh')
        df_trend = pd.read_excel(latest_file, sheet_name='2_Xu_Huong_21_Ngay')
        df_daily = pd.read_excel(latest_file, sheet_name='1_Tin_Hieu_Hom_Nay')
    except Exception as e:
        st.error(f"❌ Lỗi đọc file Excel: {e}")
        st.warning("Gợi ý: Hãy kiểm tra lại tên các Sheet trong file Excel có đúng là: '3_Song_Nganh', '2_Xu_Huong_21_Ngay', '1_Tin_Hieu_Hom_Nay' hay không?")
        st.stop()

# --- KHU VỰC 1: VĨ MÔ (SECTOR HEATMAP) ---
st.subheader("1. BẢN ĐỒ DÒNG TIỀN NGÀNH")
if not df_sector.empty:
    # Kiểm tra tên cột chính xác để tránh lỗi
    col_size = 'Tổng GTGD (Tỷ)'
    col_color = 'TB % Tăng (1M)'
    col_hover = 'Điểm (0-100)' # Đã cập nhật theo file mới của bạn

    # Nếu file Excel thiếu cột, code sẽ báo lỗi cụ thể thay vì crash
    if col_hover not in df_sector.columns:
        st.warning(f"⚠️ Cảnh báo: Không tìm thấy cột '{col_hover}' trong dữ liệu. Biểu đồ sẽ thiếu thông tin này.")
        hover_data = []
    else:
        hover_data = [col_hover]

    fig_map = px.treemap(
        df_sector, 
        path=['Ngành'], 
        values=col_size,      # Kích thước ô
        color=col_color,      # Màu sắc
        color_continuous_scale='RdYlGn', 
        hover_data=hover_data,
        title=f"Diện tích = {col_size} | Màu sắc = {col_color}"
    )
    st.plotly_chart(fig_map, use_container_width=True)
else:
    st.info("Chưa có dữ liệu ngành.")

# --- KHU VỰC 2 & 3: VI MÔ (CỔ PHIẾU) ---
col1, col2 = st.columns([2, 1]) 

with col1:
    st.subheader("2. PHÂN LOẠI CỔ PHIẾU (Quadrant)")
    
    # Lấy danh sách ngành để tạo bộ lọc
    if 'Ngành' in df_sector.columns:
        sectors = df_sector['Ngành'].unique().tolist()
        selected_sector = st.selectbox("Chọn Ngành để soi chi tiết:", sectors)
        
        # Lọc cổ phiếu thuộc ngành đó
        subset = df_trend[df_trend['Ngành'] == selected_sector]
        
        if not subset.empty:
            # Vẽ Scatter Plot
            fig_scatter = px.scatter(
                subset,
                x="Sức_Mạnh_Dòng_Tiền",
                y="%_Tăng_1_Tháng",
                size="GTGD_TB_Tỷ",
                color="Sức_Mạnh_Dòng_Tiền",
                text="Mã",
                title=f"Vị thế các mã trong ngành {selected_sector}",
                labels={"Sức_Mạnh_Dòng_Tiền": "Sức Hút Tiền (Money Flow)", "%_Tăng_1_Tháng": "Đà Tăng Giá (Momentum)"},
                color_continuous_scale='Viridis'
            )
            # Thêm đường kẻ chia 4 vùng
            fig_scatter.add_hline(y=0, line_dash="dot", annotation_text="Tăng/Giảm")
            fig_scatter.add_vline(x=1.0, line_dash="dot", annotation_text="Tiền Vào/Ra")
            st.plotly_chart(fig_scatter, use_container_width=True)
        else:
            st.warning(f"Không tìm thấy cổ phiếu nào thuộc ngành {selected_sector} trong top theo dõi.")
    else:
        st.error("Dữ liệu ngành bị lỗi cấu trúc.")

with col2:
    st.subheader("3. TOP BÙNG NỔ HÔM NAY")
    if not df_daily.empty:
        # Lấy top 10 mã nổ Vol
        top_vol = df_daily.sort_values(by='%_Vol_vs_TB', ascending=False).head(15)
        
        # Tô màu cho bảng đẹp hơn
        st.dataframe(
            top_vol[['Mã', 'Giá', '%_Vol_vs_TB', 'Tín_Hiệu_Ngày']],
            hide_index=True,
            use_container_width=True,
            column_config={
                "%_Vol_vs_TB": st.column_config.NumberColumn(
                    "Đột biến Vol (%)",
                    format="%d %%"
                ),
                "Giá": st.column_config.NumberColumn(
                    "Giá (nghìn đ)",
                    format="%.2f"
                )
            }
        )
    else:

        st.info("Hôm nay thị trường ảm đạm, không có mã nào bùng nổ đặc biệt.")


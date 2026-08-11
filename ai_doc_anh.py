import streamlit as st
import pandas as pd
import io
import os
import re
import datetime

# ==========================================
# CẤU HÌNH TRANG & GIAO DIỆN (UI/UX)
# ==========================================
st.set_page_config(page_title="Hệ Thống Quản Trị F&B", layout="wide", page_icon="📊")

# CSS Tuỳ chỉnh: Tông màu pastel, không gian sạch sẽ, typography thanh lịch
st.markdown("""
<style>
    /* Tổng thể nền và font */
    .stApp {
        background-color: #FAFAFA;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    /* Màu chữ tiêu đề */
    h1, h2, h3 {
        color: #4A4036 !important; 
        font-weight: 600 !important;
    }
    
    /* Tuỳ chỉnh Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        padding-bottom: 5px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #F0EAE1;
        border-radius: 8px 8px 0px 0px;
        padding: 10px 24px;
        color: #6B5E53;
        font-weight: 500;
        border: none;
    }
    .stTabs [aria-selected="true"] {
        background-color: #E2D4C6 !important;
        color: #4A4036 !important;
        font-weight: bold;
        border-bottom: 3px solid #C4A484 !important;
    }
    
    /* Tuỳ chỉnh các hộp Metrics (Số liệu) */
    div[data-testid="metric-container"] {
        background-color: #FFFFFF;
        border: 1px solid #EAEAEA;
        padding: 15px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
    }
    
    /* Nút bấm */
    .stButton>button {
        background-color: #4A4036;
        color: #FFFFFF;
        border-radius: 6px;
        border: none;
        transition: all 0.3s;
    }
    .stButton>button:hover {
        background-color: #6B5E53;
        color: #FFFFFF;
    }
    
    /* Bảng dữ liệu */
    .stDataFrame {
        border-radius: 8px;
        overflow: hidden;
        border: 1px solid #EAEAEA;
    }
</style>
""", unsafe_allow_html=True)

st.title("📊 Hệ Thống Bóc Tách & Phân Tích Dữ Liệu F&B")

# ==========================================
# HÀM HỖ TRỢ XỬ LÝ TIỀN TỆ & ĐỌC DỮ LIỆU
# ==========================================
def clean_money(val):
    if pd.isna(val):
        return 0
    if isinstance(val, (int, float)):
        return int(val)
    val_str = str(val).strip()
    try:
        return int(float(val_str))
    except ValueError:
        val_str = val_str.split(',')[0]
        val_str = re.sub(r'[^\d]', '', val_str)
        if val_str == '':
            return 0
        return int(val_str)

@st.cache_data(ttl=600, show_spinner=False)
def doc_du_lieu_gg_sheet(link, ten_tab):
    file_id = re.search(r'/d/([a-zA-Z0-9-_]+)', link).group(1)
    xlsx_url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
    return pd.read_excel(xlsx_url, sheet_name=ten_tab)

tab1, tab2, tab3, tab4 = st.tabs(["📦 Xử Lý Tồn Kho", "📈 Phân Tích Bán Hàng Chi Tiết", "📋 Quản Lý Menu Gốc", "🧮 Phân Tích % Nguyên Liệu"])

# ==========================================
# TAB 1: XỬ LÝ TỒN KHO (TÍCH HỢP LƯU F5)
# ==========================================
with tab1:
    st.markdown("### 📦 Quản Lý Dữ Liệu Tồn Kho")
    st.info("Dữ liệu sau khi tải lên sẽ được hệ thống lưu lại tạm thời. Anh có thể F5 thoải mái mà không sợ mất số liệu.")
    
    loai_nguon_kho = st.radio("Nguồn cấp dữ liệu Tồn Kho:", 
                              options=["🔗 Liên kết Google Sheet (Dữ liệu nền AppSheet)", "📁 Tải file (Excel/CSV) thủ công"], horizontal=True)
    
    df_kho = pd.DataFrame()
    
    # Kịch bản 1: AppSheet
    if loai_nguon_kho == "🔗 Liên kết Google Sheet (Dữ liệu nền AppSheet)":
        col_link, col_tab, col_btn = st.columns([3, 1, 1])
        with col_link:
            link_appsheet = st.text_input("🔗 Nhập link Google Sheet (AppSheet):")
        with col_tab:
            tab_appsheet = st.text_input("📌 Tên Tab:", value="Phiếu Kiểm Kho")
        with col_btn:
            st.write("") # Căn chỉnh nút bấm
            st.write("")
            if st.button("🔄 Lấy Dữ Liệu"):
                doc_du_lieu_gg_sheet.clear()
            
        if link_appsheet:
            try:
                df_kho = doc_du_lieu_gg_sheet(link_appsheet, tab_appsheet)
                st.success(f"✅ Đã kết nối thành công với kho dữ liệu `{tab_appsheet}`!")
            except Exception as e:
                st.error("❌ Không thể kết nối. Vui lòng kiểm tra lại link Google Sheet.")
                
    # Kịch bản 2: Tải file thủ công + Lưu F5
    else:
        # Kiểm tra xem có file cũ đang lưu không
        if os.path.exists("temp_kho.csv"):
            df_kho = pd.read_csv("temp_kho.csv")
            st.success("✅ Đã khôi phục dữ liệu Tồn Kho từ phiên làm việc trước.")
            if st.button("🗑️ Xóa dữ liệu cũ để tải file mới"):
                os.remove("temp_kho.csv")
                st.rerun()
        else:
            file_kho = st.file_uploader("Tải file Tồn Kho lên đây", type=['xlsx', 'xls', 'csv'], key="kho")
            if file_kho is not None:
                if file_kho.name.endswith('.csv'):
                    df_kho = pd.read_csv(file_kho)
                else:
                    df_kho = pd.read_excel(file_kho)
                # Tự động lưu file vào máy chủ để chống F5
                df_kho.to_csv("temp_kho.csv", index=False)
                st.rerun() # Tải lại trang để áp dụng

    # Hiển thị và xử lý nếu có dữ liệu
    if not df_kho.empty:
        if "Chi nhánh" in df_kho.columns:
            st.write("---")
            danh_sach_cn = ["Tất cả các chi nhánh"] + df_kho["Chi nhánh"].dropna().unique().tolist()
            cn_chon = st.selectbox("🏢 Lọc theo Chi nhánh:", options=danh_sach_cn)
            if cn_chon != "Tất cả các chi nhánh":
                df_kho = df_kho[df_kho["Chi nhánh"] == cn_chon]
                
        st.write("🔍 **Bản xem trước dữ liệu:**")
        st.dataframe(df_kho.head(5), use_container_width=True)
        
        st.write("---")
        st.write("### 🛠️ Trích Xuất Báo Cáo Chuẩn")
        danh_sach_cot = df_kho.columns.tolist()
        col1, col2, col3 = st.columns(3)
        with col1:
            ma_hang_col = st.selectbox("Cột Mã/Tên Hàng:", options=["Không chọn"] + danh_sach_cot)
        with col2:
            ton_kho_col = st.selectbox("Cột Số Lượng Tồn:", options=["Không chọn"] + danh_sach_cot)
        with col3:
            dvt_col = st.selectbox("Cột ĐVT/Ngày:", options=["Không chọn"] + danh_sach_cot)
            
        if st.button("🚀 Bắt đầu Trích Xuất"):
            if ma_hang_col != "Không chọn" and ton_kho_col != "Không chọn" and dvt_col != "Không chọn":
                df_kq = df_kho[[ma_hang_col, ton_kho_col, dvt_col]].copy()
                df_kq.columns = ['Định danh', 'Số liệu Tồn Kho', 'Thông tin phụ (ĐVT/Ngày)']
                df_kq = df_kq.dropna(how='all')
                
                def format_so(x):
                    try: return f"{int(float(x)):,}"
                    except: return x
                df_kq['Số liệu Tồn Kho'] = df_kq['Số liệu Tồn Kho'].apply(format_so)
                
                st.success(f"🎉 Đã trích xuất {len(df_kq)} dòng dữ liệu.")
                st.dataframe(df_kq, use_container_width=True)
            else:
                st.warning("⚠️ Anh vui lòng chọn đủ 3 cột để trích xuất nhé!")

# ==========================================
# TAB 2: PHÂN TÍCH BÁN HÀNG CHI TIẾT (TÍCH HỢP LƯU F5)
# ==========================================
with tab2:
    st.markdown("### 📈 Phân Tích Bán Hàng iPOS")
    st.info("Dữ liệu iPOS sẽ được lưu giữ nguyên trên màn hình. Anh chỉ cần tải 1 lần là có thể dùng qua Tab 4 tính % Food Cost.")
    
    df_ban_goc = pd.DataFrame()
    
    if os.path.exists("temp_ipos.csv"):
        df_ban_goc = pd.read_csv("temp_ipos.csv")
        st.success("✅ Đã khôi phục báo cáo iPOS từ phiên làm việc trước.")
        if st.button("🗑️ Xóa báo cáo iPOS cũ"):
            os.remove("temp_ipos.csv")
            st.rerun()
    else:
        file_ipos = st.file_uploader("Tải báo cáo chi tiết giao dịch iPOS lên đây", type=['xlsx', 'xls', 'csv'], key="ipos_chitiet")
        if file_ipos is not None:
            try:
                if file_ipos.name.endswith('.csv'):
                    df_ban_goc = pd.read_csv(file_ipos)
                else:
                    df_ban_goc = pd.read_excel(file_ipos)
                    if 'Tên hàng' not in df_ban_goc.columns and len(df_ban_goc) > 0:
                        df_ban_goc = pd.read_excel(file_ipos, header=1) 
                
                # Lưu file lại để chống F5
                df_ban_goc.to_csv("temp_ipos.csv", index=False)
                st.rerun()
            except Exception as e:
                st.error(f"❌ Lỗi đọc file: {e}")

    # Xử lý phân tích nếu có data
    if not df_ban_goc.empty:
        df_ban = df_ban_goc.dropna(subset=['Tên hàng']).copy()
        df_ban = df_ban[~df_ban['Tên hàng'].astype(str).str.strip().isin(['-', 'Tổng cộng'])]
        df_ban['Số lượng'] = pd.to_numeric(df_ban['Số lượng'], errors='coerce').fillna(0)
        df_ban['Tổng tiền'] = pd.to_numeric(df_ban['Tổng tiền'], errors='coerce').fillna(0)
        df_ban = df_ban[df_ban['Tổng tiền'] > 0] 
        
        if 'Thời gian' in df_ban.columns:
            df_daily = df_ban.groupby('Thời gian')['Tổng tiền'].sum().to_dict()
            st.session_state['nhat_ky_doanh_thu_offline'] = df_daily

        def extract_size(name):
            match = re.search(r'\((S|M|L)\)$', str(name).strip(), re.IGNORECASE)
            return match.group(1).upper() if match else 'Không ghi rõ'

        def extract_base_name(name):
            return re.sub(r'\s*\((S|M|L)\)$', '', str(name).strip(), flags=re.IGNORECASE).strip()

        df_ban['Size'] = df_ban['Tên hàng'].apply(extract_size)
        df_ban['Tên món gốc'] = df_ban['Tên hàng'].apply(extract_base_name)

        st.write("---")
        st.subheader("🏆 TỔNG QUAN KINH DOANH")
        col1, col2 = st.columns(2)
        col1.metric("🥤 Tổng Ly Bán Ra", f"{df_ban['Số lượng'].sum():,.0f} ly")
        col2.metric("💰 Tổng Doanh Thu", f"{df_ban['Tổng tiền'].sum():,.0f} VNĐ")

        if 'Thời gian' in df_ban.columns:
            st.write("---")
            st.subheader("📈 1. Biểu Đồ Doanh Thu Theo Ngày")
            df_trend = df_ban.groupby('Thời gian')['Tổng tiền'].sum().reset_index()
            df_trend = df_trend.set_index('Thời gian')
            st.line_chart(df_trend)
            
        st.write("---")
        st.subheader("🔍 2. Bảng Xếp Hạng Món (Theo Size)")
        
        if 'Thời gian' in df_ban.columns:
            df_ban['Date_Obj'] = pd.to_datetime(df_ban['Thời gian'], format='%d/%m/%Y', errors='coerce').dt.date
            min_date = df_ban['Date_Obj'].min()
            max_date = df_ban['Date_Obj'].max()
            
            if pd.notnull(min_date) and pd.notnull(max_date):
                ngay_loc_mon = st.date_input("🗓️ Lọc xếp hạng món theo ngày:", 
                                             value=(min_date, max_date), min_value=min_date, max_value=max_date)
                if isinstance(ngay_loc_mon, tuple):
                    start_d = ngay_loc_mon[0]
                    end_d = ngay_loc_mon[1] if len(ngay_loc_mon) > 1 else start_d
                else:
                    start_d = end_d = ngay_loc_mon
                df_ban_mon = df_ban[(df_ban['Date_Obj'] >= start_d) & (df_ban['Date_Obj'] <= end_d)]
            else:
                df_ban_mon = df_ban
        else:
            df_ban_mon = df_ban
            
        df_mon = df_ban_mon.groupby(['Tên món gốc', 'Size'], as_index=False).agg({
            'Số lượng': 'sum',
            'Tổng tiền': 'sum'
        }).sort_values(by='Số lượng', ascending=False)
        
        if not df_mon.empty:
            df_mon_hien_thi = df_mon.copy()
            df_mon_hien_thi['Tổng tiền'] = df_mon_hien_thi['Tổng tiền'].apply(lambda x: f"{int(x):,} đ")
            st.dataframe(df_mon_hien_thi, use_container_width=True)

        if 'Số điện thoại' in df_ban.columns and 'Mã hoá đơn' in df_ban.columns:
            st.write("---")
            st.subheader("🤝 3. Phân Tích Lượt Khách Quay Lại (SĐT)")
            df_kh = df_ban.dropna(subset=['Số điện thoại']).copy()
            df_kh_stats = df_kh.groupby('Số điện thoại').agg(
                So_Lan_Mua=('Mã hoá đơn', 'nunique'),
                Tong_Chi_Tieu=('Tổng tiền', 'sum'),
                Ten_Khach=('Tên khách', 'first')
            ).reset_index()
            
            tong_kh = len(df_kh_stats)
            kh_quay_lai = len(df_kh_stats[df_kh_stats['So_Lan_Mua'] > 1])
            ty_le = (kh_quay_lai / tong_kh * 100) if tong_kh > 0 else 0
            
            c1, c2, c3 = st.columns(3)
            c1.metric("👤 Số Khách Để Lại Info", f"{tong_kh} người")
            c2.metric("🔄 Khách VIP (Quay Lại)", f"{kh_quay_lai} người")
            c3.metric("🔥 Tỷ Lệ Giữ Chân Khách", f"{ty_le:.1f}%")
            
            top_kh = df_kh_stats.sort_values(by='So_Lan_Mua', ascending=False).head(10)
            top_kh.columns = ['Số Điện Thoại', 'Số Lần Mua', 'Tổng Chi Tiêu', 'Tên Khách Hàng']
            top_kh_hien_thi = top_kh.copy()
            top_kh_hien_thi['Tổng Chi Tiêu'] = top_kh_hien_thi['Tổng Chi Tiêu'].apply(lambda x: f"{int(x):,} đ")
            st.dataframe(top_kh_hien_thi[['Tên Khách Hàng', 'Số Điện Thoại', 'Số Lần Mua', 'Tổng Chi Tiêu']], use_container_width=True)

# ==========================================
# TAB 3: QUẢN LÝ MENU GỐC
# ==========================================
with tab3:
    st.markdown("### 📋 Dữ Liệu Bảng Giá (Menu)")
    st.info("Tải file Excel chứa Menu lên để hệ thống phân tích đối chiếu sau này.")
    
    if os.path.exists("menu_goc.csv"):
        df_menu = pd.read_csv("menu_goc.csv")
        st.success("✅ Đã khôi phục Menu gốc.")
        st.dataframe(df_menu, use_container_width=True)
        if st.button("🗑️ Xóa Menu hiện tại"):
            os.remove("menu_goc.csv")
            st.rerun()
    else:
        file_menu = st.file_uploader("Tải file Menu Excel lên", type=['xlsx', 'csv'], key="menu_upload")
        if file_menu:
            df_menu = pd.read_csv(file_menu) if file_menu.name.endswith('.csv') else pd.read_excel(file_menu)
            df_menu.to_csv("menu_goc.csv", index=False, encoding='utf-8-sig')
            st.rerun()

# ==========================================
# TAB 4: PHÂN TÍCH % NGUYÊN LIỆU
# ==========================================
with tab4:
    st.markdown("### 🧮 Quản Trị Tỷ Lệ % Nguyên Liệu (Food Cost)")
    st.info("Công cụ tính toán Tỷ Lệ Thực tế. Chọn ngày để hệ thống tự động bóc tách doanh thu.")
    
    today = datetime.date.today()
    ngay_tinh = st.date_input("🎯 Chọn ngày kiểm toán:", value=(today, today))
    
    if isinstance(ngay_tinh, tuple):
        start_date = ngay_tinh[0]
        end_date = ngay_tinh[1] if len(ngay_tinh) > 1 else start_date
    else:
        start_date = end_date = ngay_tinh
        
    date_list = pd.date_range(start_date, end_date).strftime('%d/%m/%Y').tolist()
    chuoi_hien_thi = f"từ {start_date.strftime('%d/%m')} đến {end_date.strftime('%d/%m')}" if start_date != end_date else f"ngày {start_date.strftime('%d/%m')}"
    
    st.write(f"#### 1. Doanh Thu {chuoi_hien_thi}")
    col_dt1, col_dt2 = st.columns(2)
    
    with col_dt1:
        st.write("**💵 Doanh thu Offline (tại chỗ)**")
        dt_offline_default = 0
        if 'nhat_ky_doanh_thu_offline' in st.session_state:
            nhat_ky = st.session_state['nhat_ky_doanh_thu_offline']
            dt_goi_y = sum([int(nhat_ky.get(d, 0)) for d in date_list])
            if dt_goi_y > 0:
                dt_offline_default = dt_goi_y
                st.caption(f"*(Hệ thống tự động lấp số: {dt_goi_y:,} đ từ Tab 2)*")
        
        dt_offline = st.number_input("Chỉnh sửa Doanh thu Offline:", min_value=0, value=dt_offline_default, step=10000)

    with col_dt2:
        st.write("**🛵 Doanh thu Online (ShopeeFood, Grab...)**")
        link_gg_sheet = st.text_input("🔗 Link đối soát (Google Sheet):")
        
        col_tab, col_btn = st.columns([2, 1])
        with col_tab:
            ten_tab = st.text_input("📌 Tên Tab:", value="Lục_TS") 
        with col_btn:
            st.write("")
            st.write("")
            if st.button("🔄 Cập nhật"):
                doc_du_lieu_gg_sheet.clear()
            
        dt_online = 0
        if link_gg_sheet:
            try:
                df_onl = doc_du_lieu_gg_sheet(link_gg_sheet, ten_tab)
                
                c1_onl, c2_onl = st.columns(2)
                with c1_onl:
                    cot_ngay_onl = st.selectbox("Cột Ngày:", options=["Không chọn"] + df_onl.columns.tolist())
                with c2_onl:
                    cot_dt_onl = st.selectbox("Cột Doanh Thu:", options=["Không chọn"] + df_onl.columns.tolist())
                
                if cot_ngay_onl != "Không chọn" and cot_dt_onl != "Không chọn":
                    df_onl['Ngay_Chuan'] = pd.to_datetime(df_onl[cot_ngay_onl], dayfirst=True, errors='coerce').dt.date
                    df_onl_ngay = df_onl[(df_onl['Ngay_Chuan'] >= start_date) & (df_onl['Ngay_Chuan'] <= end_date)].copy()
                    
                    if not df_onl_ngay.empty:
                        df_onl_ngay[cot_dt_onl] = df_onl_ngay[cot_dt_onl].apply(clean_money)
                        dt_online_tong = int(df_onl_ngay[cot_dt_onl].sum())
                        st.caption(f"*(Doanh thu Online hệ thống tìm thấy: {dt_online_tong:,} đ)*")
                        dt_online = dt_online_tong
                    else:
                        dt_online = st.number_input("Nhập tay Doanh thu Online:", min_value=0, value=0, step=10000)
            except Exception as e:
                dt_online = st.number_input("Nhập tay Doanh thu Online:", min_value=0, value=0, step=10000)
        else:
            dt_online = st.number_input("Nhập Doanh thu Online:", min_value=0, value=0, step=10000)

    tong_dt_hien_tai = dt_offline + dt_online
    st.write("")
    st.markdown(f"##### 🌟 TỔNG CỘNG DOANH THU: <span style='color:#C4A484;'>{tong_dt_hien_tai:,} VNĐ</span>", unsafe_allow_html=True)
    
    st.write("---")
    st.write(f"#### 2. Dữ Liệu Kho {chuoi_hien_thi}")
    col_k1, col_k2, col_k3 = st.columns(3)
    with col_k1:
        ton_dau = st.number_input("📦 Tồn Đầu (VNĐ)", min_value=0, value=0, step=100000)
    with col_k2:
        nhap_trong_ngay = st.number_input("🛒 Nhập Hàng (VNĐ)", min_value=0, value=0, step=100000)
    with col_k3:
        ton_cuoi = st.number_input("📉 Tồn Cuối (VNĐ)", min_value=0, value=0, step=100000)

    if st.button(f"🧮 Báo Cáo % Nguyên Liệu {chuoi_hien_thi}"):
        tong_doanh_thu = dt_offline + dt_online
        chi_phi_nl = ton_dau + nhap_trong_ngay - ton_cuoi
        
        st.write("---")
        if tong_doanh_thu == 0:
            st.warning("⚠️ Tổng doanh thu đang bằng 0, không thể chia tỷ lệ % được.")
        elif chi_phi_nl < 0:
            st.error("❌ Số liệu kho đang bị âm (Tồn cuối lớn hơn tổng Tồn đầu + Nhập).")
        else:
            phan_tram_nl = (chi_phi_nl / tong_doanh_thu) * 100
            
            c1, c2, c3 = st.columns(3)
            c1.metric("💰 Doanh Thu Thực Tế", f"{tong_doanh_thu:,} đ")
            c2.metric("🔪 Chi Phí Nguyên Liệu", f"{chi_phi_nl:,} đ")
            
            if phan_tram_nl <= 35:
                c3.metric("🎯 Tỷ lệ %", f"{phan_tram_nl:.1f} %", "Tuyệt vời (≤35%)")
            elif phan_tram_nl <= 40:
                c3.metric("🎯 Tỷ lệ %", f"{phan_tram_nl:.1f} %", "Chấp nhận được", delta_color="off")
            else:
                c3.metric("🎯 Tỷ lệ %", f"{phan_tram_nl:.1f} %", "Báo Động Hụt Kho (>40%)", delta_color="inverse")

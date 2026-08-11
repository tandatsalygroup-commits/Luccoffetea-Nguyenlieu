import streamlit as st
import pandas as pd
import io
import os
import re
import datetime

st.set_page_config(page_title="Hệ Thống Quản Trị F&B", layout="wide")
st.title("📊 Hệ Thống Bóc Tách & Phân Tích Dữ Liệu")

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

tab1, tab2, tab3, tab4 = st.tabs(["📦 Xử Lý Tồn Kho (AppSheet)", "📈 Phân Tích Bán Hàng Chi Tiết", "📋 Quản Lý Menu Gốc", "🧮 Phân Tích % Nguyên Liệu"])

# ==========================================
# TAB 1: XỬ LÝ TỒN KHO (TÍCH HỢP APPSHEET)
# ==========================================
with tab1:
    st.info("Kết nối trực tiếp với dữ liệu AppSheet hoặc tải file thủ công để lọc dữ liệu Tồn Kho.")
    
    loai_nguon_kho = st.radio("Cấp dữ liệu Tồn Kho bằng cách:", 
                              options=["🔗 Liên kết Google Sheet (Dữ liệu nền của AppSheet)", "📁 Tải file (Excel/CSV) thủ công"])
    
    df_kho = pd.DataFrame()
    
    # Kịch bản 1: Kéo trực tiếp từ Google Sheet của AppSheet
    if loai_nguon_kho == "🔗 Liên kết Google Sheet (Dữ liệu nền của AppSheet)":
        col_link, col_tab = st.columns([3, 1])
        with col_link:
            link_appsheet = st.text_input("🔗 Nhập link Google Sheet chứa Data của AppSheet (Mở quyền 'Bất kỳ ai có liên kết'):")
        with col_tab:
            tab_appsheet = st.text_input("📌 Tên Tab:", value="Phiếu Kiểm Kho")
            
        if st.button("🔄 Tải / Cập nhật Dữ Liệu AppSheet"):
            doc_du_lieu_gg_sheet.clear()
            
        if link_appsheet:
            try:
                with st.spinner("Đang kết nối với AppSheet..."):
                    df_kho = doc_du_lieu_gg_sheet(link_appsheet, tab_appsheet)
                st.success(f"✅ Đã kết nối thành công với kho dữ liệu `{tab_appsheet}`!")
                
                # Bộ lọc Chi Nhánh thông minh (Tự động nhận diện cột 'Chi nhánh')
                if "Chi nhánh" in df_kho.columns:
                    st.write("---")
                    st.write("### 🏢 Lọc Dữ Liệu Theo Chi Nhánh")
                    danh_sach_cn = ["Tất cả các chi nhánh"] + df_kho["Chi nhánh"].dropna().unique().tolist()
                    cn_chon = st.selectbox("Chọn chi nhánh cần xem:", options=danh_sach_cn)
                    
                    if cn_chon != "Tất cả các chi nhánh":
                        df_kho = df_kho[df_kho["Chi nhánh"] == cn_chon]
                        
                st.write("🔍 **Bản xem trước dữ liệu AppSheet:**")
                st.dataframe(df_kho, use_container_width=True)
                
            except Exception as e:
                st.error("❌ Không thể kết nối. Vui lòng kiểm tra lại link Google Sheet, Tên Tab hoặc Quyền truy cập.")
                
    # Kịch bản 2: Tải file thủ công (Như cũ)
    else:
        file_kho = st.file_uploader("Tải file Tồn Kho lên đây", type=['xlsx', 'xls', 'csv'], key="kho")
        if file_kho is not None:
            if file_kho.name.endswith('.csv'):
                df_kho = pd.read_csv(file_kho)
            else:
                df_kho = pd.read_excel(file_kho)
            st.dataframe(df_kho, use_container_width=True)

    # KHU VỰC CHUẨN HÓA DỮ LIỆU KHO (Áp dụng cho cả 2 kịch bản)
    if not df_kho.empty:
        st.write("---")
        st.write("### 🛠️ Bóc Tách & Chuẩn Hóa Chi Tiết Kho")
        danh_sach_cot = df_kho.columns.tolist()
        col1, col2, col3 = st.columns(3)
        with col1:
            ma_hang_col = st.selectbox("Cột Mã Hàng (hoặc Tên Hàng):", options=["Không chọn"] + danh_sach_cot, key="ma")
        with col2:
            ton_kho_col = st.selectbox("Cột Số Lượng / Giá Trị Tồn Kho:", options=["Không chọn"] + danh_sach_cot, key="ton")
        with col3:
            dvt_col = st.selectbox("Cột ĐVT (hoặc Ngày Kiểm):", options=["Không chọn"] + danh_sach_cot, key="dvt")
            
        if st.button("🚀 Trích Xuất Báo Cáo Chuẩn"):
            if ma_hang_col != "Không chọn" and ton_kho_col != "Không chọn" and dvt_col != "Không chọn":
                df_kq = df_kho[[ma_hang_col, ton_kho_col, dvt_col]].copy()
                df_kq.columns = ['Định danh', 'Số liệu Tồn Kho', 'Thông tin phụ (ĐVT/Ngày)']
                df_kq = df_kq.dropna(how='all')
                
                # Format tiền tệ hoặc số lượng cho đẹp
                # Ép kiểu an toàn, nếu là số thì format, nếu là chữ thì giữ nguyên
                def format_so(x):
                    try:
                        return f"{int(float(x)):,}"
                    except:
                        return x
                df_kq['Số liệu Tồn Kho'] = df_kq['Số liệu Tồn Kho'].apply(format_so)
                
                st.success(f"🎉 Đã trích xuất gọn gàng {len(df_kq)} dòng dữ liệu.")
                st.dataframe(df_kq, use_container_width=True)
            else:
                st.warning("⚠️ Anh vui lòng chọn đủ 3 cột để hệ thống tiến hành trích xuất nhé!")

# ==========================================
# TAB 2: PHÂN TÍCH BÁN HÀNG CHI TIẾT
# ==========================================
with tab2:
    st.success("Tải 'Báo cáo chi tiết giao dịch' từ iPOS. Hệ thống sẽ lưu nhật ký doanh thu từng ngày cho Tab 4.")
    file_ipos = st.file_uploader("Tải file bán hàng iPOS lên đây", type=['xlsx', 'xls', 'csv'], key="ipos_chitiet")

    if file_ipos is not None:
        try:
            if file_ipos.name.endswith('.csv'):
                df_ipos = pd.read_csv(file_ipos)
            else:
                df_ipos = pd.read_excel(file_ipos)
                if 'Tên hàng' not in df_ipos.columns and len(df_ipos) > 0:
                    df_ipos = pd.read_excel(file_ipos, header=1) 
            
            df_ban = df_ipos.dropna(subset=['Tên hàng']).copy()
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
            st.subheader("🏆 TỔNG QUAN DOANH THU & SỐ LƯỢNG (TOÀN BỘ FILE)")
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
                    ngay_loc_mon = st.date_input("🗓️ Lọc bảng xếp hạng món theo ngày (hoặc khoảng ngày):", 
                                                 value=(min_date, max_date), 
                                                 min_value=min_date, 
                                                 max_value=max_date,
                                                 key="loc_ngay_mon")
                    
                    if isinstance(ngay_loc_mon, tuple):
                        if len(ngay_loc_mon) == 2:
                            start_d, end_d = ngay_loc_mon
                        elif len(ngay_loc_mon) == 1:
                            start_d = end_d = ngay_loc_mon[0]
                        else:
                            start_d = end_d = min_date
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
            
            if df_mon.empty:
                st.warning("⚠️ Không có món nào được bán trong khoảng ngày anh vừa chọn.")
            else:
                df_mon_hien_thi = df_mon.copy()
                df_mon_hien_thi['Tổng tiền'] = df_mon_hien_thi['Tổng tiền'].apply(lambda x: f"{int(x):,}đ")
                st.dataframe(df_mon_hien_thi, use_container_width=True)

            if 'Số điện thoại' in df_ban.columns and 'Mã hoá đơn' in df_ban.columns:
                st.write("---")
                st.subheader("🤝 3. Phân Tích Lượt Khách Quay Lại (Dựa trên SĐT)")
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
                c2.metric("🔄 Khách Quay Lại (>1 lần)", f"{kh_quay_lai} người")
                c3.metric("🔥 Tỷ Lệ Giữ Chân Khách", f"{ty_le:.1f}%")
                
                st.write("**🏆 Top 10 Khách VIP Mua Thường Xuyên Nhất:**")
                top_kh = df_kh_stats.sort_values(by='So_Lan_Mua', ascending=False).head(10)
                top_kh.columns = ['Số Điện Thoại', 'Số Lần Mua', 'Tổng Chi Tiêu', 'Tên Khách Hàng']
                top_kh_hien_thi = top_kh.copy()
                top_kh_hien_thi['Tổng Chi Tiêu'] = top_kh_hien_thi['Tổng Chi Tiêu'].apply(lambda x: f"{int(x):,}đ")
                st.dataframe(top_kh_hien_thi[['Tên Khách Hàng', 'Số Điện Thoại', 'Số Lần Mua', 'Tổng Chi Tiêu']], use_container_width=True)

        except Exception as e:
            st.error(f"❌ Có lỗi trong quá trình bóc tách: {e}")

# ==========================================
# TAB 3: QUẢN LÝ MENU GỐC
# ==========================================
with tab3:
    st.info("Tải file Excel chứa Menu của quán lên đây để làm dữ liệu đối chiếu.")
    file_menu = st.file_uploader("Tải file Menu Excel lên", type=['xlsx', 'csv'], key="menu_upload")
    if file_menu:
        df_menu = pd.read_csv(file_menu) if file_menu.name.endswith('.csv') else pd.read_excel(file_menu)
        st.dataframe(df_menu)
        if st.button("💾 Lưu làm Cơ Sở Dữ Liệu Menu"):
            df_menu.to_csv("menu_goc.csv", index=False, encoding='utf-8-sig')
            st.success("✅ Đã lưu Menu gốc thành công!")

# ==========================================
# TAB 4: PHÂN TÍCH % NGUYÊN LIỆU
# ==========================================
with tab4:
    st.info("Công cụ tính toán Tỷ Lệ % Nguyên Liệu (Food Cost). Có thể chọn 1 ngày hoặc kéo chọn 1 khoảng ngày.")
    
    today = datetime.date.today()
    ngay_tinh = st.date_input("🎯 Chọn ngày (hoặc khoảng ngày):", value=(today, today))
    
    if isinstance(ngay_tinh, tuple):
        if len(ngay_tinh) == 2:
            start_date, end_date = ngay_tinh
        elif len(ngay_tinh) == 1:
            start_date = end_date = ngay_tinh[0]
        else:
            start_date = end_date = today
    else:
        start_date = end_date = ngay_tinh
        
    date_list = pd.date_range(start_date, end_date).strftime('%d/%m/%Y').tolist()
    chuoi_hien_thi = f"từ {start_date.strftime('%d/%m')} đến {end_date.strftime('%d/%m')}" if start_date != end_date else f"ngày {start_date.strftime('%d/%m')}"
    
    st.write(f"### 1. Số Liệu Doanh Thu {chuoi_hien_thi} (VNĐ)")
    col_dt1, col_dt2 = st.columns(2)
    
    with col_dt1:
        st.write("**💵 Doanh thu Offline (Bán tại chỗ/Mang đi)**")
        dt_offline_default = 0
        
        if 'nhat_ky_doanh_thu_offline' in st.session_state:
            nhat_ky = st.session_state['nhat_ky_doanh_thu_offline']
            dt_goi_y = sum([int(nhat_ky.get(d, 0)) for d in date_list])
            
            if dt_goi_y > 0:
                lua_chon = st.radio(f"Tìm thấy dữ liệu {chuoi_hien_thi} từ file iPOS:", 
                                     options=["Không, tôi tự nhập tay", f"Có, tự động lấy số {dt_goi_y:,} đ"])
                if lua_chon != "Không, tôi tự nhập tay":
                    dt_offline_default = dt_goi_y
                    st.success("✅ Đã tự động điền số liệu Offline.")
            else:
                st.warning(f"⚠️ File iPOS bên Tab 2 không chứa dữ liệu của {chuoi_hien_thi}.")
        
        dt_offline = st.number_input("Nhập/Chỉnh sửa Doanh thu Offline:", min_value=0, value=dt_offline_default, step=10000)

    with col_dt2:
        st.write("**🛵 Doanh thu Online (Grab, ShopeeFood...)**")
        link_gg_sheet = st.text_input("🔗 Link Google Sheet (Nhớ mở quyền 'Bất kỳ ai có liên kết'):")
        ten_tab = st.text_input("📌 Tên Tab (Sheet) chứa dữ liệu:", value="Lục_TS") 
        
        if st.button("🔄 Làm mới dữ liệu Google Sheet"):
            doc_du_lieu_gg_sheet.clear()
            st.toast("Đã dọn dẹp bộ nhớ đệm, hệ thống sẽ tải lại bản cập nhật mới nhất!")
            
        dt_online = 0
        
        if link_gg_sheet:
            try:
                with st.spinner("Đang tải và đồng bộ hóa dữ liệu..."):
                    df_onl = doc_du_lieu_gg_sheet(link_gg_sheet, ten_tab)
                    
                st.success(f"✅ Đã tải dữ liệu từ Tab `{ten_tab}`!")
                
                c1_onl, c2_onl = st.columns(2)
                with c1_onl:
                    cot_ngay_onl = st.selectbox("Chọn cột chứa Ngày (Ví dụ: cột BI):", options=["Không chọn"] + df_onl.columns.tolist())
                with c2_onl:
                    cot_dt_onl = st.selectbox("Chọn cột Doanh Thu (Ví dụ: cột BH):", options=["Không chọn"] + df_onl.columns.tolist())
                
                if cot_ngay_onl != "Không chọn" and cot_dt_onl != "Không chọn":
                    df_onl['Ngay_Chuan'] = pd.to_datetime(df_onl[cot_ngay_onl], dayfirst=True, errors='coerce').dt.date
                    df_onl_ngay = df_onl[(df_onl['Ngay_Chuan'] >= start_date) & (df_onl['Ngay_Chuan'] <= end_date)].copy()
                    
                    if not df_onl_ngay.empty:
                        df_onl_ngay[cot_dt_onl] = df_onl_ngay[cot_dt_onl].apply(clean_money)
                        dt_online_tong = int(df_onl_ngay[cot_dt_onl].sum())
                        
                        st.info(f"✅ Tổng Doanh thu Online {chuoi_hien_thi}: {dt_online_tong:,} đ")
                        dt_online = dt_online_tong
                    else:
                        st.warning(f"⚠️ Trong Tab `{ten_tab}` chưa có số liệu của {chuoi_hien_thi}")
                        dt_online = st.number_input("Nhập tay Doanh thu Online:", min_value=0, value=0, step=10000)
            except Exception as e:
                st.error(f"❌ Lỗi đọc Google Sheet: {e}")
                dt_online = st.number_input("Nhập tay Doanh thu Online:", min_value=0, value=0, step=10000)
        else:
            dt_online = st.number_input("Nhập Doanh thu Online:", min_value=0, value=0, step=10000)

    tong_dt_hien_tai = dt_offline + dt_online
    st.write("")
    st.success(f"🌟 **TỔNG DOANH THU (OFFLINE + ONLINE): {tong_dt_hien_tai:,} đ**")
    
    st.write("---")
    st.write(f"### 2. Số Liệu Kho Hàng {chuoi_hien_thi} (VNĐ)")
    st.caption("*(Nếu chọn khoảng ngày: Tồn đầu là của ngày đầu tiên, Tồn cuối là của ngày cuối cùng)*")
    col_k1, col_k2, col_k3 = st.columns(3)
    with col_k1:
        ton_dau = st.number_input("📦 Tổng giá trị Tồn Đầu", min_value=0, value=0, step=100000)
    with col_k2:
        nhap_trong_ngay = st.number_input("🛒 Tổng giá trị Nhập Hàng", min_value=0, value=0, step=100000)
    with col_k3:
        ton_cuoi = st.number_input("📉 Tổng giá trị Tồn Cuối", min_value=0, value=0, step=100000)

    if st.button(f"🧮 Chốt Số Tỷ Lệ % Nguyên Liệu {chuoi_hien_thi}"):
        tong_doanh_thu = dt_offline + dt_online
        chi_phi_nl = ton_dau + nhap_trong_ngay - ton_cuoi
        
        st.write("---")
        if tong_doanh_thu == 0:
            st.warning("⚠️ Tổng doanh thu đang bằng 0, không thể chia tỷ lệ % được. Anh vui lòng kiểm tra lại nhé!")
        elif chi_phi_nl < 0:
            st.error("❌ Số liệu kho đang bị âm (Tồn cuối lớn hơn tổng Tồn đầu + Nhập). Anh kiểm tra lại giá trị kiểm kho nhé!")
        else:
            phan_tram_nl = (chi_phi_nl / tong_doanh_thu) * 100
            
            c1, c2, c3 = st.columns(3)
            c1.metric("💰 Tổng Doanh Thu (Onl + Off)", f"{tong_doanh_thu:,} đ")
            c2.metric("🔪 Chi Phí Nguyên Liệu (Thực Tế)", f"{chi_phi_nl:,} đ")
            
            if phan_tram_nl <= 35:
                c3.metric("🎯 Tỷ lệ % Nguyên Liệu", f"{phan_tram_nl:.1f} %", "Tuyệt vời (≤35%)")
            elif phan_tram_nl <= 40:
                c3.metric("🎯 Tỷ lệ % Nguyên Liệu", f"{phan_tram_nl:.1f} %", "Chấp nhận được", delta_color="off")
            else:
                c3.metric("🎯 Tỷ lệ % Nguyên Liệu", f"{phan_tram_nl:.1f} %", "Báo Động Hụt Kho (>40%)", delta_color="inverse")
import streamlit as st
import pandas as pd
import io
import os
import re
import datetime
import json
import time

# ==========================================
# BỘ NHỚ LƯU TRỮ CẤU HÌNH (JSON)
# ==========================================
CONFIG_FILE = "config_links.json"

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_config(config_data):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config_data, f, ensure_ascii=False, indent=4)

app_config = load_config()

# ==========================================
# CẤU HÌNH TRANG & GIAO DIỆN (UI/UX)
# ==========================================
st.set_page_config(page_title="Hệ Thống Quản Trị F&B", layout="wide", page_icon="📊")

st.markdown("""
<style>
    .stApp { background-color: #FDFBF7; background-image: radial-gradient(#EFEBE5 1px, transparent 1px); background-size: 20px 20px; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    h1, h2, h3, h4, h5 { color: #5C4D42 !important; font-weight: 600 !important; letter-spacing: 0.5px; }
    hr { border: 0; height: 1px; background-image: linear-gradient(to right, rgba(196, 164, 132, 0), rgba(196, 164, 132, 0.6), rgba(196, 164, 132, 0)); margin: 2em 0; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; padding-bottom: 5px; border-bottom: 1px solid #EAEAEA; flex-wrap: wrap; }
    .stTabs [data-baseweb="tab"] { background-color: #F5EFEB; border-radius: 8px 8px 0px 0px; padding: 10px 20px; color: #8C7B6D; font-weight: 500; border: none; transition: all 0.3s ease; font-size: 15px; }
    .stTabs [data-baseweb="tab"]:hover { background-color: #EAE0D5; color: #5C4D42; }
    .stTabs [aria-selected="true"] { background-color: #FFFFFF !important; color: #4A4036 !important; font-weight: 700; border-bottom: 3px solid #D2B48C !important; box-shadow: 0 -3px 5px rgba(0,0,0,0.01); }
    div[data-testid="metric-container"] { background-color: #FFFFFF; border: 1px solid #F0F0F0; border-left: 4px solid #C4A484; padding: 20px; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.03); transition: transform 0.2s ease; }
    div[data-testid="metric-container"]:hover { transform: translateY(-2px); box-shadow: 0 6px 20px rgba(0,0,0,0.06); }
    .stButton>button { background-color: #6B5E53; color: #FFFFFF; border-radius: 6px; border: none; padding: 0.5rem 1rem; font-weight: 500; transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1); box-shadow: 0 2px 4px rgba(107, 94, 83, 0.2); }
    .stButton>button:hover { background-color: #4A4036; color: #FFFFFF; box-shadow: 0 4px 8px rgba(74, 64, 54, 0.3); }
    .stDataFrame { border-radius: 10px; overflow: hidden; border: 1px solid #EAEAEA; box-shadow: 0 2px 8px rgba(0,0,0,0.02); }
    .streamlit-expanderHeader { background-color: #FFFFFF; border-radius: 8px; border: 1px solid #EAEAEA; color: #5C4D42; font-weight: 500; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# HÀM XỬ LÝ DỮ LIỆU CHUẨN XÁC
# ==========================================
def clean_money(val):
    if pd.isna(val): return 0
    if isinstance(val, (int, float)): return int(val)
    val_str = str(val).strip()
    if not val_str or val_str in ['-', 'nan', '##########']: return 0
    val_str = re.sub(r'[^\d\.,\-]', '', val_str)
    val_str = re.sub(r'[\.,]\d{1,2}$', '', val_str)
    val_str = re.sub(r'[^\d\-]', '', val_str)
    if not val_str or val_str == '-': return 0
    return int(val_str)

def clean_date_robust(date_series):
    ds_str = date_series.astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    cleaned = pd.to_datetime(ds_str, format='%Y-%m-%d', errors='coerce')
    mask_na = cleaned.isna()
    if mask_na.any():
        cleaned.loc[mask_na] = pd.to_datetime(ds_str.loc[mask_na], format='%d/%m/%Y', errors='coerce')
    mask_na = cleaned.isna()
    if mask_na.any():
        cleaned.loc[mask_na] = pd.to_datetime(ds_str.loc[mask_na], errors='coerce', dayfirst=True)
    return cleaned.dt.normalize()

def clean_ipos_date(date_series):
    ds_str = date_series.astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    ds_date_part = ds_str.str.split().str[0]
    cleaned = pd.to_datetime(ds_date_part, format='%d/%m/%Y', errors='coerce')
    mask_na = cleaned.isna()
    if mask_na.any():
        cleaned.loc[mask_na] = pd.to_datetime(ds_date_part.loc[mask_na], format='%Y-%m-%d', errors='coerce')
    mask_na = cleaned.isna()
    if mask_na.any():
        cleaned.loc[mask_na] = pd.to_datetime(ds_str.loc[mask_na], errors='coerce', dayfirst=True)
    return cleaned.dt.normalize()

def process_dataframe_header(df_raw, tu_khoa_nhan_dien=[]):
    header_idx = 0
    if tu_khoa_nhan_dien:
        for i in range(min(15, len(df_raw))):
            row_vals = [str(x).lower().strip() for x in df_raw.iloc[i].values if pd.notna(x)]
            if any(kw.lower() in row_vals for kw in tu_khoa_nhan_dien):
                header_idx = i
                break
                
    cols = []
    counts = {}
    for i, val in enumerate(df_raw.iloc[header_idx]):
        if pd.isna(val) or str(val).strip() == '':
            col_name = f"Unnamed_{i}"
        else:
            col_name = str(val).strip()
            
        if col_name in counts:
            counts[col_name] += 1
            cols.append(f"{col_name}_{counts[col_name]}")
        else:
            counts[col_name] = 0
            cols.append(col_name)
            
    df_raw.columns = cols
    df_raw = df_raw.iloc[header_idx+1:].reset_index(drop=True)
    return df_raw

@st.cache_data(ttl=600, show_spinner=False)
def doc_sheet_thong_minh(link, ten_tab, tu_khoa_nhan_dien=[], refresh_trigger=0):
    file_id = re.search(r'/d/([a-zA-Z0-9-_]+)', link).group(1)
    xlsx_url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
    df_raw = pd.read_excel(xlsx_url, sheet_name=ten_tab, header=None)
    return process_dataframe_header(df_raw, tu_khoa_nhan_dien)

danh_sach_cn_he_thong = ["Trường Sa", "Lê Quang Định", "Trần Huy Liệu"]

# ==========================================
# THANH SIDEBAR TỰ ĐỘNG ĐỒNG BỘ ONLINE 
# ==========================================
st.sidebar.markdown("### ⚙️ CẤU HÌNH ĐỐI SOÁT ONLINE")
st.sidebar.info("Thiết lập Link 1 lần duy nhất ở đây. Bấm nút dưới để hệ thống kéo toàn bộ dữ liệu về chạy cục bộ siêu tốc.")

def_onl_link = app_config.get("online_master", {}).get("link", "")
global_onl_link = st.sidebar.text_input("🔗 Link Google Sheet (Tổng):", value=def_onl_link)

if st.sidebar.button("🚀 LƯU & CẬP NHẬT ĐỒNG LOẠT", use_container_width=True):
    if global_onl_link:
        if "online_master" not in app_config: app_config["online_master"] = {}
        app_config["online_master"]["link"] = global_onl_link
        save_config(app_config)
        
        with st.spinner("⏳ Đang quét tiêu đề và kéo dữ liệu từ mây cho 3 chi nhánh..."):
            for cn_mac_dinh, def_tab in [("Trường Sa", "Lục_TS"), ("Lê Quang Định", "Lục_LQD"), ("Trần Huy Liệu", "Lục_THL")]:
                tab_name = app_config.get("food_cost", {}).get(cn_mac_dinh, {}).get("tab", def_tab)
                try:
                    df_onl = doc_sheet_thong_minh(global_onl_link, tab_name, ['theo ngày', 'doanh thu', 'tổng doanh thu'], time.time())
                    df_onl.to_csv(f"temp_onl_{cn_mac_dinh}.csv", index=False)
                except Exception as e:
                    st.sidebar.error(f"❌ Lỗi tải {cn_mac_dinh} (Sheet '{tab_name}'): Không tìm thấy hoặc link sai.")
                    
        st.sidebar.success("✅ Đã lấy số xong! Các Tab hiện tại đã có dữ liệu chuẩn xác.")
        st.rerun()
    else:
        st.sidebar.warning("⚠️ Vui lòng nhập link trước khi đồng bộ.")

st.title("📊 Hệ Thống Bóc Tách & Phân Tích Dữ Liệu F&B")

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📦 Xử Lý Tồn Kho", 
    "📈 Bán Hàng iPOS", 
    "📋 Quản Lý Menu", 
    "🧮 % Trường Sa", 
    "🧮 % Trần Huy Liệu", 
    "🧮 % Lê Quang Định"
])

if "refresh_kho" not in st.session_state: st.session_state["refresh_kho"] = 0
if "refresh_ipos" not in st.session_state: st.session_state["refresh_ipos"] = {cn: 0 for cn in danh_sach_cn_he_thong}

# ==========================================
# TAB 1: XỬ LÝ TỒN KHO 
# ==========================================
with tab1:
    st.markdown("### 📦 Quản Lý Dữ Liệu Tồn Kho")
    st.info("Dữ liệu sau khi tải lên sẽ được hệ thống lưu lại tạm thời. Anh có thể F5 thoải mái mà không sợ mất số liệu.")
    
    loai_nguon_kho = st.radio("Nguồn cấp dữ liệu Tồn Kho:", 
                              options=["🔗 Liên kết Google Sheet (Dữ liệu nền AppSheet)", "📁 Tải file (Excel/CSV) thủ công"], horizontal=True)
    
    df_kho_goc = pd.DataFrame()
    
    if loai_nguon_kho == "🔗 Liên kết Google Sheet (Dữ liệu nền AppSheet)":
        col_link, col_tab, col_btn = st.columns([3, 1, 1])
        with col_link:
            def_kho_link = app_config.get("kho", {}).get("link", "")
            link_appsheet = st.text_input("🔗 Nhập link Google Sheet (AppSheet):", value=def_kho_link)
        with col_tab:
            def_kho_tab = app_config.get("kho", {}).get("tab", "Phiếu Kiểm Kho")
            tab_appsheet = st.text_input("📌 Tên Tab:", value=def_kho_tab)
        with col_btn:
            st.write("") 
            st.write("")
            if st.button("🔄 Lấy Dữ Liệu", key="btn_kho_sync"):
                if "kho" not in app_config: app_config["kho"] = {}
                app_config["kho"]["link"] = link_appsheet
                app_config["kho"]["tab"] = tab_appsheet
                save_config(app_config)
                st.session_state["refresh_kho"] += 1
            
        if link_appsheet:
            try:
                with st.spinner("⏳ Đang kéo dữ liệu Tồn Kho từ mây..."):
                    df_kho_goc = doc_sheet_thong_minh(link_appsheet, tab_appsheet, [], st.session_state["refresh_kho"])
                st.success(f"✅ Đã kết nối thành công với kho dữ liệu `{tab_appsheet}`!")
            except Exception as e:
                st.error("❌ Không thể kết nối. Vui lòng kiểm tra lại link Google Sheet.")
                
    else:
        if os.path.exists("temp_kho.csv"):
            df_kho_goc = pd.read_csv("temp_kho.csv")
            st.success("✅ Đã khôi phục dữ liệu Tồn Kho từ phiên làm việc trước.")
            if st.button("🗑️ Xóa dữ liệu cũ để tải file mới"):
                os.remove("temp_kho.csv")
                st.rerun()
        else:
            file_kho = st.file_uploader("Tải file Tồn Kho lên đây", type=['xlsx', 'xls', 'csv'], key="kho")
            if file_kho is not None:
                df_raw = pd.read_csv(file_kho, header=None) if file_kho.name.endswith('.csv') else pd.read_excel(file_kho, header=None)
                df_kho_goc = process_dataframe_header(df_raw, [])
                df_kho_goc.to_csv("temp_kho.csv", index=False)
                st.rerun()

    if not df_kho_goc.empty:
        st.session_state['df_kho_goc'] = df_kho_goc.copy()
        df_kho = df_kho_goc.copy()
        
        st.write("---")
        st.write("### 🎛️ BỘ LỌC TÙY CHỈNH (CHI NHÁNH & THỜI GIAN)")
        
        cols = df_kho.columns.tolist()
        col_f1, col_f2 = st.columns(2)
        
        with col_f1:
            cot_chi_nhanh = "Chi nhánh" if "Chi nhánh" in cols else ("Chi Nhánh" if "Chi Nhánh" in cols else None)
            if cot_chi_nhanh:
                danh_sach_cn = df_kho[cot_chi_nhanh].dropna().unique().tolist()
                cn_chon = st.multiselect("🏢 Chọn Chi Nhánh Cần Phân Tích:", options=danh_sach_cn, default=danh_sach_cn)
                if cn_chon:
                    df_kho = df_kho[df_kho[cot_chi_nhanh].isin(cn_chon)]
            else:
                st.info("Bảng không có cột định danh 'Chi nhánh'")
                
        with col_f2:
            date_cols = [c for c in df_kho.columns if 'ngày' in c.lower() or 'date' in c.lower() or 'thời gian' in c.lower()]
            if date_cols:
                date_col = date_cols[0]
                df_kho['_temp_date'] = clean_date_robust(df_kho[date_col])
                df_kho['_temp_date_obj'] = df_kho['_temp_date'].dt.date
                
                min_date = df_kho['_temp_date_obj'].min()
                max_date = df_kho['_temp_date_obj'].max()
                
                if pd.notnull(min_date) and pd.notnull(max_date):
                    ngay_loc = st.date_input(f"🗓️ Lọc khoảng thời gian ({date_col}):", value=(min_date, max_date), min_value=min_date, max_value=max_date)
                    if isinstance(ngay_loc, tuple):
                        start_d = ngay_loc[0]
                        end_d = ngay_loc[1] if len(ngay_loc) > 1 else start_d
                    else:
                        start_d = end_d = ngay_loc
                    
                    df_kho = df_kho[(df_kho['_temp_date_obj'] >= start_d) & (df_kho['_temp_date_obj'] <= end_d)]
                df_kho = df_kho.drop(columns=['_temp_date', '_temp_date_obj'])

        st.write("🔍 **Bản xem trước dữ liệu (Sau khi lọc):**")
        st.dataframe(df_kho.head(5), use_container_width=True)

# ==========================================
# TAB 2: PHÂN TÍCH BÁN HÀNG CHI TIẾT
# ==========================================
with tab2:
    st.markdown("### 📈 Phân Tích Bán Hàng iPOS (Đa Chi Nhánh)")
    st.info("Sử dụng Google Sheet để đảm bảo dữ liệu bán hàng được lưu trữ vĩnh viễn, không bị mất khi ứng dụng khởi động lại.")
    
    with st.expander("📥 QUẢN LÝ DỮ LIỆU IPOS TỪNG CHI NHÁNH", expanded=True):
        loai_nguon_ipos = st.radio("Cấp dữ liệu báo cáo iPOS bằng cách:", 
                                   options=["🔗 Liên kết Google Sheet (Lưu Vĩnh Viễn)", "📁 Tải file thủ công (Bị mất khi qua ngày)"], horizontal=True)
        st.write("---")
        
        if loai_nguon_ipos == "🔗 Liên kết Google Sheet (Lưu Vĩnh Viễn)":
            st.write("**1. Cấu hình nguồn dữ liệu iPOS chung cho toàn hệ thống**")
            def_ipos_link = app_config.get("ipos_master", {}).get("link", "")
            link_ipos = st.text_input("🔗 Link Google Sheet (Dùng chung cho cả 3 chi nhánh):", value=def_ipos_link)
            
            st.write("**2. Khai báo Tên Tab (Sheet) tương ứng với từng chi nhánh:**")
            col_t1, col_t2, col_t3 = st.columns(3)
            with col_t1: tab_ts = st.text_input("Trường Sa:", value=app_config.get("ipos_master", {}).get("tab_TS", "iPOS_Trường Sa"))
            with col_t2: tab_lqd = st.text_input("Lê Quang Định:", value=app_config.get("ipos_master", {}).get("tab_LQD", "iPOS_Lê Quang Định"))
            with col_t3: tab_thl = st.text_input("Trần Huy Liệu:", value=app_config.get("ipos_master", {}).get("tab_THL", "iPOS_Trần Huy Liệu"))
            
            if st.button("🚀 ĐỒNG BỘ CẢ 3 CHI NHÁNH TỪ IPOS", use_container_width=True):
                if "ipos_master" not in app_config: app_config["ipos_master"] = {}
                app_config["ipos_master"]["link"] = link_ipos
                app_config["ipos_master"]["tab_TS"] = tab_ts
                app_config["ipos_master"]["tab_LQD"] = tab_lqd
                app_config["ipos_master"]["tab_THL"] = tab_thl
                save_config(app_config)
                
                if link_ipos:
                    tabs_map = {"Trường Sa": tab_ts, "Lê Quang Định": tab_lqd, "Trần Huy Liệu": tab_thl}
                    with st.spinner("⏳ Hệ thống đang quét và đồng bộ dữ liệu iPOS của cả 3 chi nhánh từ mây..."):
                        for cn_upload, tab_name in tabs_map.items():
                            try:
                                df_new = doc_sheet_thong_minh(link_ipos, tab_name, ['tên hàng', 'mã hàng'])
                                if df_new.empty:
                                    st.warning(f"⚠️ {cn_upload} (Tab: {tab_name}): Không có dữ liệu hợp lệ.")
                                else:
                                    df_new['Chi Nhánh Hệ Thống'] = cn_upload
                                    df_new.to_csv(f"temp_ipos_{cn_upload}.csv", index=False)
                            except Exception as e:
                                st.error(f"❌ Lỗi tải {cn_upload} (Tab '{tab_name}'): {e}")
                        st.success("✅ Đã hoàn tất đồng bộ toàn hệ thống!")
                        st.rerun()

    df_ban_goc_list = []
    for cn in danh_sach_cn_he_thong:
        file_path = f"temp_ipos_{cn}.csv"
        if os.path.exists(file_path):
            try:
                if os.path.getsize(file_path) > 0:
                    df_temp = pd.read_csv(file_path)
                    if not df_temp.empty:
                        df_ban_goc_list.append(df_temp)
            except Exception:
                pass
            
    if df_ban_goc_list:
        df_ban_master = pd.concat(df_ban_goc_list, ignore_index=True)
        col_ten_hang = next((c for c in df_ban_master.columns if 'tên hàng' in str(c).lower()), None)
        if not col_ten_hang: col_ten_hang = next((c for c in df_ban_master.columns if 'mã hàng' in str(c).lower()), None)
        
        if col_ten_hang:
            df_ban = df_ban_master.dropna(subset=[col_ten_hang]).copy()
            df_ban = df_ban[~df_ban[col_ten_hang].astype(str).str.strip().isin(['-', 'Tổng cộng'])]
            
            col_sl = next((c for c in df_ban.columns if 'số lượng' in str(c).lower()), 'Số lượng')
            col_tt = None
            if len(df_ban.columns) > 38 and 'tiền' in str(df_ban.columns[38]).lower():
                col_tt = df_ban.columns[38]
            else:
                for c in reversed(df_ban.columns.tolist()):
                    if str(c).strip().lower() == 'tổng tiền' or str(c).strip().lower().startswith('tổng tiền'):
                        col_tt = c
                        break
            if not col_tt:
                col_tt = next((c for c in reversed(df_ban.columns.tolist()) if 'tổng tiền' in str(c).lower() or 'doanh thu' in str(c).lower()), 'Tổng tiền')
            
            time_col_master = df_ban.columns[12] if len(df_ban.columns) > 12 else next((c for c in df_ban.columns if 'thời gian' in str(c).lower() or 'ngày' in str(c).lower()), None)
            
            if col_sl in df_ban.columns: df_ban['Số lượng'] = pd.to_numeric(df_ban[col_sl], errors='coerce').fillna(0)
            if col_tt in df_ban.columns: df_ban['Tổng tiền_Clean'] = df_ban[col_tt].apply(clean_money)
            
            if 'Tổng tiền_Clean' in df_ban.columns and time_col_master:
                df_ban['Date_Obj'] = clean_ipos_date(df_ban[time_col_master])
                df_ban = df_ban[df_ban['Date_Obj'].notna()]
                
                df_ban['Ngày_Chuan_Str'] = pd.to_datetime(df_ban['Date_Obj']).dt.strftime('%d/%m/%Y')
                dict_doanh_thu = {}
                for cn in df_ban['Chi Nhánh Hệ Thống'].unique():
                    df_cn_sub = df_ban[df_ban['Chi Nhánh Hệ Thống'] == cn]
                    dict_doanh_thu[cn] = df_cn_sub.groupby('Ngày_Chuan_Str')['Tổng tiền_Clean'].sum().to_dict()
                st.session_state['nhat_ky_doanh_thu_offline_multi'] = dict_doanh_thu

        st.write("---")
        st.write("### 🎛️ BỘ LỌC PHÂN TÍCH IPOS")
        col_f_ipos1, col_f_ipos2 = st.columns(2)
        with col_f_ipos1:
            cn_chon_ipos = st.multiselect("🏢 Lọc xem báo cáo theo Chi nhánh:", options=danh_sach_cn_he_thong, default=danh_sach_cn_he_thong)
            
        start_d_ipos, end_d_ipos = None, None
        with col_f_ipos2:
            if 'df_ban' in locals() and 'Date_Obj' in df_ban.columns and not df_ban['Date_Obj'].empty:
                min_date_ipos = df_ban['Date_Obj'].dropna().min()
                max_date_ipos = df_ban['Date_Obj'].dropna().max()
                if pd.notnull(min_date_ipos) and pd.notnull(max_date_ipos):
                    min_d_ipos = min_date_ipos.date()
                    max_d_ipos = max_date_ipos.date()
                    ngay_loc_ipos = st.date_input("🗓️ Lọc khoảng thời gian phân tích:", value=(min_d_ipos, max_d_ipos), min_value=min_d_ipos, max_value=max_d_ipos, key="date_range_ipos_master")
                    if isinstance(ngay_loc_ipos, tuple):
                        start_d_ipos = ngay_loc_ipos[0]
                        end_d_ipos = ngay_loc_ipos[1] if len(ngay_loc_ipos) > 1 else start_d_ipos
                    else:
                        start_d_ipos = end_d_ipos = ngay_loc_ipos

        if 'df_ban' in locals() and cn_chon_ipos and col_ten_hang:
            df_ban_view = df_ban[df_ban['Chi Nhánh Hệ Thống'].isin(cn_chon_ipos)].copy()
            if start_d_ipos and end_d_ipos and 'Date_Obj' in df_ban_view.columns:
                df_ban_view = df_ban_view[(df_ban_view['Date_Obj'] >= pd.to_datetime(start_d_ipos)) & (df_ban_view['Date_Obj'] <= pd.to_datetime(end_d_ipos))]
            
            df_ban_view['Size'] = df_ban_view[col_ten_hang].apply(lambda name: re.search(r'\((S|M|L)\)$', str(name).strip(), re.IGNORECASE).group(1).upper() if re.search(r'\((S|M|L)\)$', str(name).strip(), re.IGNORECASE) else 'Không ghi rõ')
            df_ban_view['Tên món gốc'] = df_ban_view[col_ten_hang].apply(lambda name: re.sub(r'\s*\((S|M|L)\)$', '', str(name).strip(), flags=re.IGNORECASE).strip())

            st.write("---")
            st.subheader("🏆 TỔNG QUAN KINH DOANH")
            col1, col2 = st.columns(2)
            if 'Số lượng' in df_ban_view.columns: col1.metric("🥤 Tổng Ly Bán Ra", f"{df_ban_view['Số lượng'].sum():,.0f} ly")
            if 'Tổng tiền_Clean' in df_ban_view.columns: col2.metric("💰 Tổng Doanh Thu", f"{df_ban_view['Tổng tiền_Clean'].sum():,.0f} VNĐ")

            if 'Date_Obj' in df_ban_view.columns and 'Tổng tiền_Clean' in df_ban_view.columns:
                st.write("---")
                st.subheader("📈 1. Biểu Đồ Doanh Thu Theo Ngày")
                df_chart_data = df_ban_view.dropna(subset=['Date_Obj'])
                if not df_chart_data.empty:
                    df_trend = df_chart_data.groupby(['Date_Obj', 'Chi Nhánh Hệ Thống'])['Tổng tiền_Clean'].sum().reset_index()
                    df_pivot_trend = df_trend.pivot(index='Date_Obj', columns='Chi Nhánh Hệ Thống', values='Tổng tiền_Clean').fillna(0)
                    df_pivot_trend.index = df_pivot_trend.index.date
                    st.line_chart(df_pivot_trend)

# ==========================================
# TAB 3: QUẢN LÝ MENU GỐC
# ==========================================
with tab3:
    st.markdown("### 📋 Dữ Liệu Bảng Giá (Menu)")
    if os.path.exists("menu_goc.csv"):
        df_menu = pd.read_csv("menu_goc.csv")
        st.success("✅ Đã khôi phục Menu gốc.")
        st.dataframe(df_menu, use_container_width=True)
        if st.button("🗑️ Xóa Menu hiện tại"):
            os.remove("menu_goc.csv")
            st.rerun()

# ==========================================
# HÀM LÕI KÉO DỮ LIỆU FOOD COST & NHẬP HÀNG TỪ SHEET "Nhập Nguyên Liệu"
# ==========================================
def render_food_cost_tab(cn_mac_dinh, prefix_key, default_tab_sheet):
    st.markdown(f"### 🧮 Quản Trị Tỷ Lệ % Nguyên Liệu (Food Cost) - {cn_mac_dinh}")
    st.info("Hệ thống tự động lấp dữ liệu Tồn Kho, Doanh Thu iPOS và Chi Phí Nhập Hàng dựa theo chi nhánh và ngày chọn.")
    
    col_filter1, col_filter2 = st.columns(2)
    with col_filter1:
        today = datetime.date.today()
        ngay_tinh = st.date_input("🎯 Chọn ngày kiểm toán:", value=(today, today), key=f"date_{prefix_key}")
        if isinstance(ngay_tinh, tuple):
            start_date = ngay_tinh[0]
            end_date = ngay_tinh[1] if len(ngay_tinh) > 1 else start_date
        else:
            start_date = end_date = ngay_tinh
            
        date_list = pd.date_range(start_date, end_date).strftime('%d/%m/%Y').tolist()
        chuoi_hien_thi = f"từ {start_date.strftime('%d/%m')} đến {end_date.strftime('%d/%m')}" if start_date != end_date else f"ngày {start_date.strftime('%d/%m')}"
        
    with col_filter2:
        danh_sach_cn_tab = danh_sach_cn_he_thong.copy()
        if cn_mac_dinh in danh_sach_cn_tab:
            danh_sach_cn_tab.remove(cn_mac_dinh)
            danh_sach_cn_tab.insert(0, cn_mac_dinh)
            
        cn_doisoat = st.selectbox("🏢 Chọn Chi nhánh đối chiếu:", options=danh_sach_cn_tab, key=f"cn_{prefix_key}")

    # Tự động lấy Tồn Đầu & Tồn Cuối từ Tab 1 (Kho)
    ton_dau_auto = 0
    ton_cuoi_auto = 0
    if 'df_kho_goc' in st.session_state and not st.session_state['df_kho_goc'].empty:
        df_k = st.session_state['df_kho_goc'].copy()
        cols_k = df_k.columns.tolist()
        cot_cn = next((c for c in cols_k if 'chi nhánh' in str(c).lower()), None)
        cot_ngay = next((c for c in cols_k if 'ngày' in str(c).lower() or 'date' in str(c).lower()), None)
        cot_gt = next((c for c in cols_k if 'giá trị' in str(c).lower() or 'tổng' in str(c).lower()), None)
        
        if cot_cn and cot_ngay and cot_gt:
            df_k_cn = df_k[df_k[cot_cn].astype(str).str.strip() == cn_doisoat].copy()
            df_k_cn['Date_Obj'] = clean_date_robust(df_k_cn[cot_ngay])
            df_k_cn[cot_gt] = df_k_cn[cot_gt].apply(clean_money)
            
            td_df = df_k_cn[df_k_cn['Date_Obj'] == pd.to_datetime(start_date)]
            if not td_df.empty: ton_dau_auto = td_df[cot_gt].sum()
                
            tc_df = df_k_cn[df_k_cn['Date_Obj'] == (pd.to_datetime(end_date) + datetime.timedelta(days=1))]
            if not tc_df.empty: ton_cuoi_auto = tc_df[cot_gt].sum()

    # Tự động tính tiền Nhập Hàng từ Sheet "Nhập Nguyên Liệu" (Cột A: Chi nhánh, Cột B: Ngày, Cột N: Tổng Cộng)
    nhap_hang_auto = 0
    link_master = app_config.get("online_master", {}).get("link", "")
    if link_master:
        try:
            df_nhap = doc_sheet_thong_minh(link_master, "Nhập Nguyên Liệu", ['chi nhánh', 'tổng cộng'])
            if not df_nhap.empty and len(df_nhap.columns) >= 14:
                col_cn_nhap = df_nhap.columns[0]
                col_ngay_nhap = df_nhap.columns[1]
                col_tong_nhap = df_nhap.columns[13] # Cột N là cột thứ 14 (index 13)
                
                df_nhap['Date_Obj'] = clean_date_robust(df_nhap[col_ngay_nhap])
                df_nhap['Tien_Nhap'] = df_nhap[col_tong_nhap].apply(clean_money)
                
                start_ts = pd.to_datetime(start_date)
                end_ts = pd.to_datetime(end_date)
                
                df_nhap_loc = df_nhap[
                    (df_nhap[col_cn_nhap].astype(str).str.strip() == cn_doisoat) & 
                    (df_nhap['Date_Obj'] >= start_ts) & 
                    (df_nhap['Date_Obj'] <= end_ts)
                ]
                nhap_hang_auto = int(df_nhap_loc['Tien_Nhap'].sum())
        except Exception:
            pass # Nếu chưa tạo sheet Nhập Nguyên Liệu thì bỏ qua không lỗi app

    st.write(f"#### 1. Doanh Thu {chuoi_hien_thi}")
    col_dt1, col_dt2 = st.columns(2)
    
    with col_dt1:
        st.write("**💵 Doanh thu Offline (tại chỗ)**")
        dt_offline_default = 0
        if 'nhat_ky_doanh_thu_offline_multi' in st.session_state:
            nhat_ky_tong = st.session_state['nhat_ky_doanh_thu_offline_multi']
            if cn_doisoat in nhat_ky_tong:
                nhat_ky_cn = nhat_ky_tong[cn_doisoat]
                dt_goi_y = sum([int(nhat_ky_cn.get(d, 0)) for d in date_list])
                if dt_goi_y > 0: dt_offline_default = dt_goi_y
        
        dk = f"{prefix_key}_{cn_doisoat}_{start_date}_{end_date}"
        dt_offline = st.number_input("Chỉnh sửa Doanh thu Offline:", min_value=0, value=int(dt_offline_default), step=10000, key=f"dtoff_{dk}")

    with col_dt2:
        st.write("**🛵 Doanh thu Online (ShopeeFood, Grab...)**")
        dt_online = 0
        file_onl = f"temp_onl_{cn_mac_dinh}.csv"
        if os.path.exists(file_onl):
            try:
                df_onl = pd.read_csv(file_onl)
                df_onl.columns = [str(c).strip() for c in df_onl.columns]
                col_ngay_onl = next((c for c in df_onl.columns if 'theo ngày' in str(c).lower()), None)
                col_dt_onl = next((c for c in df_onl.columns if 'tổng doanh thu' in str(c).lower()), None)
                
                if col_ngay_onl and col_dt_onl:
                    df_onl['Ngay_Chuan'] = clean_date_robust(df_onl[col_ngay_onl])
                    df_onl_ngay = df_onl[(df_onl['Ngay_Chuan'] >= pd.to_datetime(start_date)) & (df_onl['Ngay_Chuan'] <= pd.to_datetime(end_date))]
                    if not df_onl_ngay.empty:
                        dt_online = int(df_onl_ngay[col_dt_onl].apply(clean_money).sum())
            except Exception:
                pass
        dt_online_input = st.number_input("Chỉnh sửa Doanh thu Online:", min_value=0, value=int(dt_online), step=10000, key=f"dtonl_edit_{dk}")

    tong_dt_hien_tai = dt_offline + dt_online_input
    st.write(f"##### 🌟 TỔNG CỘNG DOANH THU: <span style='color:#C4A484;'>{tong_dt_hien_tai:,} VNĐ</span>", unsafe_allow_html=True)
    
    st.write("---")
    st.write(f"#### 2. Dữ Liệu Kho {chuoi_hien_thi}")
    if ton_dau_auto > 0 or ton_cuoi_auto > 0:
        st.success(f"🤖 Đã tự động điền Tồn Đầu, Nhập Hàng và Tồn Cuối của CN {cn_doisoat} từ Google Sheet.")
        
    col_k1, col_k2, col_k3 = st.columns(3)
    with col_k1:
        ton_dau = st.number_input("📦 Tồn Đầu (VNĐ)", min_value=0, value=int(ton_dau_auto), step=100000, key=f"td_{dk}")
    with col_k2:
        nhap_trong_ngay = st.number_input("🛒 Nhập Hàng (VNĐ)", min_value=0, value=int(nhap_hang_auto), step=100000, key=f"nhap_{dk}")
    with col_k3:
        ton_cuoi = st.number_input("📉 Tồn Cuối (VNĐ)", min_value=0, value=int(ton_cuoi_auto), step=100000, key=f"tc_{dk}")

    if st.button(f"🧮 Báo Cáo % Nguyên Liệu {chuoi_hien_thi}", key=f"btn_bc_{prefix_key}"):
        tong_doanh_thu = tong_dt_hien_tai
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

# ==========================================
# KHỞI TẠO TAB 4, 5, 6
# ==========================================
with tab4: render_food_cost_tab("Trường Sa", "TS", "Lục_TS")
with tab5: render_food_cost_tab("Trần Huy Liệu", "THL", "Lục_THL")
with tab6: render_food_cost_tab("Lê Quang Định", "LQD", "Lục_LQD")

import streamlit as st
import pandas as pd
import os
import base64
from datetime import datetime, timedelta
import qrcode
from io import BytesIO
from PIL import Image

# ==========================================
# 1. CẤU HÌNH HỆ THỐNG & GIAO DIỆN
# ==========================================
st.set_page_config(page_title="Dược Khoa Y NTTU", layout="wide", page_icon="🏥", initial_sidebar_state="collapsed")

# Danh sách file hệ thống
FILE_KHO = 'kho_thuoc_tong.csv'
FILE_LICH_SU = 'lich_su_chi_tiet.csv'
FILE_NHAN_SU = 'danh_sach_nhan_su.csv'
FILE_CHUONG_TRINH = 'danh_sach_chuong_trinh.csv'
FILE_NHOM_THUOC = 'danh_sach_nhom.csv' 
FILE_DU_TRU = 'du_tru_thuoc.csv'
FILE_BG = 'background_custom.png'
FILE_SESSION = 'session_token.txt'
FILE_COLOR = 'theme_color.txt'
FOLDER_AVATAR = 'avatars'

# Định nghĩa cột chuẩn
BASE_COLS_KHO = ['Barcode', 'Tên Biệt Dược', 'Chương Trình', 'Nhóm Thuốc', 'Thành Phần', 'Đơn Vị Tính', 'Hạn Sử Dụng', 'Nhập Mới', 'Đã Xuất']
BASE_COLS_NS = ['Username', 'Password', 'Quyền', 'Họ Tên', 'SĐT', 'Gmail', 'MSSV', 'Lớp']

if not os.path.exists(FOLDER_AVATAR): os.makedirs(FOLDER_AVATAR)
if not os.path.exists(FILE_COLOR):
    with open(FILE_COLOR, "w") as f: f.write("#004a99")
with open(FILE_COLOR, "r") as f: main_color = f.read().strip()

# ==========================================
# 2. HÀM HỖ TRỢ (UI, IN ẤN, LOGIC)
# ==========================================
def apply_styles(color, bg_path=None):
    bg_css = ""
    if bg_path and os.path.exists(bg_path):
        try:
            with open(bg_path, 'rb') as f: bin_str = base64.b64encode(f.read()).decode()
            bg_css = f'background-image: url("data:image/png;base64,{bin_str}"); background-size: cover; background-attachment: fixed;'
        except: pass
    st.markdown(f'''
    <style>
    .stApp {{ {bg_css} }}
    .main .block-container {{ background-color: rgba(255, 255, 255, 0.98); border-radius: 15px; padding: 30px; box-shadow: 0 4px 20px rgba(0,0,0,0.2); }}
    .stButton>button {{ background-color: {color}; color: white; border-radius: 8px; font-weight: bold; width: 100%; height: 45px; border: none; transition: 0.3s; }}
    .stButton>button:hover {{ background-color: white; color: {color}; border: 2px solid {color}; }}
    div[data-testid="stMetricValue"] {{ color: {color}; font-weight: bold; font-size: 32px; }}
    .stTabs [aria-selected="true"] {{ background-color: {color} !important; color: white !important; border-radius: 5px; }}
    </style>
    ''', unsafe_allow_html=True)

# Hàm kiểm tra trạng thái hạn sử dụng (Chuẩn hóa định dạng dd/mm/yyyy)
def check_hsd_status(hsd_str):
    try:
        # Ép kiểu về string và xóa khoảng trắng thừa
        hsd_str = str(hsd_str).strip()
        # Xử lý trường hợp Excel tự đổi sang dạng timestamp dài
        if " " in hsd_str: hsd_str = hsd_str.split(" ")[0]
        
        hsd_date = datetime.strptime(hsd_str, '%d/%m/%Y')
        today = datetime.now()
        if hsd_date < today:
            return "❌ Hết hạn"
        elif hsd_date <= today + timedelta(days=180): # Sắp hết hạn trong 6 tháng
            return "⚠️ Sắp hết hạn"
        else:
            return "✅ Còn hạn"
    except:
        return "❓ Lỗi định dạng"

# Hàm tô màu trạng thái HSD
def color_hsd(val):
    if val == "❌ Hết hạn": color = '#ff4b4b'
    elif val == "⚠️ Sắp hết hạn": color = '#ffa500'
    elif val == "✅ Còn hạn": color = '#28a745'
    else: color = '#6c757d'
    return f'color: {color}; font-weight: bold'

def create_print_html(df, title):
    u = st.session_state.u_data
    html = f"""
    <div style="font-family: Arial; padding: 25px; border: 2px solid {main_color}; border-radius: 10px; background: white; color: black;">
        <h2 style="text-align: center; color: {main_color}; text-transform: uppercase;">{title}</h2>
        <p style="text-align: center; font-style: italic;">Khoa Y - Trường Đại học Nguyễn Tất Thành</p>
        <hr style="border: 1px solid {main_color};">
        <p style="text-align: right; font-size: 12px;">Ngày xuất: {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
        <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
            <tr style="background-color: #f2f2f2;">{''.join([f'<th style="border: 1px solid #ddd; padding: 10px; text-align: left;">{c}</th>' for c in df.columns])}</tr>
            {''.join(['<tr>' + ''.join([f'<td style="border: 1px solid #ddd; padding: 8px;">{v}</td>' for v in r]) + '</tr>' for r in df.values])}
        </table>
        <div style="margin-top: 40px; display: flex; justify-content: space-between;">
            <div style="text-align: center; width: 45%;"><b>Người lập biểu</b><br><br><br>{u.get('Họ Tên')}<br>(MSSV: {u.get('MSSV')})</div>
            <div style="text-align: center; width: 45%;"><b>Xác nhận đơn vị</b><br><br><br>......................................</div>
        </div>
    </div>
    <button onclick="window.print()" style="margin-top:15px; padding:12px 25px; background:{main_color}; color:white; border:none; border-radius:5px; cursor:pointer; font-weight:bold;">🖨️ XÁC NHẬN IN / XUẤT PDF</button>
    """
    return html

def create_qr_pdf_html(df_nhom_thuoc, ten_nhom):
    qr_cards = ""
    for _, row in df_nhom_thuoc.iterrows():
        val = row['Barcode']
        ten = row['Tên Biệt Dược']
        qr = qrcode.make(val)
        b = BytesIO()
        qr.save(b, format="PNG")
        qr_base64 = base64.b64encode(b.getvalue()).decode()
        qr_cards += f"""
        <div style="width: 30%; border: 1px solid #ccc; padding: 10px; margin: 5px; display: inline-block; text-align: center; border-radius: 5px;">
            <img src="data:image/png;base64,{qr_base64}" style="width: 100px; height: 100px;"><br>
            <b style="font-size: 14px;">{ten}</b><br>
            <span style="font-size: 12px;">Mã: {val}</span><br>
            <span style="font-size: 10px;">({ten_nhom})</span>
        </div>
        """
    html = f"""
    <div style="font-family: Arial; padding: 20px; background: white; color: black;">
        <h3 style="text-align: center;">DANH SÁCH MÃ QR - NHÓM: {ten_nhom.upper()}</h3>
        <hr>
        <div style="display: flex; flex-wrap: wrap; justify-content: center;">{qr_cards}</div>
    </div>
    <button onclick="window.print()" style="margin-top:20px; width: 100%; padding:15px; background:{main_color}; color:white; border:none; border-radius:5px; cursor:pointer; font-weight:bold;">🖨️ XÁC NHẬN IN NHÃN QR (PDF)</button>
    """
    return html

def get_excel_template():
    df_temp = pd.DataFrame(columns=BASE_COLS_KHO)
    example = {'Barcode': 'AUTO', 'Tên Biệt Dược': 'Paracetamol 500mg', 'Chương Trình': 'Kho Tổng', 'Nhóm Thuốc': 'Giảm đau', 'Thành Phần': 'Paracetamol', 'Đơn Vị Tính': 'Viên', 'Hạn Sử Dụng': '31/12/2026', 'Nhập Mới': 100, 'Đã Xuất': 0}
    df_temp = pd.concat([df_temp, pd.DataFrame([example])], ignore_index=True)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_temp.to_excel(writer, index=False, sheet_name='Template')
    return output.getvalue()

# ==========================================
# 3. QUẢN LÝ DỮ LIỆU & LOGIC THÔNG MINH
# ==========================================

# XÓA BỎ HÀM SAVE_ALL CŨ VÀ DÁN ĐOẠN NÀY VÀO:
def save_all():
    # 1. Sắp xếp lại kho thuốc theo tên Thành Phần trước khi lưu
    if not st.session_state.df_kho.empty:
        st.session_state.df_kho = st.session_state.df_kho.sort_values(by='Thành Phần', ascending=True)
    
    # 2. ĐỒNG BỘ DỮ LIỆU LÊN CLOUD (GOOGLE SHEETS)
    # LƯU Ý: Đã xóa bỏ hoàn toàn các dòng .to_csv() để tránh lỗi Permission Denied (Lỗi 13)
    try:
        # Cập nhật bảng Kho Thuốc
        conn.update(worksheet="KhoThuoc", data=st.session_state.df_kho)
        
        # Cập nhật bảng Lịch Sử Xuất Thuốc
        conn.update(worksheet="LichSu", data=st.session_state.df_ls)
        
        # Cập nhật bảng Nhân Sự (Xử lý an toàn để tránh lỗi thiếu cột)
        cols_ns_safe = [c for c in BASE_COLS_NS if c in st.session_state.df_ns.columns]
        conn.update(worksheet="NhanSu", data=st.session_state.df_ns[cols_ns_safe])
        
        # Cập nhật các bảng danh mục bổ trợ
        conn.update(worksheet="ChuongTrinh", data=st.session_state.df_ct)
        conn.update(worksheet="DuTru", data=st.session_state.df_dt)
        conn.update(worksheet="NhomThuoc", data=st.session_state.df_nhom)
        
        st.toast("☁️ Đã đồng bộ dữ liệu lên Google Sheets thành công!", icon='✅')
    except Exception as e:
        st.error(f"Lỗi kết nối khi lưu dữ liệu: {e}")
        st.info("Mẹo: Đảm bảo bạn đã tắt các file Excel đang mở và kiểm tra kết nối mạng.")
def save_all():
    # 1. Sắp xếp lại kho thuốc theo tên Thành Phần trước khi lưu
    if not st.session_state.df_kho.empty:
        st.session_state.df_kho = st.session_state.df_kho.sort_values(by='Thành Phần', ascending=True)
        st.toast("💾 Đã đồng bộ dữ liệu!", icon='✅')

def generate_code(nhom, df):
    words = str(nhom).split()
    if len(words) >= 2:
        prefix = (words[0][0] + words[1][0]).upper()
    elif len(words) == 1:
        prefix = (words[0][:2]).upper() if len(words[0]) >= 2 else (words[0][0] + "X").upper()
    else:
        prefix = "TH"

    existing_codes = df['Barcode'].astype(str).tolist()
    relevant_nums = []
    for c in existing_codes:
        if c.startswith(prefix) and len(c) == (len(prefix) + 5):
            suffix = c[len(prefix):]
            if suffix.isdigit():
                relevant_nums.append(int(suffix))
    
    next_num = max(relevant_nums) + 1 if relevant_nums else 1
    return f"{prefix}{next_num:05d}"

if 'df_kho' not in st.session_state:
    st.session_state.df_kho, st.session_state.df_ls, st.session_state.df_ns, st.session_state.df_ct, st.session_state.df_dt, st.session_state.df_nhom = load_data()

st.session_state.df_kho['Tồn Kho'] = pd.to_numeric(st.session_state.df_kho['Nhập Mới'], errors='coerce').fillna(0) - pd.to_numeric(st.session_state.df_kho['Đã Xuất'], errors='coerce').fillna(0)

# --- LOGIN LOGIC ---
if "logged_in" not in st.session_state: st.session_state.logged_in = False
if not st.session_state.logged_in and os.path.exists(FILE_SESSION):
    try:
        with open(FILE_SESSION, "r") as f:
            u_save = f.read().strip()
            if u_save in st.session_state.df_ns['Username'].values:
                st.session_state.logged_in = True
                st.session_state.u_data = st.session_state.df_ns[st.session_state.df_ns['Username'] == u_save].iloc[0].to_dict()
    except: pass

apply_styles(main_color, FILE_BG)

if not st.session_state.logged_in:
    st.markdown(f"<h1 style='text-align:center; color:{main_color};'>🏥 DƯỢC KHOA Y NTTU</h1>", unsafe_allow_html=True)
    _, c2, _ = st.columns([1, 1.2, 1])
    with c2:
        with st.form("login"):
            u = st.text_input("Tên đăng nhập")
            p = st.text_input("Mật khẩu", type="password")
            if st.form_submit_button("ĐĂNG NHẬP"):
                ns = st.session_state.df_ns
                if u in ns['Username'].values and str(ns[ns['Username'] == u].iloc[0]['Password']) == p:
                    st.session_state.logged_in = True
                    st.session_state.u_data = ns[ns['Username'] == u].iloc[0].to_dict()
                    with open(FILE_SESSION, "w") as f: f.write(u)
                    st.rerun()
                else: st.error("Sai thông tin đăng nhập!")
else:
    with st.sidebar:
        st.write(f"Chào, **{st.session_state.u_data['Họ Tên']}**")
        if st.button("🚪 Đăng xuất"):
            if os.path.exists(FILE_SESSION): os.remove(FILE_SESSION)
            st.session_state.logged_in = False
            st.rerun()
        st.divider()
        nc = st.color_picker("Màu chủ đạo", main_color)
        if nc != main_color:
            with open(FILE_COLOR, "w") as f: f.write(nc)
            st.rerun()

    tabs = st.tabs(["📊 DASHBOARD", "📤 XUẤT THUỐC", "📥 NHẬP KHO", "🏷️ MÃ THUỐC VÀ NHÓM THUỐC", "📦 KHO TỔNG", "📝 DỰ TRÙ", "👤 CÁ NHÂN"])

    with tabs[0]:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Mặt hàng", len(st.session_state.df_kho))
        c2.metric("Tổng tồn", int(st.session_state.df_kho['Tồn Kho'].sum()))
        c3.metric("Chương trình", len(st.session_state.df_ct))
        c4.metric("Thành viên", len(st.session_state.df_ns))
        st.divider()
        st.bar_chart(st.session_state.df_kho.groupby('Nhóm Thuốc')['Tồn Kho'].sum())

    # --- TAB 2: XUẤT THUỐC ---
    with tabs[1]:
        st.subheader("📤 Cấp phát thuốc Chiến dịch")
        sel_ct = st.selectbox("Chọn chương trình:", st.session_state.df_ct[st.session_state.df_ct['Trạng Thái'] == 'Đang mở']['Tên Chương Trình'].tolist())
        col_x1, col_x2 = st.columns(2)
        m_find = col_x1.radio("Phương thức tìm:", ["Chọn Tên Thuốc", "Tìm theo Thành Phần", "Quét mã Barcode"], horizontal=True)
        
        res = pd.DataFrame()
        if m_find == "Quét mã Barcode":
            barcode = col_x2.text_input("📳 Đưa mã vào máy quét...")
            if barcode: res = st.session_state.df_kho[st.session_state.df_kho['Barcode'] == barcode]
        elif m_find == "Tìm theo Thành Phần":
            tp_list = sorted(st.session_state.df_kho['Thành Phần'].dropna().unique().tolist())
            sel_tp = col_x2.selectbox("Chọn Thành Phần (Hoạt chất):", ["---"] + tp_list)
            if sel_tp != "---":
                sub_df = st.session_state.df_kho[st.session_state.df_kho['Thành Phần'] == sel_tp]
                t_sel_tp = col_x2.selectbox("Khớp biệt dược phù hợp:", sub_df['Tên Biệt Dược'].tolist())
                res = sub_df[sub_df['Tên Biệt Dược'] == t_sel_tp]
        else:
            sorted_drugs = sorted(list(st.session_state.df_kho['Tên Biệt Dược'].unique()))
            t_sel = col_x2.selectbox("Chọn tên biệt dược:", ["---"] + sorted_drugs)
            if t_sel != "---": res = st.session_state.df_kho[st.session_state.df_kho['Tên Biệt Dược'] == t_sel]
            
        if not res.empty:
            st.success(f"Khớp: {res.iloc[0]['Tên Biệt Dược']} | Hoạt chất: {res.iloc[0]['Thành Phần']} (Tồn: {res.iloc[0]['Tồn Kho']})")
            with st.form("xuat_thuoc"):
                cx1, cx2 = st.columns(2)
                nguoi = cx1.text_input("Nơi nhận/Người nhận")
                sl_x = cx2.number_input("Số lượng xuất", min_value=1)
                if st.form_submit_button("🚀 XÁC NHẬN CẤP THUỐC"):
                    idx = res.index[0]
                    if st.session_state.df_kho.at[idx, 'Tồn Kho'] >= sl_x:
                        st.session_state.df_kho.at[idx, 'Đã Xuất'] += sl_x
                        row_ls = pd.DataFrame([{'Thời Gian': datetime.now().strftime("%d/%m/%Y %H:%M"), 'Chương Trình': sel_ct, 'Nơi Xuất': nguoi, 'Tên Thuốc': res.iloc[0]['Tên Biệt Dược'], 'Số Lượng': sl_x}])
                        st.session_state.df_ls = pd.concat([st.session_state.df_ls, row_ls], ignore_index=True)
                        save_all()
                        st.markdown(create_print_html(row_ls, "PHIẾU XUẤT THUỐC"), unsafe_allow_html=True)
                    else: st.error("Kho không đủ số lượng!")

    # --- TAB 3: NHẬP KHO (XỬ LÝ TỰ ĐỘNG EXCEL) ---
    with tabs[2]:
        st.subheader("📥 Nhập thuốc vào hệ thống")
        cn1, cn2 = st.columns(2)
        with cn1:
            with st.form("nhap_le"):
                st.write("#### 📝 Thêm thuốc lẻ (Mã tự động)")
                t_ten = st.text_input("Tên biệt dược")
                t_tp = st.text_input("Thành phần hoạt chất")
                t_nhom = st.selectbox("Nhóm thuốc", st.session_state.df_nhom['Tên Nhóm'].tolist())
                t_sl = st.number_input("Số lượng nhập", min_value=1)
                # Đảm bảo input từ lịch sang chuỗi dd/mm/yyyy
                t_hsd = st.date_input("Hạn sử dụng", format="DD/MM/YYYY")
                btn_nhap = st.form_submit_button("➕ THÊM VÀO KHO & TẠO MÃ")
                if btn_nhap:
                    m_code = generate_code(t_nhom, st.session_state.df_kho)
                    new_r = pd.DataFrame([{'Barcode': m_code, 'Tên Biệt Dược': t_ten, 'Thành Phần': t_tp, 'Nhóm Thuốc': t_nhom, 'Hạn Sử Dụng': t_hsd.strftime('%d/%m/%Y'), 'Nhập Mới': t_sl, 'Đã Xuất': 0, 'Chương Trình': 'Kho Tổng', 'Đơn Vị Tính': 'Viên'}])
                    st.session_state.df_kho = pd.concat([st.session_state.df_kho, new_r], ignore_index=True)
                    save_all()
                    st.session_state.last_added = {'code': m_code, 'name': t_ten}
                    st.rerun()

            if 'last_added' in st.session_state:
                st.success(f"Đã thêm: {st.session_state.last_added['name']} với mã: {st.session_state.last_added['code']}")
                val_last = st.session_state.last_added['code']
                qr_l = qrcode.make(val_last); buf_l = BytesIO(); qr_l.save(buf_l, format="PNG")
                st.image(buf_l, width=150, caption=f"Mã QR: {val_last}")
                st.download_button("🖨️ IN MÃ QR VỪA TẠO", buf_l.getvalue(), file_name=f"QR_{val_last}.png", mime="image/png")
                if st.button("Xóa thông báo"): del st.session_state.last_added; st.rerun()

        with cn2:
            st.write("#### 📥 Import Excel hàng loạt")
            st.info("Mã Barcode để trống hoặc 'AUTO' sẽ được tự động tạo theo nhóm.")
            st.download_button("📥 TẢI FILE MẪU EXCEL", data=get_excel_template(), file_name="mau_nhap_thuoc.xlsx")
            f_ex = st.file_uploader("Chọn file Excel:", type=['xlsx'])
            if f_ex and st.button("🚀 NẠP DỮ LIỆU EXCEL"):
                try:
                    df_up = pd.read_excel(f_ex)
                    # Chuyển định dạng ngày trong Excel nếu nó là object/datetime sang dd/mm/yyyy
                    if 'Hạn Sử Dụng' in df_up.columns:
                        df_up['Hạn Sử Dụng'] = pd.to_datetime(df_up['Hạn Sử Dụng'], errors='coerce').dt.strftime('%d/%m/%Y')
                    
                    for index, row in df_up.iterrows():
                        if pd.isna(row['Barcode']) or str(row['Barcode']).strip().upper() in ['', 'AUTO', 'NAN']:
                            nhom_hien_tai = row['Nhóm Thuốc'] if pd.notna(row['Nhóm Thuốc']) else "Khác"
                            new_code = generate_code(nhom_hien_tai, st.session_state.df_kho)
                            df_up.at[index, 'Barcode'] = new_code
                            st.session_state.df_kho = pd.concat([st.session_state.df_kho, df_up.iloc[[index]]], ignore_index=True)
                        else:
                            st.session_state.df_kho = pd.concat([st.session_state.df_kho, df_up.iloc[[index]]], ignore_index=True)
                    save_all()
                    st.success(f"Đã nạp thành công {len(df_up)} loại thuốc!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Lỗi định dạng: {e}")

    with tabs[3]:
        st.subheader("🏷️ Mã & Nhóm Thuốc")
        with st.expander("🛠️ Quản lý Nhóm Thuốc"):
            ed_nhom = st.data_editor(st.session_state.df_nhom, use_container_width=True, hide_index=True, num_rows="dynamic")
            if st.button("💾 LƯU NHÓM"):
                st.session_state.df_nhom = ed_nhom
                save_all(); st.rerun()
        st.divider()
        c_qr1, c_qr2 = st.columns([1, 1.2])
        with c_qr1:
            st.write("#### 🔍 QR đơn lẻ")
            s_thuoc = st.selectbox("Chọn thuốc:", ["---"] + sorted(st.session_state.df_kho['Tên Biệt Dược'].unique().tolist()), key="qr_tab")
            if s_thuoc != "---":
                row_d = st.session_state.df_kho[st.session_state.df_kho['Tên Biệt Dược'] == s_thuoc].iloc[0]
                val = row_d['Barcode']
                qr = qrcode.make(val); b = BytesIO(); qr.save(b, format="PNG")
                st.image(b, width=250); st.download_button("📥 Tải QR", b.getvalue(), f"QR_{val}.png")
        with c_qr2:
            st.write("#### 📂 In nhãn nhóm (A-Z Thành Phần)")
            sel_nh = st.selectbox("Chọn nhóm:", st.session_state.df_nhom['Tên Nhóm'].tolist())
            df_nh_res = st.session_state.df_kho[st.session_state.df_kho['Nhóm Thuốc'] == sel_nh].sort_values(by='Thành Phần')
            if not df_nh_res.empty:
                st.dataframe(df_nh_res[['Barcode', 'Thành Phần', 'Tên Biệt Dược', 'Tồn Kho']], use_container_width=True, hide_index=True)
                if st.button("🖨️ TẠO PDF IN NHÃN QR"):
                    st.markdown(create_qr_pdf_html(df_nh_res, sel_nh), unsafe_allow_html=True)

   # --- TAB 4: KHO TỔNG (PHIÊN BẢN SỬA LỖI KEYERROR) ---
    with tabs[4]:
        st.subheader("📦 Kho tổng (A-Z Thành Phần)")
        
        # 1. Kiểm tra nếu kho trống
        if st.session_state.df_kho.empty:
            st.info("Kho hiện đang trống. Hãy nhập thuốc ở tab 'Nhập Kho' hoặc kiểm tra file Google Sheets.")
        else:
            # Tự động tạo cột Trạng Thái HSD để đảm bảo nó luôn tồn tại trong DataFrame
            st.session_state.df_kho['Trạng Thái HSD'] = st.session_state.df_kho['Hạn Sử Dụng'].apply(check_hsd_status)
            st.session_state.df_kho = st.session_state.df_kho.sort_values(by='Thành Phần')
            
            # 2. Widgets Thống kê HSD
            c_hsd1, c_hsd2, c_hsd3 = st.columns(3)
            expired = len(st.session_state.df_kho[st.session_state.df_kho['Trạng Thái HSD'] == "❌ Hết hạn"])
            warning = len(st.session_state.df_kho[st.session_state.df_kho['Trạng Thái HSD'] == "⚠️ Sắp hết hạn"])
            c_hsd1.metric("Hết hạn", f"{expired} mục", delta=-expired if expired > 0 else 0, delta_color="inverse")
            c_hsd2.metric("Sắp hết hạn", f"{warning} mục")
            c_hsd3.info("⚠️ Cảnh báo sắp hết hạn: Còn dưới 180 ngày.")

            # 3. Xuất file Excel
            buf_ex = BytesIO()
            with pd.ExcelWriter(buf_ex, engine='xlsxwriter') as wr:
                st.session_state.df_kho.to_excel(wr, index=False)
            st.download_button("📥 XUẤT EXCEL KHO", buf_ex.getvalue(), f"Kho_{datetime.now().strftime('%d%m%Y')}.xlsx")
            
            # 4. CHỨC NĂNG QUẢN TRỊ (ADMIN)
            is_admin = st.session_state.u_data.get('Quyền') == 'admin'
            if is_admin:
                with st.expander("🗑️ CÔNG CỤ XÓA THUỐC"):
                    cx1, cx2 = st.columns(2)
                    with cx1:
                        st.write("#### 🔍 Xóa đơn lẻ")
                        thuoc_xoa = st.selectbox("Chọn thuốc cần xóa:", ["---"] + sorted(st.session_state.df_kho['Tên Biệt Dược'].tolist()))
                        if st.button("❌ XÁC NHẬN XÓA"):
                            if thuoc_xoa != "---":
                                st.session_state.df_kho = st.session_state.df_kho[st.session_state.df_kho['Tên Biệt Dược'] != thuoc_xoa]
                                save_all(); st.rerun()

                    with cx2:
                        st.write("#### ⚠️ Xóa toàn bộ kho")
                        confirm_all = st.checkbox("Xác nhận làm trống kho")
                        if st.button("🔥 XÓA TẤT CẢ"):
                            if confirm_all:
                                st.session_state.df_kho = pd.DataFrame(columns=BASE_COLS_KHO)
                                save_all(); st.rerun()

                st.divider()
                st.write("#### ✏️ Chỉnh sửa trực tiếp")
                
                # --- GIẢI PHÁP SỬA LỖI KEYERROR ---
                # Bước 1: Tạo bản sao để hiển thị
                df_to_show = st.session_state.df_kho.copy()
                
                # Bước 2: Kiểm tra cột tồn tại rồi mới áp dụng Style
                if 'Trạng Thái HSD' in df_to_show.columns:
                    try:
                        df_styled = df_to_show.style.applymap(color_hsd, subset=['Trạng Thái HSD'])
                    except:
                        df_styled = df_to_show # Nếu style lỗi thì hiện bảng thô
                else:
                    df_styled = df_to_show
                
                # Bước 3: Hiển thị trình chỉnh sửa dữ liệu
                ed_k = st.data_editor(df_styled, use_container_width=True, hide_index=True, num_rows="dynamic")
                
                if st.button("💾 LƯU THAY ĐỔI XUỐNG CLOUD"):
                    # Chỉ lưu các cột gốc, không lưu cột 'Trạng Thái HSD' tự tạo để tránh làm nặng Sheets
                    cols_to_save = [c for c in ed_k.columns if c != 'Trạng Thái HSD']
                    st.session_state.df_kho = ed_k[cols_to_save]
                    save_all(); st.rerun()
            else:
                st.dataframe(st.session_state.df_kho, use_container_width=True)

    with tabs[5]:
        st.subheader("📝 Dự trù Chiến dịch")
        with st.form("form_dt"):
            dt_ten = st.selectbox("Chọn thuốc:", sorted(st.session_state.df_kho['Tên Biệt Dược'].unique().tolist()))
            dt_sl = st.number_input("Số lượng dự kiến:", min_value=1)
            if st.form_submit_button("Thêm dự trù"):
                st.session_state.df_dt = pd.concat([st.session_state.df_dt, pd.DataFrame([{'Tên Thuốc': dt_ten, 'Số Lượng Dự Trù': dt_sl}])], ignore_index=True)
                save_all(); st.rerun()
        st.table(st.session_state.df_dt)

    with tabs[6]:
        st.subheader("👤 Hồ sơ")
        u_idx = st.session_state.df_ns[st.session_state.df_ns['Username'] == st.session_state.u_data['Username']].index[0]
        c_h1, c_h2 = st.columns(2)
        with c_h1:
            st.write("#### ✏️ Thông tin")
            new_ten = st.text_input("Họ Tên:", value=st.session_state.u_data['Họ Tên'])
            if st.button("Lưu tên"):
                st.session_state.df_ns.at[u_idx, 'Họ Tên'] = new_ten
                st.session_state.u_data['Họ Tên'] = new_ten
                save_all(); st.rerun()
        with c_h2:
            st.write("#### 🔐 Mật khẩu")
            old_p = st.text_input("Cũ:", type="password")
            new_p = st.text_input("Mới:", type="password")
            if st.button("Đổi mật khẩu"):
                if str(st.session_state.df_ns.at[u_idx, 'Password']) == old_p:
                    st.session_state.df_ns.at[u_idx, 'Password'] = new_p
                    save_all(); st.success("Thành công!"); st.rerun()
                else: st.error("Sai mật khẩu cũ!")
        if st.session_state.u_data.get('Quyền') == 'admin':
            st.divider()
            st.subheader("👥 Nhân sự")
            ed_ns = st.data_editor(st.session_state.df_ns, use_container_width=True, hide_index=True, num_rows="dynamic")
            if st.button("💾 LƯU NHÂN SỰ"):
                st.session_state.df_ns = ed_ns
                save_all(); st.rerun()


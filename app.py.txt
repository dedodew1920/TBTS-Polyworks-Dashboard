import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
from folium.plugins import Fullscreen
import requests
import re
import sqlite3
import os

# --- 🎨 1. ตั้งค่าหน้าเพจเริ่มต้น ---
st.set_page_config(layout="wide", page_title="Polyworks Dashboard")

# --- 🔑 2. ระบบหน้าจอ Login (ดักตั้งแต่หน้าแรกสุด) ---
def check_password():
    """ฟังก์ชันล็อกหน้าเว็บ คืนค่า True ก็ต่อเมื่อใส่รหัสผ่านถูกเท่านั้น"""
    def password_entered():
        # ค้นหารหัสผ่านจากระบบหลังบ้าน (ถ้าทดสอบเครื่องตัวเองแล้วไม่ได้ตั้ง ค่าเริ่มต้นจะเป็น admin1234)
        correct_password = st.secrets.get("PASSWORD", "admin1234")
        if st.session_state["password_input"] == correct_password:
            st.session_state["password_correct"] = True
            del st.session_state["password_input"]  # ลบรหัสออกเพื่อความปลอดภัย
        else:
            st.session_state["password_correct"] = False

    # กรณีเข้ามาครั้งแรกสุด (ยังไม่ได้ล็อกอิน) ให้ล็อกหน้าจอและหยุดการทำงานไว้แค่นี้
    if "password_correct" not in st.session_state:
        st.markdown("<h2 style='text-align: center;'>🔒 Polyworks Dashboard Login</h2>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.text_input("กรุณากรอกรหัสผ่านเพื่อเข้าสู่ระบบ", type="password", on_change=password_entered, key="password_input")
        return False
        
    # กรณีกรอกรหัสผ่านผิด ให้ขึ้นข้อความเตือนแดงๆ และไม่ให้ผ่านไปไหน
    elif not st.session_state["password_correct"]:
        st.markdown("<h2 style='text-align: center;'>🔒 Polyworks Dashboard Login</h2>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.text_input("กรุณากรอกรหัสผ่านเพื่อเข้าสู่ระบบ", type="password", on_change=password_entered, key="password_input")
            st.error("❌ รหัสผ่านไม่ถูกต้อง! กรุณาตรวจสอบใหม่อีกครั้ง")
        return False
    return True


# --- 3. ตรวจสอบเงื่อนไขการเข้าถึงหน้าใช้งานหลัก ---
if check_password():

    # ถ้ารหัสผ่านถูกต้อง ให้ประยุกต์ใช้สไตล์ความสวยงามของเว็บคุณทันที
    st.markdown("""
        <style>
        .stExpander { 
            border: 1px solid var(--secondary-background-color) !important; 
            border-radius: 8px !important; 
            margin-bottom: 10px !important;
            background-color: var(--background-color) !important; 
        }
        .info-card {
            padding: 10px;
            border-left: 5px solid;
            color: var(--text-color) !important;
        }
        .info-card b, .info-card p, .info-card span {
            color: var(--text-color) !important;
        }
        hr { opacity: 0.2; }
        </style>
    """, unsafe_allow_html=True)

    # --- ปุ่มออกจากระบบ (Logout) บนแถบเมนูด้านซ้าย ---
    if st.sidebar.button("🚪 ออกจากระบบ (Log Out)"):
        del st.session_state["password_correct"]
        st.rerun()

    # ---------------------------------------------------------------
    # --- ⚙️ 4. หน้าจอหลัก (เริ่มทำงานเมื่อปลดล็อกรหัสผ่านสำเร็จ) ---
    # ---------------------------------------------------------------
    DB_NAME = "polyworks_data.db"
    FILE_PATH = "Polyworks Contract.xlsx" 
    SHEET_NAME = "PolyWorks MA Contract"

    def sync_db_from_excel():
        try:
            if os.path.exists(DB_NAME): os.remove(DB_NAME)
            df = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME, header=2)
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            if 'DongleNo.' in df.columns:
                df['DongleNo.'] = df['DongleNo.'].astype(str).str.strip()
            with sqlite3.connect(DB_NAME) as conn:
                df.to_sql("data", conn, if_exists="replace", index=False)
            return True
        except Exception as e:
            st.error(f"Error syncing: {e}")
            return False

    def load_data():
        if not os.path.exists(DB_NAME): sync_db_from_excel()
        if not os.path.exists(DB_NAME): return pd.DataFrame()
        with sqlite3.connect(DB_NAME) as conn:
            df = pd.read_sql("SELECT * FROM data", conn) 
        df = df.dropna(subset=['Company', 'DongleNo.'], how='all')
        return df

    @st.cache_data
    def get_coords_optimized(lat_val, lon_val, url_map):
        if pd.notna(lat_val):
            val_str = str(lat_val).strip()
            if "," in val_str:
                try:
                    p = val_str.split(",")
                    return float(p[0].strip()), float(p[1].strip())
                except: pass
        if pd.notna(url_map) and isinstance(url_map, str) and url_map.strip().startswith("http"):
            try:
                with requests.Session() as s:
                    resp = s.head(url_map, allow_redirects=True, timeout=1)
                    match = re.search(r'@([-+]?\d+\.\d+),([-+]?\d+\.\d+)', resp.url)
                    if match: return float(match.group(1)), float(match.group(2))
            except: pass
        return None, None

    # --- Initialization ---
    if 'main_df' not in st.session_state:
        st.session_state.main_df = load_data()

    # --- Sidebar ---
    st.sidebar.header("🔍 ตัวกรองข้อมูล")
    display_mode = st.sidebar.radio("โหมดการทำงาน:", ["📦 View Mode", "📝 Edit Mode"])

    df = st.session_state.main_df

    if not df.empty:
        suggestions = sorted(list(set(df['Company'].dropna().astype(str).unique().tolist() + df['Division'].dropna().astype(str).unique().tolist())))
        search_query = st.sidebar.selectbox("🎯 ค้นหา บริษัท/แผนก", options=[""] + suggestions)
        
        all_statuses = sorted(df['Status'].dropna().unique().tolist())
        selected_statuses = st.sidebar.multiselect("🚦 สถานะ Filter", options=all_statuses)

        st.sidebar.markdown("---")
        st.sidebar.subheader("📅 สิทธิ์การอัปเดตข้อมูล")
        f24, f25, f26 = st.sidebar.checkbox("2024"), st.sidebar.checkbox("2025"), st.sidebar.checkbox("2026")

        filtered_df = df.copy()
        if f24: filtered_df = filtered_df[filtered_df['2024 Can Update'] == 'OK']
        if f25: filtered_df = filtered_df[filtered_df['2025 Can Update'] == 'OK']
        if f26: filtered_df = filtered_df[filtered_df['2026 Can Update'] == 'OK']
        if selected_statuses: filtered_df = filtered_df[filtered_df['Status'].isin(selected_statuses)]
        if search_query:
            filtered_df = filtered_df[filtered_df['Company'].astype(str).str.contains(search_query, case=False) | 
                                     filtered_df['Division'].astype(str).str.contains(search_query, case=False)]
    else:
        filtered_df = pd.DataFrame()

    # --- หน้าหลักแสดงผล ---
    st.title("📍 Polyworks Maintenance Dashboard")
    if not filtered_df.empty:
        total = len(filtered_df)
        expired = len(filtered_df[filtered_df['Status'].str.lower().str.contains('expired', na=False)])
        ok = total - expired
        c1, c2, c3 = st.columns(3)
        c1.metric("📋 ข้อมูลทั้งหมด", total)
        c2.metric("✅ ใช้งานได้ปกติ", ok)
        c3.metric("❌ หมดอายุ (Expired)", expired, delta=f"-{expired}" if expired > 0 else 0, delta_color="inverse")

    st.divider()

    # --- 🗺️ แสดงแผนที่ ---
    with st.expander("🗺️ แผนที่ระบุพิกัดที่ตั้งลูกค้า", expanded=True):
        if not filtered_df.empty:
            coords = filtered_df.apply(lambda r: get_coords_optimized(r.get('Lat'), r.get('Lon'), r.get('Map')), axis=1)
            temp_df = filtered_df.copy(); temp_df['Lat_f'], temp_df['Lon_f'] = zip(*coords)
            map_data = temp_df.dropna(subset=['Lat_f', 'Lon_f'])
            if not map_data.empty:
                m = folium.Map(location=[map_data['Lat_f'].mean(), map_data['Lon_f'].mean()], zoom_start=7, tiles="CartoDB positron")
                Fullscreen().add_to(m)
                for _, r in map_data.iterrows():
                    is_exp = 'expired' in str(r['Status']).lower()
                    marker_color = 'red' if is_exp else 'green'
                    popup_html = f"<div style='font-family:sans-serif;'><b>{r['Company']}</b><br>S/N: {r['DongleNo.']}</div>"
                    folium.Marker([r['Lat_f'], r['Lon_f']], popup=folium.Popup(popup_html, max_width=200), icon=folium.Icon(color=marker_color)).add_to(m)
                st_folium(m, width="100%", height=400, key="main_map")

    st.divider()

    # --- แสดงผลโหมดพนักงานอ่าน หรือ โหมดแก้ไขแอดมิน ---
    if display_mode == "📦 View Mode":
        if filtered_df.empty:
            st.info("ไม่พบข้อมูลพิกัดหรือรายชื่อ")
        else:
            cols = st.columns(4) 
            for idx, row in filtered_df.reset_index().iterrows():
                if pd.isna(row.get('Company')) or str(row.get('Company')).strip() == "":
                    continue
                    
                with cols[idx % 4]:
                    status_val = str(row.get('Status', '-'))
                    is_exp = 'expired' in status_val.lower()
                    status_color = "#FF4B4B" if is_exp else "#00C49A"
                    header = f"{'🔴' if is_exp else '🟢'} {str(row.get('Company'))[:12]}.."
                    with st.expander(header):
                        st.markdown(f"""<div class="info-card" style="border-color: {status_color};"><p><b>Company:</b> {row.get('Company')}</p><p><b>Division:</b> {row.get('Division')}</p><hr><p><b>S/N:</b> {row.get('DongleNo.')}</p><p><b>Expire:</b> {row.get('Expire')}</p><p><b>Status:</b> <span style="color:{status_color};font-weight:bold;">{status_val}</span></p></div>""", unsafe_allow_html=True)
                        if pd.notna(row.get('Map')) and str(row.get('Map')).startswith("http"):
                            st.link_button("📍 เปิดดูในแผนที่จริง", str(row.get('Map')).strip(), use_container_width=True)

    else: # 📝 Edit Mode
        st.subheader("📝 ระบบแก้ไขตารางข้อมูลลูกค้า")
        edited_df = st.data_editor(
            filtered_df, 
            use_container_width=True, 
            num_rows="dynamic",
            column_config={"Map": st.column_config.LinkColumn("Map Link", display_text="📍 View Map")},
            key="editor_key"
        )
        
        if st.button("💾 บันทึกการเปลี่ยนแปลงข้อมูล"):
            try:
                with sqlite3.connect(DB_NAME) as conn:
                    old_ids = filtered_df['DongleNo.'].astype(str).str.strip().tolist()
                    if old_ids:
                        # ใช้ SQL Parameter Binding เพื่อป้องกันการเจาะระบบ (SQL Injection)
                        placeholders = ','.join(['?'] * len(old_ids))
                        conn.execute(f"DELETE FROM data WHERE CAST(`DongleNo.` AS TEXT) IN ({placeholders})", old_ids)
                    
                    final_save_df = edited_df.dropna(subset=['Company', 'DongleNo.'], how='all')
                    final_save_df.to_sql("data", conn, if_exists="append", index=False)
                    conn.commit()
                
                st.session_state.main_df = load_data()
                st.success("✅ บันทึกข้อมูลและอัปเดตลงฐานข้อมูลเรียบร้อย!")
                st.rerun()
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดขณะบันทึก: {e}")

        if st.sidebar.button("🔄 อัปเดตตารางใหม่จากไฟล์ Excel"):
            if sync_db_from_excel():
                st.session_state.main_df = load_data()
                st.rerun()
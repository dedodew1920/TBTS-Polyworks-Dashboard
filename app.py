import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
from folium.plugins import Fullscreen
import sqlite3
import os
import re

# --- 🎨 Page Config ---
st.set_page_config(layout="wide", page_title="Polyworks Dashboard")

# --- Configuration ---
FILE_PATH = "Polyworks Contract.xlsx"
DB_NAME = "polyworks_data.db"
SHEET_NAME = "PolyWorks MA Contract"

# --- Functions ---

def mask_phone_number(phone):
    """ฟังก์ชันสำหรับซ่อนเบอร์โทรศัพท์ (เช่น 081-xxx-xx99)"""
    phone_str = str(phone).strip()
    if len(phone_str) > 5:
        return f"{phone_str[:3]}-xxx-xx{phone_str[-2:]}"
    return phone_str

def sync_db_from_excel():
    try:
        if not os.path.exists(FILE_PATH):
            st.error(f"❌ ไม่พบไฟล์ {FILE_PATH} ใน GitHub")
            return False
        if os.path.exists(DB_NAME): 
            os.remove(DB_NAME)
        df = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME, header=2)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        conn = sqlite3.connect(DB_NAME)
        df.to_sql("data", conn, if_exists="replace", index=False)
        conn.close()
        return True
    except Exception as e:
        st.error(f"⚠️ Error: {e}")
        return False

def load_data():
    if not os.path.exists(DB_NAME): 
        sync_db_from_excel()
    if os.path.exists(DB_NAME):
        conn = sqlite3.connect(DB_NAME)
        df = pd.read_sql("SELECT * FROM data", conn)
        conn.close()
        return df
    return pd.DataFrame()

def save_data(df_to_save):
    conn = sqlite3.connect(DB_NAME)
    full_df = pd.read_sql("SELECT * FROM data", conn)
    full_df.update(df_to_save)
    new_rows = df_to_save[~df_to_save.index.isin(full_df.index)]
    if not new_rows.empty:
        full_df = pd.concat([full_df, new_rows], ignore_index=True)
    full_df.to_sql("data", conn, if_exists="replace", index=False)
    conn.close()

@st.cache_data
def get_coords_optimized(lat_val, lon_val, url_map):
    if pd.notna(lat_val):
        val_str = str(lat_val).strip()
        if "," in val_str:
            try:
                parts = val_str.split(",")
                return float(parts[0].strip()), float(parts[1].strip())
            except: pass
    if pd.notna(url_map) and isinstance(url_map, str) and url_map.strip().startswith("http"):
        import requests
        try:
            with requests.Session() as s:
                response = s.head(url_map, allow_redirects=True, timeout=1)
                match = re.search(r'@([-+]?\d+\.\d+),([-+]?\d+\.\d+)', response.url)
                if match: return float(match.group(1)), float(match.group(2))
        except: pass
    return None, None

# --- 🚀 Main Logic ---
df = load_data()

st.sidebar.header("🔍 ค้นหาและตัวกรอง")
display_mode = st.sidebar.radio("โหมดการใช้งาน:", ["📦 View Mode", "📝 Edit Mode"])

if st.sidebar.button("🔄 Sync ข้อมูลจาก GitHub", use_container_width=True):
    if sync_db_from_excel():
        st.cache_data.clear()
        st.rerun()

# --- Dashboard UI ---
st.title("📍 Polyworks Maintenance Dashboard")

if not df.empty:
    filtered_df = df.copy()
    
    # Summary Statistics
    total_count = len(filtered_df)
    expired_count = len(filtered_df[filtered_df['Status'].str.lower().str.contains('expired', na=False)])
    not_expired_count = total_count - expired_count
    
    col_s1, col_s2, col_s3 = st.columns(3)
    with col_s1: st.metric("📋 ทั้งหมด", f"{total_count} รายการ")
    with col_s2: st.metric("✅ ปกติ", f"{not_expired_count} รายการ")
    with col_s3: st.metric("❌ หมดอายุ", f"{expired_count} รายการ", delta=f"-{expired_count}" if expired_count > 0 else 0, delta_color="inverse")
    
    st.divider()

    if display_mode == "📦 View Mode":
        for idx, row in filtered_df.iterrows():
            is_expired = 'expired' in str(row['Status']).lower()
            st.markdown(f"""
                <div style="background-color:{'#FF4B4B' if is_expired else '#00C49A'}; padding:10px; border-radius:5px 5px 0 0; color:white; font-weight:bold;">
                    🏢 {row['Company']} | 📂 {row.get('Division','-')}
                </div>
            """, unsafe_allow_html=True)
            with st.expander("รายละเอียด"):
                c1, c2 = st.columns(2)
                with c1:
                    st.write(f"**S/N:** {row.get('DongleNo.','-')}")
                    st.write(f"**ผู้ติดต่อ:** {row.get('Contact','-')}")
                    st.write(f"**เบอร์โทร:** {mask_phone_number(row.get('TEL','-'))}")
                with c2:
                    st.write(f"**สิ้นสุดสัญญา:** {row.get('Expire','-')}")
                    if pd.notna(row.get('Map')):
                        st.link_button("🚀 เปิด Google Maps", str(row.get('Map')))
    else:
        edited_df = st.data_editor(filtered_df, use_container_width=True)
        if st.button("💾 บันทึกการแก้ไข"):
            save_data(edited_df)
            st.success("บันทึกเรียบร้อย!")
            st.rerun()
else:
    st.info("โปรดอัปโหลดไฟล์ Polyworks Contract.xlsx ขึ้น GitHub และกดปุ่ม Sync")
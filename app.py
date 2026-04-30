import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
from folium.plugins import Fullscreen
import requests
import re
import sqlite3
import os
import io

# --- 🎨 Custom CSS & Page Config ---
st.set_page_config(layout="wide", page_title="Polyworks Dashboard")

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .detail-label { font-weight: 600; color: #666; font-size: 0.85rem; text-transform: uppercase; }
    .detail-value { font-size: 0.95rem; color: #333; margin-bottom: 8px; }
    .stExpander { border: none !important; margin-bottom: 5px !important; }
    </style>
""", unsafe_allow_html=True)

# --- Configuration ---
# ✅ แก้ไข URL ให้พร้อมสำหรับการ Download โดยตรง
SHAREPOINT_URL = "https://tbtsonline.sharepoint.com/sites/TBTS-Engineering/Shared%20Documents/PRODUCT%20LIST/004)%20Software/PolyWorks/my%20dashboard/Polyworks%20Contract.xlsx?download=1"
DB_NAME = "polyworks_data.db"
SHEET_NAME = "PolyWorks MA Contract"

# --- Functions ---

def sync_db_from_excel():
    """ดึงข้อมูลจาก SharePoint และบันทึกลง SQLite DB"""[cite: 1]
    try:
        response = requests.get(SHAREPOINT_URL)
        if response.status_code == 200:
            if os.path.exists(DB_NAME): os.remove(DB_NAME)
            
            # อ่านจาก Excel (เริ่มที่แถว 3 ตาม Header=2)
            df = pd.read_excel(io.BytesIO(response.content), sheet_name=SHEET_NAME, header=2)[cite: 1]
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            
            conn = sqlite3.connect(DB_NAME)
            df.to_sql("data", conn, if_exists="replace", index=False)[cite: 1]
            conn.close()
            return True
        else:
            st.error(f"❌ ไม่สามารถโหลดไฟล์จาก SharePoint ได้ (Error {response.status_code})")
            return False
    except Exception as e:
        st.error(f"⚠️ เกิดข้อผิดพลาดในการ Sync: {e}")
        return False

def load_data():
    """โหลดข้อมูลจาก SQLite DB"""[cite: 1]
    if not os.path.exists(DB_NAME): 
        sync_db_from_excel()
    
    if os.path.exists(DB_NAME):
        conn = sqlite3.connect(DB_NAME)
        df = pd.read_sql("SELECT * FROM data", conn)[cite: 1]
        conn.close()
        return df
    return pd.DataFrame()

def save_data(df_to_save):
    """บันทึกข้อมูลที่แก้ไขกลับลง DB"""[cite: 1]
    conn = sqlite3.connect(DB_NAME)
    full_df = pd.read_sql("SELECT * FROM data", conn)
    full_df.update(df_to_save)
    new_rows = df_to_save[~df_to_save.index.isin(full_df.index)]
    if not new_rows.empty:
        full_df = pd.concat([full_df, new_rows], ignore_index=True)
    full_df.to_sql("data", conn, if_exists="replace", index=False)[cite: 1]
    conn.close()

@st.cache_data
def get_coords_optimized(lat_val, lon_val, url_map):
    """แปลงพิกัดจากค่าในตารางหรือ Google Maps URL"""[cite: 1]
    if pd.notna(lat_val):
        val_str = str(lat_val).strip()
        if "," in val_str:
            try:
                parts = val_str.split(",")
                return float(parts[0].strip()), float(parts[1].strip())
            except: pass
    if pd.notna(url_map) and isinstance(url_map, str) and url_map.strip().startswith("http"):
        try:
            with requests.Session() as s:
                response = s.head(url_map, allow_redirects=True, timeout=1)
                match = re.search(r'@([-+]?\d+\.\d+),([-+]?\d+\.\d+)', response.url)
                if match: return float(match.group(1)), float(match.group(2))
        except: pass
    return None, None

def get_status_color(status_val):
    """กำหนดสีตามสถานะสัญญา"""[cite: 1]
    return '#FF4B4B' if 'expired' in str(status_val).lower() else '#00C49A'

# --- 🚀 Main Logic ---

df = load_data()

# --- Sidebar ---
st.sidebar.header("🔍 ค้นหาและตัวกรอง")
display_mode = st.sidebar.radio("โหมดการใช้งาน:", ["📦 View Mode", "📝 Edit Mode"])[cite: 1]

if not df.empty:
    # สร้างรายการค้นหาจาก Company, Division, DongleNo.
    suggestions = sorted(list(set(
        df['Company'].dropna().astype(str).unique().tolist() + 
        df['Division'].dropna().astype(str).unique().tolist() + 
        df['DongleNo.'].dropna().astype(str).unique().tolist()
    )))
    search_query = st.sidebar.selectbox("🎯 ค้นหา", options=[""] + suggestions, index=0)
    
    all_statuses = sorted(df['Status'].dropna().unique().tolist())
    selected_statuses = st.sidebar.multiselect("🚦 สถานะสัญญา", options=all_statuses)
else:
    search_query = ""
    selected_statuses = []

# Filter
filtered_df = df.copy()
if search_query:
    filtered_df = filtered_df[
        filtered_df['Company'].astype(str).str.contains(search_query, case=False, na=False) |
        filtered_df['Division'].astype(str).str.contains(search_query, case=False, na=False) |
        filtered_df['DongleNo.'].astype(str).str.contains(search_query, case=False, na=False)
    ]
if selected_statuses:
    filtered_df = filtered_df[filtered_df['Status'].isin(selected_statuses)]

if st.sidebar.button("🔄 Sync ข้อมูลใหม่จาก SharePoint", use_container_width=True):
    if sync_db_from_excel():
        st.cache_data.clear() 
        st.success("อัปเดตข้อมูลสำเร็จ!")
        st.rerun()

# --- Dashboard UI ---
st.title("📍 Polyworks Maintenance Dashboard")

if not filtered_df.empty:
    total_count = len(filtered_df)
    expired_count = len(filtered_df[filtered_df['Status'].str.lower().str.contains('expired', na=False)])
    not_expired_count = total_count - expired_count
    
    col_s1, col_s2, col_s3 = st.columns(3)
    with col_s1: st.metric("📋 ทั้งหมด", f"{total_count} รายการ")
    with col_s2: st.metric("✅ ปกติ", f"{not_expired_count} รายการ")
    with col_s3: st.metric("❌ หมดอายุ", f"{expired_count} รายการ", delta=f"-{expired_count}" if expired_count > 0 else 0, delta_color="inverse")

st.divider()

# --- Map ---
with st.expander("🗺️ แผนที่พิกัดลูกค้า", expanded=True):
    if not filtered_df.empty:
        coords = filtered_df.apply(lambda r: get_coords_optimized(r.get('Lat'), r.get('Lon'), r.get('Map')), axis=1)
        temp_df = filtered_df.copy()
        temp_df['Lat_f'], temp_df['Lon_f'] = zip(*coords)
        map_data = temp_df.dropna(subset=['Lat_f', 'Lon_f'])
        
        if not map_data.empty:
            m = folium.Map(location=[map_data['Lat_f'].mean(), map_data['Lon_f'].mean()], zoom_start=8, tiles="CartoDB positron")
            Fullscreen(position="topright").add_to(m)

            for (company, lat, lon), group in map_data.groupby(['Company', 'Lat_f', 'Lon_f']):
                dept_list = "".join([f"<div style='border-bottom:1px solid #eee; padding:3px;'><b>{r['Division']}</b>: <span style='color:{get_status_color(r['Status'])};'>{r['Status']}</span></div>" for _, r in group.iterrows()])
                marker_color = 'red' if 'expired' in group['Status'].str.lower().values else 'green'
                folium.Marker([lat, lon], popup=folium.Popup(f"🏢 <b>{company}</b><br>{dept_list}", max_width=300), icon=folium.Icon(color=marker_color)).add_to(m)
            
            st_folium(m, width="100%", height=450, key="main_map")
        else:
            st.info("ไม่มีข้อมูลพิกัดแสดงบนแผนที่")

# --- Content Display ---
if display_mode == "📦 View Mode":
    for idx, row in filtered_df.iterrows():
        is_expired = 'expired' in str(row['Status']).lower()
        st.markdown(f"""
            <div style="background-color:{'#FF4B4B' if is_expired else '#00C49A'}; padding:10px; border-radius:5px 5px 0 0; color:white; font-weight:bold;">
                🏢 {row['Company']} | 📂 {row.get('Division','-')} | 🔑 SN: {row.get('DongleNo.','-')}
            </div>
        """, unsafe_allow_html=True)
        with st.expander("ดูรายละเอียดเพิ่มเติม"):
            c1, c2 = st.columns(2)
            with c1:
                st.write(f"**ผู้ติดต่อ:** {row.get('Contact','-')}")
                st.write(f"**เบอร์โทร:** {row.get('TEL','-')}")
            with c2:
                st.write(f"**สิ้นสุดสัญญา:** {row.get('Expire','-')}")
                st.link_button("🚀 เปิด Google Maps", str(row.get('Map','')))
else:
    edited_df = st.data_editor(filtered_df, use_container_width=True)
    if st.button("💾 บันทึกการแก้ไข"):
        save_data(edited_df)
        st.success("บันทึกข้อมูลเรียบร้อย!")
        st.rerun()
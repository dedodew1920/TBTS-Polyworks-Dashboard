import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
from folium.plugins import Fullscreen
import requests
import re
import sqlite3
import os

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
DB_NAME = "polyworks_data.db"
FILE_PATH = "Polyworks Contract.xlsx"
SHEET_NAME = "PolyWorks MA Contract"

# --- Database Functions ---
def sync_db_from_excel():
    try:
        if os.path.exists(DB_NAME): os.remove(DB_NAME)
        df = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME, header=2)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        conn = sqlite3.connect(DB_NAME)
        df.to_sql("data", conn, if_exists="replace", index=False)
        conn.close()
        return True
    except:
        return False

def load_data():
    if not os.path.exists(DB_NAME): 
        sync_db_from_excel()
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql("SELECT * FROM data", conn) 
    conn.close()
    return df

def save_data(df_to_save):
    """บันทึกข้อมูลทั้งหมดลงใน Database โดยตรง"""[cite: 1]
    conn = sqlite3.connect(DB_NAME)
    df_to_save.to_sql("data", conn, if_exists="replace", index=False)
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
        try:
            with requests.Session() as s:
                response = s.head(url_map, allow_redirects=True, timeout=1)
                match = re.search(r'@([-+]?\d+\.\d+),([-+]?\d+\.\d+)', response.url)
                if match: return float(match.group(1)), float(match.group(2))
        except: pass
    return None, None

def get_status_color(status_val):
    return '#FF4B4B' if 'expired' in str(status_val).lower() else '#00C49A'

# --- Initialization ---
df_all = load_data() # โหลดข้อมูลทั้งหมดเก็บไว้ก่อน

# --- Sidebar ---
st.sidebar.header("🔍 ค้นหาและตัวกรอง")
display_mode = st.sidebar.radio("โหมดการใช้งาน:", ["📦 View Mode", "📝 Edit Mode"])[cite: 1]

# เตรียมข้อมูลตัวเลือกสำหรับตัวกรอง (อ้างอิงจากข้อมูลทั้งหมด)
if not df_all.empty:
    suggestions = sorted(list(set(
        df_all['Company'].dropna().astype(str).unique().tolist() + 
        df_all['Division'].dropna().astype(str).unique().tolist() + 
        df_all['DongleNo.'].dropna().astype(str).unique().tolist()
    )))
    search_query = st.sidebar.selectbox("🎯 ค้นหา (ในโหมด View)", options=[""] + suggestions, index=0)
    all_statuses = sorted(df_all['Status'].dropna().unique().tolist())
else:
    search_query = ""
    all_statuses = []

selected_statuses = st.sidebar.multiselect("🚦 สถานะสัญญา", options=all_statuses)

# --- 🔄 Sync Button ---
if st.sidebar.button("🔄 Sync ข้อมูลใหม่จาก Excel", use_container_width=True):[cite: 1]
    sync_db_from_excel()
    st.cache_data.clear() 
    st.rerun()

# --- Filter Logic (เฉพาะ View Mode เท่านั้น) ---
filtered_df = df_all.copy()
if search_query:
    filtered_df = filtered_df[
        filtered_df['Company'].astype(str).str.contains(search_query, case=False, na=False) |
        filtered_df['Division'].astype(str).str.contains(search_query, case=False, na=False) |
        filtered_df['DongleNo.'].astype(str).str.contains(search_query, case=False, na=False)
    ]
if selected_statuses:
    filtered_df = filtered_df[filtered_df['Status'].isin(selected_statuses)]

# --- Main Dashboard ---
st.title("📍 Polyworks Maintenance Dashboard")

# --- 📦 DISPLAY MODES ---
if display_mode == "📦 View Mode":[cite: 1]
    # Summary Metrics
    if not filtered_df.empty:
        total_count = len(filtered_df)
        expired_count = len(filtered_df[filtered_df['Status'].str.lower().str.contains('expired', na=False)])
        not_expired_count = total_count - expired_count
        
        col_s1, col_s2, col_s3 = st.columns(3)
        with col_s1: st.metric("📋 ทั้งหมด", f"{total_count} รายการ")
        with col_s2: st.metric("✅ ปกติ", f"{not_expired_count} รายการ")
        with col_s3: st.metric("❌ หมดอายุ", f"{expired_count} รายการ", delta=f"-{expired_count}" if expired_count > 0 else 0, delta_color="inverse")

    st.divider()

    # Map Section
    with st.expander("🗺️ แผนที่พิกัดลูกค้า", expanded=True):[cite: 1]
        if not filtered_df.empty:
            coords = filtered_df.apply(lambda r: get_coords_optimized(r.get('Lat'), r.get('Lon'), r.get('Map')), axis=1)
            temp_df = filtered_df.copy()
            temp_df['Lat_f'], temp_df['Lon_f'] = zip(*coords)
            map_data = temp_df.dropna(subset=['Lat_f', 'Lon_f'])
            
            if not map_data.empty:
                m = folium.Map(location=[map_data['Lat_f'].mean(), map_data['Lon_f'].mean()], zoom_start=8, tiles="CartoDB positron")
                Fullscreen(position="topright").add_to(m)[cite: 1]

                for (company, lat, lon), group in map_data.groupby(['Company', 'Lat_f', 'Lon_f']):
                    dept_html = "".join([f"<div style='border-bottom:1px solid #eee; padding:3px;'><b>{r['Division']}</b>: <span style='color:{get_status_color(r['Status'])};'>{r['Status']}</span></div>" for _, r in group.iterrows()])
                    marker_color = 'red' if 'expired' in group['Status'].str.lower().values else 'green'
                    folium.Marker([lat, lon], popup=folium.Popup(f"🏢 <b>{company}</b><br>{dept_html}", max_width=300), icon=folium.Icon(color=marker_color)).add_to(m)
                
                st_folium(m, width="100%", height=450, key="view_map")
            else:
                st.info("ไม่พบข้อมูลพิกัดในรายการที่เลือก")

    # Details Cards
    st.subheader("📋 รายละเอียดรายแผนก")
    for idx, row in filtered_df.iterrows():
        is_expired = 'expired' in str(row['Status']).lower()
        bg_color = "#FF4B4B" if is_expired else "#00C49A"
        st.markdown(f"""
            <div style="background-color:{bg_color}; padding:10px 15px; border-radius:5px 5px 0 0; color:white; font-weight:bold;">
                🏢 {row['Company']} | 📂 {row.get('Division','-')}
            </div>
        """, unsafe_allow_html=True)
        with st.expander("คลิกดูรายละเอียด"):
            c1, c2, c3 = st.columns(3)
            with c1:
                st.write(f"**S/N:** {row.get('DongleNo.','-')}")
                st.write(f"**ผู้ติดต่อ:** {row.get('Contact','-')}")
            with c2:
                st.write(f"**โทร:** {row.get('TEL','-')}")
                st.write(f"**สิ้นสุดสัญญา:** {row.get('Expire','-')}")
            with c3:
                if pd.notna(row.get('Map')) and str(row.get('Map')).startswith("http"):
                    st.link_button("🚀 Google Maps", str(row.get('Map')).strip())
        st.markdown("<div style='margin-bottom:10px;'></div>", unsafe_allow_html=True)

else: # 📝 Edit Mode
    st.subheader("📝 แก้ไขและเพิ่มข้อมูล (แสดงข้อมูลทั้งหมด)")
    st.info("💡 ในโหมดนี้จะแสดงข้อมูลทุกแถว เพื่อให้เวลาบันทึกข้อมูลจะไม่หายไป")[cite: 1]
    
    # ⚠️ ใช้ df_all เสมอ เพื่อให้บันทึกแล้วข้อมูลไม่หาย
    edited_df = st.data_editor(
        df_all, 
        use_container_width=True, 
        num_rows="dynamic", # เปิดให้เพิ่มบรรทัดใหม่ได้ด้วยปุ่ม (+)[cite: 1]
        column_config={
            "Map": st.column_config.LinkColumn("Map Link", display_text="Open Maps 🌐"),
            "Status": st.column_config.SelectboxColumn("Status", options=all_statuses)
        },
        key="main_editor"
    )
    
    if st.button("💾 บันทึกข้อมูลลงฐานข้อมูล"):[cite: 1]
        save_data(edited_df)
        st.success("บันทึกข้อมูลเรียบร้อยแล้ว! (ข้อมูลครบถ้วน)")[cite: 1]
        st.rerun()
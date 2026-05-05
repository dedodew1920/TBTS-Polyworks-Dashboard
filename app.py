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
        # Read from "Polyworks Contract.xlsx" verbatim[cite: 1]
        df = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME, header=2)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        conn = sqlite3.connect(DB_NAME)
        df.to_sql("data", conn, if_exists="replace", index=False)
        conn.close()
        return True
    except Exception as e:
        st.error(f"Error syncing database: {e}")
        return False

def load_data():
    if not os.path.exists(DB_NAME): 
        sync_db_from_excel()
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql("SELECT * FROM data", conn) 
    conn.close()
    return df

def save_data(df_to_save):
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
df = load_data()

# --- Sidebar ---
st.sidebar.header("🔍 ค้นหาและตัวกรอง")
display_mode = st.sidebar.radio("โหมดการใช้งาน:", ["📦 View Mode", "📝 Edit Mode"])

if not df.empty:
    suggestions = sorted(list(set(
        df['Company'].dropna().astype(str).unique().tolist() + 
        df['Division'].dropna().astype(str).unique().tolist() + 
        df['DongleNo.'].dropna().astype(str).unique().tolist()
    )))
    search_query = st.sidebar.selectbox("🎯 ค้นหา", options=[""] + suggestions, index=0)
    all_statuses = sorted(df['Status'].dropna().unique().tolist())
    
    # --- 📅 Update Eligibility Filters (O3, P3, Q3)[cite: 1] ---
    st.sidebar.markdown("---")
    st.sidebar.subheader("📅 Update Eligibility")
    filter_2024 = st.sidebar.checkbox("2024 Can Update")
    filter_2025 = st.sidebar.checkbox("2025 Can Update")
    filter_2026 = st.sidebar.checkbox("2026 Can Update")
else:
    search_query = ""
    all_statuses = []
    filter_2024 = filter_2025 = filter_2026 = False

selected_statuses = st.sidebar.multiselect("🚦 สถานะสัญญา", options=all_statuses)

# --- Filter Logic ---
filtered_df = df.copy()

# 1. Update Eligibility Filters (Filtering for 'OK')[cite: 1]
if filter_2024:
    filtered_df = filtered_df[filtered_df['2024 Can Update'] == 'OK']
if filter_2025:
    filtered_df = filtered_df[filtered_df['2025 Can Update'] == 'OK']
if filter_2026:
    filtered_df = filtered_df[filtered_df['2026 Can Update'] == 'OK']

# 2. Search Query Filter
if search_query:
    filtered_df = filtered_df[
        filtered_df['Company'].astype(str).str.contains(search_query, case=False, na=False) |
        filtered_df['Division'].astype(str).str.contains(search_query, case=False, na=False) |
        filtered_df['DongleNo.'].astype(str).str.contains(search_query, case=False, na=False)
    ]

# 3. Status Filter
if selected_statuses:
    filtered_df = filtered_df[filtered_df['Status'].isin(selected_statuses)]

if st.sidebar.button("🔄 Sync ข้อมูลใหม่จาก Excel", use_container_width=True):
    sync_db_from_excel()
    st.cache_data.clear() 
    st.rerun()

# --- Dashboard Header ---
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

# --- 🗺️ Map Section ---
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
                dept_html = "".join([f"<div style='border-bottom:1px solid #eee; padding:3px;'><b>{r['Division']}</b>: <span style='color:{get_status_color(r['Status'])};'>{r['Status']}</span></div>" for _, r in group.iterrows()])
                marker_color = 'red' if 'expired' in group['Status'].str.lower().values else 'green'
                folium.Marker([lat, lon], popup=folium.Popup(f"🏢 <b>{company}</b><br>{dept_html}", max_width=300), icon=folium.Icon(color=marker_color)).add_to(m)
            
            st_folium(m, width="100%", height=450, key="company_map")

# --- 📦 Content Display ---
if display_mode == "📦 View Mode":
    st.subheader("📋 รายละเอียดรายแผนก")
    for idx, row in filtered_df.iterrows():
        is_expired = 'expired' in str(row['Status']).lower()
        bg_color = "#FF4B4B" if is_expired else "#00C49A"
        st.markdown(f"""
            <div style="background-color:{bg_color}; padding:10px 15px; border-radius:5px 5px 0 0; color:white; font-weight:bold;">
                🏢 {row['Company']} | 📂 {row.get('Division','-')}
            </div>
        """, unsafe_allow_html=True)
        with st.expander("รายละเอียด"):
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

else: # Edit Mode
    st.subheader("📝 แก้ไขและเพิ่มข้อมูล")
    # Clean temporary coordinates columns before editing
    df_to_edit = filtered_df.drop(columns=['Lat_f', 'Lon_f']) if 'Lat_f' in filtered_df.columns else filtered_df
    
    edited_df = st.data_editor(
        df_to_edit, 
        use_container_width=True, 
        num_rows="dynamic",
        column_config={
            "Map": st.column_config.LinkColumn("Map Link", display_text="Open Maps 🌐"),
            "Status": st.column_config.SelectboxColumn("Status", options=all_statuses)
        }
    )
    
    if st.button("💾 บันทึกข้อมูล"):
        save_data(edited_df)
        st.success("บันทึกข้อมูลเรียบร้อย!")
        st.rerun()
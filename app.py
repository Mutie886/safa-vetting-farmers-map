import streamlit as st
import pandas as pd
import numpy as np
import json
import subprocess
import sys
import streamlit.components.v1 as components

# 1. PAGE CONFIGURATION
st.set_page_config(
    page_title="SAFA — OND 2026 Vetting Farmers Dashboard",
    page_icon="🌱",
    layout="wide"
)

# Initialize Session State for Data Persistence during active session
if 'active_df' not in st.session_state:
    st.session_state.active_df = None

# 2. BRANDING & HEADER (SAFA LOGO INTEGRATION)
st.markdown("""
    <style>
        .safa-header {
            display: flex;
            align-items: center;
            background: linear-gradient(90deg, #0A3A2A 0%, #1E5E43 100%);
            padding: 18px 25px;
            border-radius: 10px;
            color: white;
            margin-bottom: 20px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .safa-logo-box {
            background-color: #FFFFFF;
            padding: 8px 14px;
            border-radius: 8px;
            display: flex;
            align-items: center;
            justify-content: center;
            margin-right: 20px;
        }
        .safa-logo-text {
            color: #0A3A2A;
            font-weight: 900;
            font-size: 26px;
            letter-spacing: 2px;
            font-family: 'Helvetica Neue', Arial, sans-serif;
        }
        .safa-logo-sub {
            color: #27AE60;
            font-size: 10px;
            font-weight: 700;
            display: block;
            margin-top: -4px;
            letter-spacing: 1px;
        }
        .safa-title {
            margin: 0;
            font-size: 24px;
            font-weight: 700;
            color: #FFFFFF !important;
        }
        .safa-subtitle {
            margin: 2px 0 0 0;
            font-size: 13px;
            color: #A3E4D7;
            font-weight: 400;
        }
    </style>
    
    <div class="safa-header">
        <div class="safa-logo-box">
            <div>
                <span class="safa-logo-text">SAFA</span>
                <span class="safa-logo-sub">SUSTAINABLE AGRI</span>
            </div>
        </div>
        <div>
            <h1 class="safa-title">OND 2026 Vetting Farmers</h1>
            <p class="safa-subtitle">Live GPS Audit & Field Analytics Dashboard</p>
        </div>
    </div>
""", unsafe_allow_html=True)

# 3. ADMIN ACCESS & UPLOAD SECTION
st.sidebar.header("🔑 Admin Access")
admin_pin = st.sidebar.text_input("Enter Admin PIN", type="password", help="Type 1234 and press ENTER")

ADMIN_PASSCODE = "1234"

# Validate passcode automatically
is_admin = admin_pin == ADMIN_PASSCODE

if is_admin:
    st.sidebar.success("✅ Admin Access Granted")
    uploaded_file = st.sidebar.file_uploader(
        "Upload New Dataset (.xlsx or .csv)", 
        type=["xlsx", "xls", "csv"]
    )
    
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                st.session_state.active_df = pd.read_csv(uploaded_file)
            else:
                try:
                    st.session_state.active_df = pd.read_excel(uploaded_file, engine='openpyxl')
                except ImportError:
                    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
                    st.session_state.active_df = pd.read_excel(uploaded_file, engine='openpyxl')
            st.sidebar.success("Data updated successfully!")
        except Exception as e:
            st.sidebar.error(f"Error reading file: {str(e)}")

# Ensure dataset exists before rendering dashboards
if st.session_state.active_df is None:
    if not is_admin:
        st.warning("⚠️ No dataset loaded. Admin must enter the PIN in the sidebar and upload the Kobo data file.")
    else:
        st.info("👈 Please use the uploader in the left sidebar to upload the initial dataset.")
    st.stop()

df_raw = st.session_state.active_df.copy()

if df_raw.empty:
    st.warning("The uploaded dataset is empty.")
    st.stop()

# 4. DATA CLEANING & TRANSFORMATION
df = df_raw.copy()

# Standardize Field Officer Names
fo_clean = {
    'Caroline': 'Caroline Kalovoto', 'Kklonzi': 'Kilonzi', 'Dou': 'Douglas',
    'Dominic': 'Dominic Kioko', 'Amani Thoya Karisa': 'Amani Thoya',
    'Paul Kamau Muraya': 'Paul Kamau', 'Paul kamau muraya': 'Paul Kamau',
    'Pk And Mary': 'PK & Mary', 'Pk and mary': 'PK & Mary', 'Peter And Mary': 'PK & Mary',
    'Dominic kioko': 'Dominic Kioko', 'Francis kisese': 'Francis Kisese',
    'Peter king\'ola': 'Peter King\'ola', 'Shadrack kieti': 'Shadrack Kieti'
}

officer_cols = [c for c in df.columns if 'officer' in c.lower() or 'Field Officer' in c]
officer_col = officer_cols[0] if officer_cols else df.columns[0]
df['officer'] = df[officer_col].astype(str).str.strip().replace(fo_clean).str.title()

# Time & Operational Date Adjustment (5 AM cutoff)
start_cols = [c for c in df.columns if 'start' in c.lower() or 'submission' in c.lower() or 'date' in c.lower()]
start_col = start_cols[0] if start_cols else df.columns[0]
df['start_dt'] = pd.to_datetime(df[start_col], errors='coerce').fillna(pd.Timestamp.now())

def assign_op_date_and_day(dt):
    op_dt = dt - pd.Timedelta(days=1) if dt.hour < 5 else dt
    return op_dt.strftime('%Y-%m-%d'), op_dt.strftime('%A')

res = df['start_dt'].apply(assign_op_date_and_day)
df['date_str'] = [r[0] for r in res]
df['day_name'] = [r[1] for r in res]
df['time_visited'] = df['start_dt'].dt.strftime('%I:%M %p')

# Farmer & Location Details
farmer_cols = [c for c in df.columns if 'farmer' in c.lower() or 'Farmer Name' in c]
farmer_col = farmer_cols[0] if farmer_cols else df.columns[0]
df['farmer'] = df[farmer_col].fillna('Unknown').astype(str).str.title()

county_cols = [c for c in df.columns if 'county' in c.lower()]
county_col = county_cols[0] if county_cols else df.columns[0]
df['county'] = df[county_col].astype(str).str.strip().str.title()

loc_cols = [c for c in df.columns if 'location' in c.lower() and 'farm' not in c.lower()]
loc_col = loc_cols[0] if loc_cols else df.columns[0]
df['location'] = df[loc_col].astype(str).str.strip().str.title()

vill_cols = [c for c in df.columns if 'village' in c.lower()]
vill_col = vill_cols[0] if vill_cols else df.columns[0]
df['village'] = df[vill_col].astype(str).str.strip().str.title()

# Acres Parsing
acre_cols = [c for c in df.columns if 'acre' in c.lower()]
if acre_cols:
    raw_acres = df[acre_cols[0]].astype(str).str.replace(',', '.', regex=False).str.strip()
    numeric_acres = pd.to_numeric(raw_acres, errors='coerce').fillna(0).clip(lower=0)
    df['acres'] = np.where(numeric_acres > 20.0, 0.0, numeric_acres)
else:
    df['acres'] = 0.0

app_cols = [c for c in df.columns if 'approve' in c.lower()]
app_col = app_cols[0] if app_cols else df.columns[0]
df['approved'] = df[app_col].fillna('Unknown').astype(str).str.title()

# Extract Latitude & Longitude
lat_cols = [c for c in df.columns if 'latitude' in c.lower() or 'lat' in c.lower()]
lon_cols = [c for c in df.columns if 'longitude' in c.lower() or 'lng' in c.lower() or 'lon' in c.lower()]

if lat_cols and lon_cols:
    df['lat'] = pd.to_numeric(df[lat_cols[0]], errors='coerce')
    df['lon'] = pd.to_numeric(df[lon_cols[0]], errors='coerce')
else:
    gps_cols = [c for c in df.columns if 'location' in c.lower() or 'gps' in c.lower() or 'geolocation' in c.lower()]
    if gps_cols:
        gps_col = gps_cols[0]
        df['lat'] = df[gps_col].apply(lambda x: float(str(x).split()[0]) if pd.notnull(x) and len(str(x).split())>=2 else np.nan)
        df['lon'] = df[gps_col].apply(lambda x: float(str(x).split()[1]) if pd.notnull(x) and len(str(x).split())>=2 else np.nan)
    else:
        df['lat'] = np.nan
        df['lon'] = np.nan

df_coords = df.dropna(subset=['lat', 'lon']).copy()

if df_coords.empty:
    st.error("No valid GPS coordinates found in file.")
    st.stop()

# 5. AUDIT CALCULATION (Haversine & Overlaps)
df_coords = df_coords.sort_values(by=['officer', 'start_dt']).reset_index(drop=True)

def haversine_km(lat1, lon1, lat2, lon2):
    if pd.isnull(lat1) or pd.isnull(lon1) or pd.isnull(lat2) or pd.isnull(lon2):
        return 0.0
    R = 6371.0
    dlat = np.radians(lat2 - lat1)
    dlon = np.radians(lon2 - lon1)
    a = np.sin(dlat / 2.0)**2 + np.cos(np.radians(lat1)) * np.cos(np.radians(lat2)) * np.sin(dlon / 2.0)**2
    c = 2 * np.arctan2(np.sqrt(a), np.sqrt(1 - a))
    return R * c

df_coords['prev_lat'] = df_coords.groupby(['officer', 'date_str'])['lat'].shift(1)
df_coords['prev_lon'] = df_coords.groupby(['officer', 'date_str'])['lon'].shift(1)
df_coords['dist_from_prev_km'] = df_coords.apply(lambda r: haversine_km(r['prev_lat'], r['prev_lon'], r['lat'], r['lon']), axis=1)

def flag_status(row):
    if pd.notnull(row['dist_from_prev_km']) and row['dist_from_prev_km'] < 0.01:
        return "⚠️ Suspicious (0km from last farm)"
    elif pd.notnull(row['dist_from_prev_km']) and row['dist_from_prev_km'] > 50:
        return "🚨 GPS Outlier (>50km jump)"
    return "✅ Valid GPS Distance"

df_coords['audit_status'] = df_coords.apply(flag_status, axis=1)
df_coords['dist_from_prev_km'] = df_coords['dist_from_prev_km'].fillna(0).round(2)
df_coords['coord_key'] = df_coords['lat'].round(5).astype(str) + '_' + df_coords['lon'].round(5).astype(str)

# 6. STREAMLIT FILTERS
st.sidebar.header("🔍 Filters")
counties = ["All"] + sorted(list(df_coords['county'].unique()))
sel_county = st.sidebar.selectbox("Select County", counties)

dates = ["All"] + sorted(list(df_coords['date_str'].unique()), reverse=True)
sel_date = st.sidebar.selectbox("Select Date", dates)

filtered_df = df_coords.copy()
if sel_county != "All":
    filtered_df = filtered_df[filtered_df['county'] == sel_county]
if sel_date != "All":
    filtered_df = filtered_df[filtered_df['date_str'] == sel_date]

officers = ["All"] + sorted(list(filtered_df['officer'].unique()))
sel_officer = st.sidebar.selectbox("Select Field Officer", officers)

if sel_officer != "All":
    filtered_df = filtered_df[filtered_df['officer'] == sel_officer]

# Metrics Overview Bar
col1, col2, col3, col4 = st.columns(4)
col1.metric("Farmers Vetted", len(filtered_df))
col2.metric("Acres Vetted", f"{filtered_df['acres'].sum():.1f}")
col3.metric("Farmers Approved", len(filtered_df[filtered_df['approved'].isin(['Yes', 'Approved'])]))

overlap_groups_count = (filtered_df.groupby('coord_key').size() > 1).sum()
col4.metric("GPS Overlap Groups", overlap_groups_count)

# 7. LEAFLET MAP GENERATION
records = filtered_df[['farmer', 'officer', 'county', 'location', 'village', 'acres', 'approved', 'lat', 'lon', 'date_str', 'day_name', 'time_visited', 'audit_status', 'dist_from_prev_km', 'coord_key']].to_dict(orient='records')
records_json = json.dumps(records)

html_code = f"""
<!DOCTYPE html>
<html>
<head>
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <style>
        body {{ margin: 0; padding: 0; font-family: Arial; }}
        #map {{ height: 680px; width: 100%; border-radius: 8px; }}
        .warning {{ color: #C00000; font-weight: bold; }}
    </style>
</head>
<body>
<div id="map"></div>
<script>
    var rawData = {records_json};
    var map = L.map('map').setView([-2.4, 38.0], 8);
    L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{ maxZoom: 18 }}).addTo(map);

    var layerGroup = L.layerGroup().addTo(map);
    var colors = ['#e6194b', '#3cb44b', '#ffe119', '#4363d8', '#f58231', '#911eb4', '#46f0f0', '#f032e6', '#008080', '#9a6324', '#e6beff', '#800000'];
    var officers = [...new Set(rawData.map(d => d.officer))].sort();
    var colorMap = {{}};
    officers.forEach((o, i) => colorMap[o] = colors[i % colors.length]);

    var groupedCoords = {{}};
    rawData.forEach(d => {{
        if(!groupedCoords[d.coord_key]) groupedCoords[d.coord_key] = [];
        groupedCoords[d.coord_key].push(d);
    }});

    Object.keys(groupedCoords).forEach(key => {{
        var items = groupedCoords[key];
        items.forEach((d, idx) => {{
            var c = colorMap[d.officer];
            var isWarn = d.audit_status.includes('⚠️') || d.audit_status.includes('🚨');
            
            var lat = d.lat, lon = d.lon;
            if (items.length > 1 && idx > 0) {{
                var angle = idx * (2 * Math.PI / 6);
                var radius = 0.00025 * Math.ceil(idx / 6);
                lat += radius * Math.cos(angle);
                lon += radius * Math.sin(angle);
            }}

            var marker = L.circleMarker([lat, lon], {{
                radius: items.length > 1 ? 7 : 6,
                fillColor: c,
                color: isWarn ? '#ff0000' : '#000',
                weight: isWarn ? 2 : 1,
                fillOpacity: 0.85
            }});

            var popup = `<div style="font-size:12px; width:220px;">
                <b style="color:#0A3A2A;">${{d.farmer}}</b> ${{items.length > 1 ? `<span style="background:#e1f5fe; color:#0288d1; padding:2px 4px; border-radius:3px; font-size:10px; float:right;">Pt ${{idx+1}} of ${{items.length}}</span>` : ''}}<br/>
                <b>Officer:</b> ${{d.officer}}<br/>
                <b>Date:</b> ${{d.date_str}} (${{d.day_name}})<br/>
                <b>Time:</b> ${{d.time_visited}}<br/><hr style="margin:4px 0;"/>
                <b>County:</b> ${{d.county}}<br/>
                <b>Loc/Vill:</b> ${{d.location}} / ${{d.village}}<br/>
                <b>Acres:</b> ${{d.acres.toFixed(1)}} (App: ${{d.approved}})<br/><hr style="margin:4px 0;"/>
                <b>Status:</b> <span class="${{isWarn ? 'warning' : ''}}">${{d.audit_status}}</span><br/>
                <i>Dist from prev: ${{d.dist_from_prev_km}} km</i>
            </div>`;

            marker.bindPopup(popup);
            layerGroup.addLayer(marker);
        }});
    }});

    if(rawData.length > 0) {{
        var bounds = rawData.map(d => [d.lat, d.lon]);
        map.fitBounds(bounds, {{padding: [30, 30], maxZoom: 14}});
    }}
</script>
</body>
</html>
"""

st.subheader("Interactive Farmer Map")
components.html(html_code, height=700)

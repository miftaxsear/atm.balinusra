import streamlit as st
import sqlite3
import pandas as pd
import os
import json
import math

# Konfigurasi halaman agar memenuhi lebar layar penuh
st.set_page_config(page_title="Peta ATM (OpenStreetMap)", layout="wide", initial_sidebar_state="expanded")

# ==========================================
# PENGAMAN HALAMAN: CEK STATUS LOGIN
# ==========================================
# Cek apakah session 'logged_in' ada dan bernilai True
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if not st.session_state["logged_in"]:
    st.warning("⚠️ Anda belum melakukan Login. Silakan masuk terlebih dahulu melalui halaman utama.")
    if st.button("Kembali ke Halaman Utama / Login"):
        st.switch_page("app.py")
    st.stop()

# ==========================================
# FUNGSI LOAD DATA ATM TERPUSAT (EXCEL & SQLITE)
# ==========================================
@st.cache_data
def load_all_atms():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    parent_dir = os.path.dirname(current_dir)
    
    # Menentukan path secara dinamis berdasarkan tempat aplikasi berjalan
    if os.path.exists(r"c:\webmaster"):
        excel_path = r"c:\webmaster\Master_Lokasi_ATM.xlsx"
        db_path = r"c:\webmaster\atm_database.db"
    else:
        excel_path = os.path.join(parent_dir, "Master_Lokasi_ATM.xlsx")
        db_path = os.path.join(parent_dir, "atm_database.db")
        if not os.path.exists(excel_path):
            excel_path = os.path.join(current_dir, "Master_Lokasi_ATM.xlsx")
            db_path = os.path.join(current_dir, "atm_database.db")

    df = pd.DataFrame()
    
    # 1. Membaca file Excel Master Lokasi
    if os.path.exists(excel_path):
        try:
            df = pd.read_excel(excel_path)
        except Exception as e:
            st.error(f"Gagal membaca file Excel Master Lokasi di {excel_path}: {e}")
            return pd.DataFrame()
    else:
        st.error(f"File master lokasi tidak ditemukan di path: {excel_path}. Pastikan file sudah tersedia.")
        return pd.DataFrame()

    if not df.empty:
        df = df.loc[:, ~df.columns.duplicated()]
        
        # Mapping kolom agar seragam dengan sistem peta
        col_map = {}
        for col in df.columns:
            lc = col.lower().strip()
            if lc in ['ws_id', 'wsid', 'id']:
                col_map[col] = 'Wsid'
            elif lc in ['ws_name', 'name', 'lokasi']:
                col_map[col] = 'Ws Name'
            elif lc in ['serial_no', 'sn', 'serial']:
                col_map[col] = 'Serial_No'
            elif 'prov' in lc:
                col_map[col] = 'Province'
            elif lc in ['city', 'kot', 'kota']:
                col_map[col] = 'City'
            elif 'svc' in lc and 'area' in lc:
                col_map[col] = 'SVC_AREA'
            elif 'se' in lc and 'name' in lc:
                col_map[col] = 'SE_NAME'
            elif 'mtype' in lc or lc == 'm_type':
                col_map[col] = 'MTYPE'
            elif lc in ['latitude', 'lat', 'gps_lat']:
                col_map[col] = 'Gps Lat'
            elif lc in ['longitude', 'long', 'lon', 'gps_long']:
                col_map[col] = 'Gps Long'
            elif lc in ['address', 'alamat']:
                col_map[col] = 'Gps Address'
                
        df = df.rename(columns=col_map)
        df = df.loc[:, ~df.columns.duplicated()]

        # 2. Menyimpan/Sinkronisasi otomatis ke database SQLite di server
        try:
            conn = sqlite3.connect(db_path)
            df.to_sql('atm_master', conn, if_exists='replace', index=False)
            conn.close()
        except Exception as e:
            st.warning(f"Catatan: Gagal menyalin ke database SQLite, menggunakan data Excel langsung. Error: {e}")

        return df
        
    return pd.DataFrame()

def calculate_distance(lat1, lon1, lat2, lon2):
    R = 6371.0
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat / 2)**2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2)**2
    c = 2 * math.asin(math.sqrt(a))
    return R * c

df_atms = load_all_atms()

if df_atms.empty:
    st.warning("⚠️ Data ATM tidak ditemukan atau gagal dimuat dari database.")
else:
    if 'Gps Lat' in df_atms.columns and 'Gps Long' in df_atms.columns:
        df_atms['Gps Lat'] = pd.to_numeric(df_atms['Gps Lat'], errors='coerce')
        df_atms['Gps Long'] = pd.to_numeric(df_atms['Gps Long'], errors='coerce')
        df_valid = df_atms.dropna(subset=['Gps Lat', 'Gps Long']).copy()
    else:
        df_valid = pd.DataFrame()

    if not df_valid.empty:
        wsid_series = df_valid['Wsid'] if 'Wsid' in df_valid.columns else pd.Series()
        if isinstance(wsid_series, pd.DataFrame):
            wsid_series = wsid_series.iloc[:, 0]

        # Inisialisasi session state
        if "zonasi_text" not in st.session_state:
            st.session_state["zonasi_text"] = ""
        if "search_wsid" not in st.session_state:
            st.session_state["search_wsid"] = ""
        if "selected_filter" not in st.session_state:
            st.session_state["selected_filter"] = "- Tampilkan Semua -"
        if "label_by" not in st.session_state:
            st.session_state["label_by"] = "Province"

        # CSS Tombol Warna Gold (#daa520)
        st.sidebar.markdown("""
        <style>
            div.stButton > button {
                background-color: #daa520 !important;
                color: white !important;
                border: 1px solid #b8860b !important;
                font-weight: bold !important;
            }
            div.stButton > button:hover {
                background-color: #b8860b !important;
                color: white !important;
            }
        </style>
        """, unsafe_allow_html=True)

        tab_jelajahi, tab_zonasi = st.sidebar.tabs(["Balinusra Site", "Zonasi Marker ATM"])
        
        filtered_df = df_valid.copy()
        cluster_enabled = True
        zonasi_active_df = pd.DataFrame()
        route_segments = []
        total_route_distance = 0.0
        missing_wsids = []
        
       # --- TAB 1: BALINUSRA SITE ---
        with tab_jelajahi:
            st.markdown("### Balinusra Site")
            st.markdown("---")
            
            group_options = ["Province", "City", "SVC_AREA", "SE_NAME", "MTYPE"]
            available_group_options = [opt for opt in group_options if opt in df_valid.columns]
            if not available_group_options:
                available_group_options = ["Province"]
                
            if st.session_state["label_by"] not in available_group_options:
                st.session_state["label_by"] = available_group_options[0]

            label_by = st.selectbox(
                "ATM site berdasarkan", 
                available_group_options, 
                key="label_by"
            )
            
            filter_values = ["- Tampilkan Semua -"] + sorted(list(df_valid[label_by].dropna().astype(str).unique())) if label_by in df_valid.columns else ["- Tampilkan Semua -"]
            
            if st.session_state["selected_filter"] not in filter_values:
                st.session_state["selected_filter"] = "- Tampilkan Semua -"

            selected_filter = st.selectbox(
                "Filter Nilai (opsional)", 
                filter_values, 
                key="selected_filter"
            )
            
            st.markdown("##### Cari WSID")
            search_wsid = st.text_input(
                "Masukkan WSID...", 
                key="search_wsid", 
                label_visibility="collapsed"
            )
            
            def clear_filter_callback():
                st.session_state["label_by"] = "Province"
                st.session_state["selected_filter"] = "- Tampilkan Semua -"
                st.session_state["search_wsid"] = ""

            col_s1, col_s2 = st.columns(2)
            btn_cari = col_s1.button("Cari", use_container_width=True)
            btn_clear = col_s2.button("Clear", use_container_width=True, on_click=clear_filter_callback)

            cluster_enabled = st.checkbox("Cluster marker", value=True)
            
            def reset_all_filters():
                st.session_state["label_by"] = "Province"
                st.session_state["selected_filter"] = "- Tampilkan Semua -"
                st.session_state["search_wsid"] = ""
                st.session_state["zonasi_text"] = ""
                st.cache_data.clear()

            st.markdown("---")
            if st.button("Muat Ulang", use_container_width=True, on_click=reset_all_filters):
                st.rerun()
                
            st.button("Fit to Markers", use_container_width=True)

            if st.session_state["selected_filter"] != "- Tampilkan Semua -" and label_by in filtered_df.columns:
                filtered_df = filtered_df[filtered_df[label_by].astype(str) == st.session_state["selected_filter"]]
                
            if st.session_state["search_wsid"].strip() and 'Wsid' in filtered_df.columns:
                f_wsid = filtered_df['Wsid']
                if isinstance(f_wsid, pd.DataFrame):
                    f_wsid = f_wsid.iloc[:, 0]
                filtered_df = filtered_df[f_wsid.astype(str).str.contains(st.session_state["search_wsid"].strip(), case=False, na=False)]

            if label_by in df_valid.columns:
                cat_series = df_valid[label_by]
                if isinstance(cat_series, pd.DataFrame):
                    cat_series = cat_series.iloc[:, 0]
                cat_counts = cat_series.value_counts().reset_index()
                cat_counts.columns = ['Category', 'Count']
                
                total_all_atms = cat_counts['Count'].sum()

                st.markdown("---")
                
                header_html = f"""
                <div style='font-size: 15px; font-weight: bold; color: #262730; margin-bottom: 8px; display: flex; justify-content: space-between; align-items: center;'>
                    <span>Berdasarkan ({label_by}):</span>
                    <span style='background-color: #f0f2f6; padding: 1px 6px; border-radius: 4px; font-weight: bold;'>{total_all_atms}</span>
                </div>
                """
                st.markdown(header_html, unsafe_allow_html=True)
                
                colors = ["🔴", "🟢", "🔵", "🟣", "🟠", "🟡", "🟤", "⚫"]
                legend_html = "<div style='font-size: 13px; line-height: 1.8;'>"
                for idx, row in cat_counts.iterrows():
                    bullet_color = colors[idx % len(colors)]
                    legend_html += f"<div style='display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;'><span>{bullet_color} <b>{row['Category']}</b></span> <span style='background-color: #f0f2f6; padding: 1px 6px; border-radius: 4px; font-weight: bold;'>{row['Count']}</span></div>"
                legend_html += "</div>"
                
                st.markdown(legend_html, unsafe_allow_html=True)

        # --- TAB 2: ZONASI MARKER ATM ---
        with tab_zonasi:
            st.markdown("""
                <style>
                    div[data-testid="stHorizontalBlock"] > div:nth-child(1) button {
                        background-color: #27ae60 !important;
                        color: white !important;
                        border-color: #219653 !important;
                    }
                    div[data-testid="stHorizontalBlock"] > div:nth-child(1) button:hover {
                        background-color: #219653 !important;
                        color: white !important;
                    }
                    div[data-testid="stHorizontalBlock"] > div:nth-child(2) button {
                        background-color: #e74c3c !important;
                        color: white !important;
                        border-color: #c0392b !important;
                    }
                    div[data-testid="stHorizontalBlock"] > div:nth-child(2) button:hover {
                        background-color: #c0392b !important;
                        color: white !important;
                    }
                </style>
            """, unsafe_allow_html=True)

            st.markdown("### Zonasi Marker ATM")
            st.markdown("---")
            st.markdown("##### Daftar WSID")

            def clear_route_callback():
                st.session_state["zonasi_text"] = ""

            col_z1, col_z2 = st.columns(2)
            btn_route = col_z1.button("Tampilkan Rute", use_container_width=True)
            btn_clear_route = col_z2.button("Hapus Rute", use_container_width=True, on_click=clear_route_callback)

            placeholder_text = "Masukkan WSID (satu per baris atau comma-separated)\nContoh:\n0001\n0002\n0003"
            zonasi_input = st.text_area(
                "Daftar WSID",
                key="zonasi_text",
                height=130,
                placeholder=placeholder_text,
                label_visibility="collapsed"
            )
            st.markdown("<small>Masukkan 2 atau lebih WSID untuk menampilkan rute.</small>", unsafe_allow_html=True)

            if zonasi_input.strip():
                raw_wsids = [w.strip() for w in zonasi_input.replace(',', '\n').split('\n') if w.strip()]
                if 'Wsid' in df_valid.columns:
                    temp_list = []
                    for w in raw_wsids:
                        match = df_valid[wsid_series.astype(str).str.strip().str.upper() == w.upper()]
                        if not match.empty:
                            temp_list.append(match.iloc[0])
                        else:
                            missing_wsids.append(w)
                            
                    if temp_list:
                        zonasi_active_df = pd.DataFrame(temp_list)
                        zonasi_active_df = zonasi_active_df.loc[:, ~zonasi_active_df.columns.duplicated()]
                        filtered_df = zonasi_active_df.copy()

                st.markdown("---")

                if missing_wsids:
                    missing_str = ", ".join(missing_wsids)
                    st.error(f"Wsid tidak ditemukan ({missing_str}), cek kembali")

                if len(filtered_df) >= 2:
                    coords = list(zip(filtered_df['Gps Lat'], filtered_df['Gps Long']))
                    st.markdown("##### Informasi Rute:")
                    
                    segment_distances = []
                    for i in range(len(coords) - 1):
                        dist = calculate_distance(coords[i][0], coords[i][1], coords[i+1][0], coords[i+1][1])
                        segment_distances.append(dist)
                        total_route_distance += dist
                        
                        curr_w = filtered_df.iloc[i].get('Wsid', '-')
                        if isinstance(curr_w, pd.Series): curr_w = curr_w.iloc[0]
                        next_w = filtered_df.iloc[i+1].get('Wsid', '-')
                        if isinstance(next_w, pd.Series): next_w = next_w.iloc[0]
                        route_segments.append(f"{i+1}. {curr_w} -> {next_w}: {dist:.2f} km")

                    for i in range(len(coords)):
                        row = filtered_df.iloc[i]
                        wsid_val = row.get('Wsid', '-')
                        if isinstance(wsid_val, pd.Series):
                            wsid_val = wsid_val.iloc[0]
                        addr_val = row.get('Gps Address', '-')
                        city_val = row.get('City', '-')
                        prov_val = row.get('Province', '-')
                        
                        segment_dist_str = ""
                        if i > 0:
                            dist_val = segment_distances[i-1]
                            segment_dist_str = f"<span style='background-color: #0066ff; color: white; padding: 2px 8px; border-radius: 4px; font-size: 11px; float: right;'>{dist_val:.2f} km</span>"

                        stop_box = f"""
                        <div style='background-color: #f8f9fa; border: 1px solid #e0e0e0; padding: 10px; border-radius: 6px; margin-bottom: 8px;'>
                            <b>Stop {i+1} - {wsid_val}</b> {segment_dist_str}<br>
                            <span style='font-size: 11px; color: #555;'>{addr_val}</span><br>
                            <span style='font-size: 11px; color: #777;'>{city_val}, {prov_val}</span>
                        </div>
                        """
                        st.markdown(stop_box, unsafe_allow_html=True)

                    breakdown_html = "<br>".join([f"&nbsp;&nbsp;{seg}" for seg in route_segments])
                    summary_box = f"""
                    <div style='background-color: #e6f0ff; border: 1px solid #b8daff; padding: 12px; border-radius: 6px; margin-top: 10px; font-size: 12px;'>
                        <b>Total Rute: {len(filtered_df)} Lokasi</b><br>
                        <b>Total Jarak: {total_route_distance:.2f} km</b><br>
                        <span style='color: #444; font-weight: bold;'>Breakdown per segment:</span><br>
                        {breakdown_html}
                    </div>
                    """
                    st.markdown(summary_box, unsafe_allow_html=True)

        # --- MAIN CONTENT AREA (PETA OPENSTREETMAP) ---
        st.markdown("### Peta ATM (OpenStreetMap)")

        target_cols = ['Wsid', 'Ws Name', 'Serial_No', 'Province', 'City', 'SVC_AREA', 'SE_NAME', 'MTYPE', 'Gps Lat', 'Gps Long', 'Gps Address']
        valid_target_cols = [c for c in target_cols if c in filtered_df.columns]
        
        clean_filtered_df = filtered_df[valid_target_cols].copy()
        clean_filtered_df = clean_filtered_df.loc[:, ~clean_filtered_df.columns.duplicated()]
        for col in clean_filtered_df.columns:
            if isinstance(clean_filtered_df[col], pd.DataFrame):
                clean_filtered_df[col] = clean_filtered_df[col].iloc[:, 0]

        map_data_json = clean_filtered_df.to_dict(orient='records')
        is_single_search = bool(st.session_state["search_wsid"].strip() and len(clean_filtered_df) == 1)

        leaflet_html = f"""
        <!DOCTYPE html>
        <html lang="id">
        <head>
            <meta charset="utf-8">
            <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
            <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
            <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.4.1/dist/MarkerCluster.css" />
            <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.4.1/dist/MarkerCluster.Default.css" />
            <script src="https://unpkg.com/leaflet.markercluster@1.4.1/dist/leaflet.markercluster.js"></script>
            <style>
                body {{ margin: 0; padding: 0; font-family: sans-serif; }}
                #map {{ width: 100%; height: 750px; border-radius: 4px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }}
                .leaflet-popup-content {{ font-size: 13px; line-height: 1.5; min-width: 220px; }}
                .btn-map {{ display: inline-block; padding: 6px 12px; margin-top: 8px; margin-right: 5px; font-size: 12px; font-weight: bold; text-align: center; text-decoration: none; border-radius: 4px; color: white !important; }}
                .btn-gmaps {{ background-color: #1a73e8; }}
                .btn-osm {{ background-color: #34a853; }}
            </style>
        </head>
        <body>
            <div id="map"></div>
            <script>
                var atmData = {json.dumps(map_data_json)};
                var map = L.map('map').setView([-2.548926, 118.014863], 5);

                L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
                    maxZoom: 19,
                    attribution: '&copy; OpenStreetMap'
                }}).addTo(map);

                var useCluster = {str(cluster_enabled).lower()};
                var markersGroup = useCluster ? L.markerClusterGroup() : L.layerGroup();
                var singleMarkerGroup = new L.FeatureGroup();
                var latLngs = [];
                var firstMarker = null;

                atmData.forEach(function(atm) {{
                    if (atm['Gps Lat'] && atm['Gps Long']) {{
                        var wsid = atm['Wsid'] || '-';
                        var wsName = atm['Ws Name'] || '-'; // Nama Lokasi dari Kolom C (WS_NAME)
                        var sn = atm['Serial_No'] || '-';
                        var mtype = atm['MTYPE'] || '-';
                        var address = atm['Gps Address'] || '-';
                        var city = atm['City'] || '-';
                        var province = atm['Province'] || '-';
                        var svcArea = atm['SVC_AREA'] || '-';
                        var seName = atm['SE_NAME'] || '-';
                        var lat = atm['Gps Lat'];
                        var lng = atm['Gps Long'];

                        latLngs.push([lat, lng]);

                        var gmapsUrl = "https://www.google.com/maps/search/?api=1&query=" + lat + "," + lng;
                        var osmUrl = "https://www.openstreetmap.org/?mlat=" + lat + "&mlon=" + lng + "#map=18/" + lat + "/" + lng;

                        // Menambahkan LOKASI (wsName) di sebelah WSID pada Popup
                        var popupContent = "<b>WSID: " + wsid + " &nbsp;|&nbsp; Lokasi: " + wsName + "</b><br>" +
                                           "<b>SN:</b> " + sn + " &nbsp;|&nbsp; <b>Tipe:</b> " + mtype + "<hr style='margin: 5px 0;'>" +
                                           "<b>Address</b><br>" + address + "<br><br>" +
                                           "<b>City:</b> " + city + "<br>" +
                                           "<b>Province:</b> " + province + "<br>" +
                                           "<b>SVC Area:</b> " + svcArea + "<br>" +
                                           "<b>SE Name:</b> " + seName + "<br>" +
                                           "<div style='margin-top: 8px;'>" +
                                           "<a href='" + gmapsUrl + "' target='_blank' class='btn-map btn-gmaps'>Google Maps</a>" +
                                           "<a href='" + osmUrl + "' target='_blank' class='btn-map btn-osm'>OpenStreetMap</a>" +
                                           "</div>";
                        
                        var marker = L.marker([lat, lng]);
                        marker.bindPopup(popupContent);
                        markersGroup.addLayer(marker);
                        singleMarkerGroup.addLayer(marker);

                        if (!firstMarker) {{
                            firstMarker = marker;
                        }}
                    }}
                }});

                map.addLayer(markersGroup);

                var isZonasiActive = "{str(len(zonasi_active_df) > 0)}".toLowerCase() === "true";
                var isSingleSearch = {str(is_single_search).lower()};

                if (isZonasiActive && latLngs.length > 0) {{
                    if (latLngs.length >= 2) {{
                        var polyline = L.polyline(latLngs, {{color: '#8e44ad', weight: 6, opacity: 0.9, dashArray: '5, 10'}}).addTo(map);
                    }}
                    if (singleMarkerGroup.getLayers().length > 0) {{
                        map.fitBounds(singleMarkerGroup.getBounds().pad(0.4));
                    }}
                }} else if (isSingleSearch && firstMarker) {{
                    map.setView(firstMarker.getLatLng(), 18);
                    firstMarker.openPopup();
                }} else if (atmData.length > 0 && atmData.length < 500) {{
                    var group = new L.featureGroup(markersGroup.getLayers ? markersGroup.getLayers() : []);
                    if (group.getBounds().isValid()) {{
                        map.fitBounds(group.getBounds().pad(0.1));
                    }}
                }}
            </script>
        </body>
        </html>
        """

        st.components.v1.html(leaflet_html, height=780, scrolling=False)
    else:
        st.warning("Data koordinat (Latitude/Longitude) tidak valid atau kosong pada database.")
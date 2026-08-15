import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
import plotly.express as px

# ---------------------------------------------------------
# 1. TETAPAN HALAMAN
# ---------------------------------------------------------
st.set_page_config(
    page_title="Sistem Automasi Klimatologi MetMalaysia",
    layout="wide",
    page_icon="🌤️"
)

# ---------------------------------------------------------
# 2. KAMUS BAHASA (BILINGUAL DICTIONARY)
# ---------------------------------------------------------
TEXTS = {
    "BM": {
        "portal_tag": "PORTAL DALAMAN RASMI",
        "agency_title": "JABATAN METEOROLOGI MALAYSIA",
        "branch_title": "Pejabat Meteorologi Sabah | Sistem Automasi Analisis Klimatologi",
        "nav_home": "UTAMA",
        "nav_rain": "HUJAN (RAINFALL)",
        "nav_temp": "SUHU (TEMPERATURE)",
        "nav_qc": "SEMAKAN KUALITI & WMO",
        "sidebar_header": "📁 PANEL KAWALAN & DATA",
        "upload_label": "Muat naik fail raw AAWS (.xls / .xlsx)",
        "upload_success": "✅ Berjaya memuat naik {count} fail AAWS!",
        "qc_mode_label": "⚙️ Piawaian Penapisan Data Hilang (WMO):",
        "home_card_title": "📌 Sistem Automasi Rekod & Verifikasi Klimatologi",
        "home_card_desc": """
        Sistem ini dibangunkan untuk memproses, memformat, dan mengesahkan data siri masa daripada stesen **Automatic Agricultural Weather Station (AAWS)** ke format piawai rekod klimatologi kebangsaan mengikut garis panduan **WMO-No. 1203**.
        """,
        "home_features": """
        * 📊 **Penjanaan Format Grid Piawai:** Menyusun siri masa kepada grid 31 hari $\\times$ 12 bulan (JAN–DEC) bagi setiap tahun stesen.
        * 🛡️ **Pemeriksaan Integriti WMO:** Menapis data hilang (*missing data*) mengikut piawaian antarabangsa sebelum statistik dihitung.
        * 📈 **Visualisasi Interaktif:** Matriks keamatan hujan harian (*Heatmap*) dan siri masa tahunan yang menyokong pelbagai skim warna.
        """,
        "wmo_announcement_title": "📢 Piawaian WMO-No. 1203",
        "wmo_announcement_desc": "Garis panduan rasmi *Calculation of Climate Normals* digunapakai untuk tapisan kesempurnaan data.",
        "btn_download_wmo": "📥 Muat Turun WMO-No. 1203 (PDF)",
        "stations_ready": "📌 {count} Stesen Tersedia",
        "download_zip": "📦 MUAT TURUN SEMUA (.ZIP)",
        "zip_filename": "Laporan_Klimatologi_Semua_Stesen.zip",
        "select_station": "Pilih Stesen:",
        "station_name": "Nama Stesen",
        "record_period": "Sela Masa Rekod",
        "total_records": "Jumlah Rekod",
        "completeness_rate": "Kesempurnaan Data",
        "invalid_months": "Bulan Tidak Sah (Incomplete)",
        "alert_incomplete": "⚠️ **Perhatian Pegawai:** Terdapat **{count} bulan** gagal melepasi piawaian ({rule}). Jumlah bulanan ditandakan sebagai `N.A (Incomplete)`.",
        "subtab_form": "📄 Borang Rekod & Muat Turun",
        "subtab_charts": "📈 Visualisasi Interaktif",
        "download_excel": "📥 Muat Turun Excel ({station})",
        "preview_title": "#### 🔍 Pratonton Data Terkini (20 Baris Terawal)",
        "plot_type": "📊 Jenis Graf:",
        "color_theme": "🎨 Tema Warna:",
        "opt_heatmap": "🌧️ Matriks Harian Hujan (Heatmap)",
        "opt_trend": "📈 Trend Jumlah Hujan Tahunan (Time-Series)",
        "opt_normals": "📊 Purata Taburan Hujan Bulanan (Normals)",
        "select_year_heat": "Pilih Tahun Heatmap:",
        "heat_title": "Matriks Keamatan Hujan Harian bagi {station} ({year})",
        "axis_month": "Bulan",
        "axis_day": "Hari",
        "axis_rain": "Hujan (mm)",
        "axis_year": "Tahun",
        "axis_avg_rain": "Purata Hujan (mm)",
        "trend_title": "Trend Jumlah Hujan Tahunan bagi {station} ({min_yr} - {max_yr})",
        "trend_avg_label": "Purata",
        "norm_title": "Corak Purata Hujan Bulanan (Normal Profile) bagi {station}",
        "temp_title": "🌡️ Modul Analisis Suhu Udara (Temperature)",
        "temp_desc": "Modul Suhu Udara sedia diselaraskan mengikut format cerapan piawai WMO.",
        "qc_title": "📋 Laporan Audit Kualiti Data & Integriti WMO",
        "qc_select_station": "Pilih Stesen untuk Log Audit:",
        "qc_filter_failed": "🔍 Paparkan Bulan Tidak Sah / Ada Data Hilang Sahaja",
        "download_qc_csv": "📥 Muat Turun Log Audit (.CSV)",
        "qc_col_year": "Tahun",
        "qc_col_month": "Bulan",
        "qc_col_na": "Hari Hilang (NA)",
        "qc_col_consec": "Maks. Berturut-turut",
        "qc_col_status": "Status Integriti",
        "qc_col_action": "Tindakan Pengiraan",
        "qc_status_valid": "✅ Sah (Valid)",
        "qc_status_incomp": "❌ Tidak Lengkap (Incomplete)",
        "qc_status_perfect": "🟢 100% Lengkap",
        "qc_act_calc": "Dikira (Abaikan NA)",
        "qc_act_reject": "Ditolak (N.A Incomplete)",
        "info_upload": "👈 Sila muat naik fail raw AAWS di menu tepi untuk memulakan pemprosesan."
    },
    "EN": {
        "portal_tag": "OFFICIAL INTERNAL PORTAL",
        "agency_title": "MALAYSIAN METEOROLOGICAL DEPARTMENT",
        "branch_title": "Sabah Meteorological Office | Climatology Analysis Automation System",
        "nav_home": "HOME",
        "nav_rain": "RAINFALL",
        "nav_temp": "TEMPERATURE",
        "nav_qc": "DATA QC & WMO",
        "sidebar_header": "📁 CONTROL PANEL & DATA",
        "upload_label": "Upload raw AAWS files (.xls / .xlsx)",
        "upload_success": "✅ Successfully uploaded {count} AAWS file(s)!",
        "qc_mode_label": "⚙️ Missing Data Standard (WMO):",
        "home_card_title": "📌 Climatology Records Automation & Verification",
        "home_card_desc": """
        This system processes, structures, and validates time-series data from **Automatic Agricultural Weather Stations (AAWS)** into national standardized climatological record sheets in compliance with **WMO-No. 1203** guidelines.
        """,
        "home_features": """
        * 📊 **Standard Grid Generation:** Restructures data into 31-day $\\times$ 12-month (JAN-DEC) grids for each year.
        * 🛡️ **WMO Integrity Screening:** Filters missing observations before statistical computation.
        * 📈 **Interactive Analytics:** Daily rainfall intensity heatmaps and time-series trends supporting customizable color themes.
        """,
        "wmo_announcement_title": "📢 WMO-No. 1203 Standard",
        "wmo_announcement_desc": "Official WMO guidelines for Calculation of Climate Normals applied for data completeness verification.",
        "btn_download_wmo": "📥 Download WMO-No. 1203 (PDF)",
        "stations_ready": "📌 {count} Station(s) Ready",
        "download_zip": "📦 DOWNLOAD ALL (.ZIP)",
        "zip_filename": "All_Stations_Climatology_Reports.zip",
        "select_station": "Select Station:",
        "station_name": "Station Name",
        "record_period": "Record Period",
        "total_records": "Total Records",
        "completeness_rate": "Data Completeness",
        "invalid_months": "Invalid Months (Incomplete)",
        "alert_incomplete": "⚠️ **Officer Advisory:** There are **{count} month(s)** failing completeness standards ({rule}). Monthly totals are marked as `N.A (Incomplete)`.",
        "subtab_form": "📄 Record Sheets & Download",
        "subtab_charts": "📈 Interactive Analytics",
        "download_excel": "📥 Download Excel ({station})",
        "preview_title": "#### 🔍 Raw Data Preview (First 20 Records)",
        "plot_type": "📊 Chart Type:",
        "color_theme": "🎨 Color Theme:",
        "opt_heatmap": "🌧️ Daily Rainfall Matrix (Heatmap)",
        "opt_trend": "📈 Annual Total Rainfall Trend (Time-Series)",
        "opt_normals": "📊 Monthly Average Rainfall (Normals)",
        "select_year_heat": "Select Year for Heatmap:",
        "heat_title": "Daily Rainfall Intensity Matrix for {station} ({year})",
        "axis_month": "Month",
        "axis_day": "Day",
        "axis_rain": "Rainfall (mm)",
        "axis_year": "Year",
        "axis_avg_rain": "Average Rainfall (mm)",
        "trend_title": "Annual Total Rainfall Trend for {station} ({min_yr} - {max_yr})",
        "trend_avg_label": "Average",
        "norm_title": "Monthly Average Rainfall Pattern (Normal Profile) for {station}",
        "temp_title": "🌡️ Air Temperature Analysis Module",
        "temp_desc": "The Temperature Module is structured for standard WMO meteorological parameters.",
        "qc_title": "📋 Data Quality Audit Report & WMO Integrity",
        "qc_select_station": "Select Station for Audit Log:",
        "qc_filter_failed": "🔍 Display Incomplete / Missing Data Months Only",
        "download_qc_csv": "📥 Download Audit Log (.CSV)",
        "qc_col_year": "Year",
        "qc_col_month": "Month",
        "qc_col_na": "Missing Days (NA)",
        "qc_col_consec": "Max Consecutive",
        "qc_col_status": "Integrity Status",
        "qc_col_action": "Calculation Action",
        "qc_status_valid": "✅ Valid",
        "qc_status_incomp": "❌ Incomplete",
        "qc_status_perfect": "🟢 100% Complete",
        "qc_act_calc": "Calculated (Exclude NA)",
        "qc_act_reject": "Rejected (N.A Incomplete)",
        "info_upload": "👈 Please upload raw AAWS files in the sidebar to begin processing."
    }
}

# ---------------------------------------------------------
# 3. SIDEBAR: KAWALAN & DATA
# ---------------------------------------------------------
st.sidebar.header("⚙️ Tetapan / Settings")
selected_lang = st.sidebar.selectbox("🌐 Bahasa / Language", options=["Bahasa Melayu", "English"])
lang_key = "BM" if selected_lang == "Bahasa Melayu" else "EN"
t = TEXTS[lang_key]

st.sidebar.divider()
qc_rule = st.sidebar.radio(
    t["qc_mode_label"],
    options=["WMO Standard (11/5 Rule)", "Strict Rule (5/3 Rule)", "No Filter (Raw Data)"],
    index=0
)

st.sidebar.header(t["sidebar_header"])
uploaded_files = st.sidebar.file_uploader(
    t["upload_label"], 
    type=["xls", "xlsx"],
    accept_multiple_files=True
)

# ---------------------------------------------------------
# 4. HEADER RASMI METMALAYSIA (GOVERNMENT PORTAL STYLE)
# ---------------------------------------------------------
st.markdown("""
<style>
    /* Pengepala Utama Putih Rasmi */
    .gov-header-container {
        display: flex;
        align-items: center;
        padding: 10px 0 15px 0;
        border-bottom: 2px solid #e2e8f0;
    }
    .gov-tag {
        font-size: 11px;
        letter-spacing: 1.5px;
        font-weight: 700;
        color: #4a5568;
        margin-bottom: 2px;
    }
    .gov-main-title {
        font-size: 26px;
        font-weight: 800;
        color: #1a2a4b;
        letter-spacing: 0.5px;
        margin: 0;
        line-height: 1.2;
    }
    .gov-sub-title {
        font-size: 13px;
        color: #596780;
        margin-top: 4px;
        font-weight: 500;
    }
    
    /* Kad Pengumuman / Side Box Hijau Zaitun */
    .announcement-box {
        background: linear-gradient(135deg, #8ba858 0%, #6e8b3d 100%);
        border-radius: 8px;
        padding: 16px 20px;
        color: #ffffff;
        box-shadow: 0 4px 10px rgba(0,0,0,0.06);
    }
    .announcement-title {
        font-size: 16px;
        font-weight: 700;
        margin-bottom: 6px;
    }
    .announcement-text {
        font-size: 13px;
        line-height: 1.4;
        color: #f7fafc;
    }
    
    /* Kad Utama Putih */
    .main-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 24px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.03);
    }
</style>
""", unsafe_allow_html=True)

# Paparan Pengepala
col_logo, col_heading = st.columns([1.2, 7])
with col_logo:
    try:
        st.image("logo_met.png", width=165)
    except:
        st.write("🌤️")

with col_heading:
    st.markdown(f"""
    <div class="gov-tag">{t['portal_tag']}</div>
    <div class="gov-main-title">{t['agency_title']}</div>
    <div class="gov-sub-title">🏛️ {t['branch_title']}</div>
    """, unsafe_allow_html=True)

# ---------------------------------------------------------
# 5. NAVIGASI BAR HORIZONTAL BIRU GELAP (NAVY TABS)
# ---------------------------------------------------------
tab_home, tab_rain, tab_temp, tab_qc = st.tabs([
    f"🔹 {t['nav_home']}", 
    f"🌧️ {t['nav_rain']}", 
    f"🌡️ {t['nav_temp']}", 
    f"📋 {t['nav_qc']}"
])

# ---------------------------------------------------------
# 6. FUNGSI PEMPROSESAN DATA AAWS
# ---------------------------------------------------------
def process_multiple_aaws_files(files_list):
    all_stations_data = {}
    for file in files_list:
        try:
            xls = pd.ExcelFile(file)
            for sheet in xls.sheet_names:
                if sheet.lower() == 'datalist':
                    continue
                df = pd.read_excel(xls, sheet_name=sheet)
                raw_station_text = str(df.iloc[2, 0]) if df.shape[0] > 2 else "Station"
                station_name = raw_station_text.split(':', 1)[1].strip() if ':' in raw_station_text else raw_station_text.replace("Station", "").strip()
                if not station_name or station_name == "nan":
                    station_name = f"Station_{sheet}"
                    
                data = df.iloc[11:].copy().iloc[:, :4]
                data.columns = ['Year', 'Month', 'Day', 'Rainfall']
                data['Year'] = pd.to_numeric(data['Year'], errors='coerce')
                data['Month'] = pd.to_numeric(data['Month'], errors='coerce')
                data['Day'] = pd.to_numeric(data['Day'], errors='coerce')
                data['Rainfall_Numeric'] = pd.to_numeric(data['Rainfall'], errors='coerce')
                data['Rainfall_Display'] = data['Rainfall']
                
                data = data.dropna(subset=['Year', 'Month', 'Day'])
                data = data[(data['Year'] > 1900) & (data['Year'] < 2100)]
                data['Year'] = data['Year'].astype(int)
                data['Month'] = data['Month'].astype(int)
                data['Day'] = data['Day'].astype(int)
                
                if not data.empty:
                    if station_name in all_stations_data:
                        combined_df = pd.concat([all_stations_data[station_name], data], ignore_index=True)
                        combined_df = combined_df.drop_duplicates(subset=['Year', 'Month', 'Day'])
                        all_stations_data[station_name] = combined_df
                    else:
                        all_stations_data[station_name] = data
        except Exception as e:
            st.error(f"Error: {file.name} - {e}")
    return all_stations_data

def evaluate_month_qc(series, rule):
    missing_count = series.isna().sum()
    is_na = series.isna().astype(int)
    blocks = (is_na != is_na.shift()).cumsum()
    consecutive_na = is_na.groupby(blocks).transform('sum') * is_na
    max_consecutive = consecutive_na.max() if not consecutive_na.empty else 0
    
    is_valid = True
    if rule == "WMO Standard (11/5 Rule)":
        if missing_count >= 11 or max_consecutive >= 5:
            is_valid = False
    elif rule == "Strict Rule (5/3 Rule)":
        if missing_count > 5 or max_consecutive > 3:
            is_valid = False
    return is_valid, missing_count, max_consecutive

def generate_qc_audit_table(df_station, rule, t_dict):
    qc_records = []
    month_names = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
    for yr in sorted(df_station['Year'].unique()):
        df_yr = df_station[df_station['Year'] == yr]
        pivot_num = df_yr.pivot(index='Day', columns='Month', values='Rainfall_Numeric')
        pivot_num = pivot_num.reindex(index=range(1, 32), columns=range(1, 13))
        for m in range(1, 13):
            col_data = pivot_num[m]
            is_valid, miss_cnt, max_consec = evaluate_month_qc(col_data, rule)
            status_str = t_dict["qc_status_perfect"] if miss_cnt == 0 else (t_dict["qc_status_valid"] if is_valid else t_dict["qc_status_incomp"])
            action_str = t_dict["qc_act_calc"] if is_valid else t_dict["qc_act_reject"]
            qc_records.append({
                t_dict["qc_col_year"]: yr,
                t_dict["qc_col_month"]: month_names[m-1],
                t_dict["qc_col_na"]: miss_cnt,
                t_dict["qc_col_consec"]: max_consec,
                t_dict["qc_col_status"]: status_str,
                t_dict["qc_col_action"]: action_str
            })
    return pd.DataFrame(qc_records)

def generate_excel_for_station(station_name, df_station, rule):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        years = sorted(df_station['Year'].unique())
        month_names = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
        has_written = False
        for yr in years:
            df_yr = df_station[df_station['Year'] == yr]
            if df_yr.empty:
                continue
            pivot_num = df_yr.pivot(index='Day', columns='Month', values='Rainfall_Numeric').reindex(index=range(1, 32), columns=range(1, 13))
            pivot_display = df_yr.pivot(index='Day', columns='Month', values='Rainfall_Display').reindex(index=range(1, 32), columns=range(1, 13))
            total_rain, rain_days, highest_fall, highest_date = [], [], [], []
            for m in range(1, 13):
                col_data = pivot_num[m]
                is_valid, _, _ = evaluate_month_qc(col_data, rule)
                if is_valid:
                    tot = col_data.sum(skipna=True)
                    total_rain.append(round(tot, 1) if pd.notna(tot) else "N.A")
                    rain_days.append((col_data > 0.1).sum())
                    highest_fall.append(round(col_data.max(skipna=True), 1) if pd.notna(col_data.max(skipna=True)) else "N.A")
                    highest_date.append(int(col_data.idxmax(skipna=True)) if pd.notna(col_data.idxmax(skipna=True)) else "-")
                else:
                    total_rain.append("N.A (Incomplete)")
                    rain_days.append("N.A")
                    highest_fall.append("N.A")
                    highest_date.append("-")
            report_df = pivot_display.copy()
            report_df.loc['TOTAL'] = total_rain
            report_df.loc['No. Of Days (>0.1mm)'] = rain_days
            report_df.loc['Highest Fall'] = highest_fall
            report_df.loc['Date of Highest'] = highest_date
            report_df.columns = month_names
            report_df.index.name = "DATE"
            report_df.to_excel(writer, sheet_name=str(yr)[:31])
            has_written = True
        if not has_written:
            pd.DataFrame({"Note": ["No Data"]}).to_excel(writer, sheet_name="No Data")
    output.seek(0)
    return output

stations_data = {}
if uploaded_files:
    stations_data = process_multiple_aaws_files(uploaded_files)

# ---------------------------------------------------------
# 7. KANDUNGAN SETIAP TAB UTAMA
# ---------------------------------------------------------

# === TAB 1: UTAMA ===
with tab_home:
    col_main, col_side = st.columns([2.2, 1])
    with col_main:
        st.markdown(f"""
        <div class="main-card">
            <h3>{t['home_card_title']}</h3>
            <p>{t['home_card_desc']}</p>
        </div>
        """, unsafe_allow_html=True)
        st.markdown(t["home_features"])
        if not uploaded_files:
            st.info(t["info_upload"])
            
    with col_side:
        st.markdown(f"""
        <div class="announcement-box">
            <div class="announcement-title">{t['wmo_announcement_title']}</div>
            <div class="announcement-text">{t['wmo_announcement_desc']}</div>
        </div>
        """, unsafe_allow_html=True)
        st.write("")
        try:
            with open("WMO Guidelines on the Calculation of Climate Normals_en.pdf", "rb") as f:
                pdf_bytes = f.read()
            st.download_button(
                label=t["btn_download_wmo"],
                data=pdf_bytes,
                file_name="WMO_Guidelines_Calculation_Climate_Normals_No1203.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        except:
            st.caption("ℹ️ Fail rujukan WMO PDF tersedia di sistem.")

# === TAB 2: HUJAN ===
with tab_rain:
    if uploaded_files and stations_data:
        st.success(t["upload_success"].format(count=len(uploaded_files)))
        col_zip1, col_zip2 = st.columns([3, 1])
        with col_zip1:
            st.subheader(t["stations_ready"].format(count=len(stations_data)))
        with col_zip2:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for st_name, st_df in stations_data.items():
                    excel_bytes = generate_excel_for_station(st_name, st_df, qc_rule)
                    clean_st_name = "".join([c for c in st_name if c.isalnum() or c in (' ', '_', '-')]).strip()
                    zip_file.writestr(f"Climatology_{clean_st_name.replace(' ', '_')}.xlsx", excel_bytes.getvalue())
            zip_buffer.seek(0)
            st.download_button(
                label=t["download_zip"],
                data=zip_buffer,
                file_name=t["zip_filename"],
                mime="application/zip",
                type="primary"
            )
        st.divider()
        
        selected_stesen = st.selectbox(t["select_station"], options=list(stations_data.keys()))
        df_stesen = stations_data[selected_stesen]
        min_yr, max_yr = df_stesen['Year'].min(), df_stesen['Year'].max()
        
        qc_df_station = generate_qc_audit_table(df_stesen, qc_rule, t)
        total_months = len(qc_df_station)
        incomplete_months_count = (qc_df_station[t["qc_col_status"]] == t["qc_status_incomp"]).sum()
        total_missing_days = df_stesen['Rainfall_Numeric'].isna().sum()
        completeness_pct = ((len(df_stesen) - total_missing_days) / len(df_stesen)) * 100
        
        m1, m2, m3, m4 = st.columns(4)
        m1.metric(t["station_name"], selected_stesen)
        m2.metric(t["record_period"], f"{min_yr} - {max_yr}")
        m3.metric(t["completeness_rate"], f"{completeness_pct:.1f}%")
        m4.metric(t["invalid_months"], f"{incomplete_months_count} / {total_months}", delta=f"-{incomplete_months_count}" if incomplete_months_count > 0 else None, delta_color="inverse")
        
        if incomplete_months_count > 0:
            st.warning(t["alert_incomplete"].format(count=incomplete_months_count, rule=qc_rule))
            
        sub_form, sub_plots = st.tabs([t["subtab_form"], t["subtab_charts"]])
        with sub_form:
            excel_file = generate_excel_for_station(selected_stesen, df_stesen, qc_rule)
            clean_st_name = "".join([c for c in selected_stesen if c.isalnum() or c in (' ', '_', '-')]).strip()
            st.download_button(
                label=t["download_excel"].format(station=selected_stesen),
                data=excel_file,
                file_name=f"Borang_Klimatologi_{clean_st_name.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.write(t["preview_title"])
            st.dataframe(df_stesen.head(20), use_container_width=True)
            
        with sub_plots:
            ctrl_col1, ctrl_col2 = st.columns(2)
            chart_options_map = {
                t["opt_heatmap"]: "Heatmap",
                t["opt_trend"]: "Trend",
                t["opt_normals"]: "Normals"
            }
            with ctrl_col1:
                chart_choice_label = st.selectbox(t["plot_type"], options=list(chart_options_map.keys()))
                chart_choice = chart_options_map[chart_choice_label]
            with ctrl_col2:
                color_choice = st.selectbox(t["color_theme"], options=["Blues", "Viridis", "YlGnBu", "Spectral", "Plasma", "Turbo", "Teal"])
                
            if chart_choice == "Heatmap":
                years_list = sorted(df_stesen['Year'].unique())
                chosen_year = st.selectbox(t["select_year_heat"], options=years_list, index=len(years_list)-1)
                df_heat = df_stesen[df_stesen['Year'] == chosen_year]
                heat_pivot = df_heat.pivot(index='Day', columns='Month', values='Rainfall_Numeric').reindex(index=range(1, 32), columns=range(1, 13))
                month_labels = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
                fig_heat = px.imshow(
                    heat_pivot,
                    labels=dict(x=t["axis_month"], y=t["axis_day"], color=t["axis_rain"]),
                    x=month_labels,
                    y=[str(d) for d in range(1, 32)],
                    color_continuous_scale=color_choice,
                    aspect="auto",
                    title=t["heat_title"].format(station=selected_stesen, year=chosen_year)
                )
                fig_heat.update_layout(height=650)
                st.plotly_chart(fig_heat, use_container_width=True)
                
            elif chart_choice == "Trend":
                annual_df = df_stesen.groupby('Year')['Rainfall_Numeric'].sum().reset_index()
                mean_val = annual_df['Rainfall_Numeric'].mean()
                fig_trend = px.bar(
                    annual_df, x='Year', y='Rainfall_Numeric', color='Rainfall_Numeric',
                    color_continuous_scale=color_choice,
                    labels={'Rainfall_Numeric': t["axis_rain"], 'Year': t["axis_year"]},
                    title=t["trend_title"].format(station=selected_stesen, min_yr=min_yr, max_yr=max_yr)
                )
                fig_trend.add_hline(y=mean_val, line_dash="dash", line_color="red", annotation_text=f"{t['trend_avg_label']}: {mean_val:.1f} mm")
                st.plotly_chart(fig_trend, use_container_width=True)
                
            elif chart_choice == "Normals":
                month_names = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
                month_df = df_stesen.groupby('Month')['Rainfall_Numeric'].mean().reset_index()
                month_df['Month_Name'] = month_df['Month'].apply(lambda x: month_names[x-1])
                fig_norm = px.line(
                    month_df, x='Month_Name', y='Rainfall_Numeric', markers=True,
                    title=t["norm_title"].format(station=selected_stesen),
                    labels={'Rainfall_Numeric': t["axis_avg_rain"], 'Month_Name': t["axis_month"]}
                )
                fig_norm.update_traces(line_color="#1f77b4", marker=dict(size=9))
                st.plotly_chart(fig_norm, use_container_width=True)
    else:
        st.info(t["info_upload"])

# === TAB 3: SUHU ===
with tab_temp:
    st.subheader(t["temp_title"])
    st.info(t["temp_desc"])

# === TAB 4: QC ===
with tab_qc:
    st.subheader(t["qc_title"])
    if uploaded_files and stations_data:
        qc_station_choice = st.selectbox(t["qc_select_station"], options=list(stations_data.keys()), key="qc_select")
        qc_table = generate_qc_audit_table(stations_data[qc_station_choice], qc_rule, t)
        filter_col1, _ = st.columns([2, 1])
        with filter_col1:
            show_failed_only = st.checkbox(t["qc_filter_failed"], value=False)
        display_qc_table = qc_table[qc_table[t["qc_col_na"]] > 0] if show_failed_only else qc_table
        st.dataframe(display_qc_table, use_container_width=True)
        csv_buffer = display_qc_table.to_csv(index=False).encode('utf-8')
        st.download_button(
            label=t["download_qc_csv"],
            data=csv_buffer,
            file_name=f"Log_Audit_QC_{qc_station_choice}.csv",
            mime="text/csv"
        )
    else:
        st.info(t["info_upload"])
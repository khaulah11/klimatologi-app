import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
import plotly.express as px

# ---------------------------------------------------------
# 1. TETAPAN HALAMAN & TEMA
# ---------------------------------------------------------
st.set_page_config(
    page_title="MetMalaysia Climatology System",
    layout="wide",
    page_icon="🌤️"
)

# ---------------------------------------------------------
# 2. KAMUS BAHASA (BILINGUAL DICTIONARY)
# ---------------------------------------------------------
TEXTS = {
    "BM": {
        "page_title": "Sistem Automasi Klimatologi MetMalaysia",
        "title": "Sistem Automasi Analisis Klimatologi",
        "subtitle": "Jabatan Meteorologi Malaysia (MetMalaysia) | Pejabat Meteorologi Sabah",
        "desc": "Aplikasi ini memproses siri masa data AAWS kepada Format Borang Rekod Hujan Piawai, Audit Integriti WMO, dan Analisis Visual Interaktif.",
        
        # Navigation
        "nav_header": "🧭 Menu Navigasi",
        "menu_home": "🏠 Utama & Panduan",
        "menu_rain": "🌧️ Hujan (Rainfall)",
        "menu_temp": "🌡️ Suhu (Temperature)",
        "menu_qc": "📋 Semakan Kualiti & WMO",
        
        # Sidebar
        "sidebar_settings": "⚙️ Tetapan / Settings",
        "lang_select": "🌐 Bahasa / Language",
        "sidebar_header": "📁 Muat Naik Fail Data",
        "upload_label": "Muat naik fail raw AAWS (.xls / .xlsx)",
        "upload_success": "✅ Berjaya memuat naik {count} fail AAWS!",
        "qc_mode_label": "⚙️ Piawaian Penapisan Data Hilang (WMO):",
        
        # Home Page
        "home_title": "📌 Panduan Penggunaan & Standard WMO-No. 1203",
        "home_body": """
        Selamat datang ke **Sistem Automasi Analisis Klimatologi MetMalaysia**. Sistem ini dibangunkan khusus untuk memudahkan pengurusan data iklim:
        
        * 📋 **Borang Piawai MetMalaysia:** Menjana jadual matriks harian bulanan (JAN-DEC) secara automatik bagi setiap tahun.
        * 🛡️ **Penapisan Integriti WMO:** Mengesan data hilang (NA) mengikut ketetapan WMO-No. 1203 untuk mengelakkan ralat anggaran.
        * 📈 **Visualisasi Interaktif:** Menyediakan Heatmap keamatan hujan dan analisis siri masa dengan pilihan penukar tema warna.
        """,
        "home_wmo_box": "📄 **Dokumen Rujukan Rasmi WMO-No. 1203** sedia ada dalam sistem untuk panduan kiraan Climate Normals.",
        "btn_download_wmo": "📥 Muat Turun Garis Panduan WMO-No. 1203 (PDF)",
        "home_info": "👈 Sila muat naik fail raw AAWS di menu bar sisi kiri untuk memulakan pemprosesan.",
        
        # Rainfall Page
        "stations_ready": "📌 {count} Stesen Tersedia",
        "download_zip": "📦 MUAT TURUN SEMUA STESEN (.ZIP)",
        "zip_filename": "Laporan_Klimatologi_Semua_Stesen.zip",
        "select_station": "Pilih Stesen untuk Analisis & Laporan:",
        "station_name": "Nama Stesen",
        "record_period": "Sela Masa Rekod",
        "total_records": "Jumlah Rekod Data",
        "completeness_rate": "Tahap Kesempurnaan Data",
        "invalid_months": "Bulan Tidak Sah (Incomplete)",
        "alert_incomplete": "⚠️ **Perhatian Pegawai:** Terdapat **{count} bulan** yang gagal melepasi piawaian kesempurnaan data ({rule}). Nilai jumlah bagi bulan-bulan tersebut ditandakan secara automatik sebagai `N.A (Incomplete)` pada borang Excel untuk mengelakkan ralat anggaran.",
        "subtab_form": "📄 Borang Rekod & Muat Turun",
        "subtab_charts": "📈 Visualisasi & Analisis Interaktif",
        "download_excel": "📥 Muat Turun Borang Excel ({station})",
        "preview_title": "#### 🔍 Pratonton Data Terkini (20 Baris Terawal)",
        "plot_type": "📊 Pilih Jenis Graf Interaktif:",
        "color_theme": "🎨 Pilih Tema Warna Graf:",
        "opt_heatmap": "🌧️ Matriks Harian Hujan (Heatmap)",
        "opt_trend": "📈 Trend Jumlah Hujan Tahunan (Time-Series)",
        "opt_normals": "📊 Purata Taburan Hujan Bulanan (Climatological Normals)",
        "select_year_heat": "Pilih Tahun untuk Heatmap:",
        "heat_title": "Matriks Keamatan Hujan Harian bagi {station} ({year})",
        "axis_month": "Bulan",
        "axis_day": "Hari",
        "axis_rain": "Hujan (mm)",
        "axis_year": "Tahun",
        "axis_avg_rain": "Purata Hujan (mm)",
        "trend_title": "Trend Jumlah Hujan Tahunan bagi {station} ({min_yr} - {max_yr})",
        "trend_avg_label": "Purata",
        "norm_title": "Corak Purata Hujan Bulanan (Normal Profile) bagi {station}",
        
        # Temperature Page
        "temp_title": "🌡️ Modul Analisis Suhu Udara (Temperature)",
        "temp_desc": "Modul Suhu Udara sedang diselaraskan untuk menyokong format Maximum / Minimum Temperature mengikut garis panduan WMO.",
        
        # QC Page
        "qc_title": "📋 Laporan Audit Kualiti Data & Piawaian WMO-No. 1203",
        "qc_desc": """
        Garis panduan rasmi **WMO-No. 1203 (Calculation of Climate Normals)** menetapkan:
        * **Peraturan 11/5 Hari (Data Bulanan):** Data bulanan dianggap tidak sah sekiranya mengandungi $\ge 11$ hari hilang atau $\ge 5$ hari hilang berturut-turut.
        * **Kriteria 80% (Normal 30-Tahun):** Pengiraan *Climate Normals* 30-tahun rasmi memerlukan sekurang-kurangnya 80% data bulanan yang lengkap ($\ge 24$ tahun daripada tempoh 30 tahun).
        """,
        "qc_select_station": "Pilih Stesen untuk Semakan Log QC:",
        "qc_filter_failed": "🔍 Paparkan Bulan Tidak Sah / Ada Data Hilang Sahaja",
        "download_qc_csv": "📥 Muat Turun Laporan Log Audit QC (.CSV)",
        "qc_col_year": "Tahun",
        "qc_col_month": "Bulan",
        "qc_col_na": "Hari Hilang (NA)",
        "qc_col_consec": "Maks. Berturut-turut (Hari)",
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
        "page_title": "MetMalaysia Climatology Automation System",
        "title": "Climatology Analysis Automation System",
        "subtitle": "Malaysian Meteorological Department (MetMalaysia) | Sabah Meteorological Office",
        "desc": "This application processes AAWS time-series data into Standard Rainfall Record Sheets, WMO Integrity Audit, and Interactive Visual Analytics.",
        
        # Navigation
        "nav_header": "🧭 Navigation Menu",
        "menu_home": "🏠 Home & Guide",
        "menu_rain": "🌧️ Rainfall",
        "menu_temp": "🌡️ Temperature",
        "menu_qc": "📋 QC & WMO Completeness",
        
        # Sidebar
        "sidebar_settings": "⚙️ Settings",
        "lang_select": "🌐 Language",
        "sidebar_header": "📁 Upload Data Files",
        "upload_label": "Upload raw AAWS files (.xls / .xlsx)",
        "upload_success": "✅ Successfully uploaded {count} AAWS file(s)!",
        "qc_mode_label": "⚙️ Missing Data Standard (WMO):",
        
        # Home Page
        "home_title": "📌 System User Guide & WMO-No. 1203 Standard",
        "home_body": """
        Welcome to the **MetMalaysia Climatology Analysis Automation System**. This application is designed to automate climate data processing and verification:
        
        * 📋 **Standard MetMalaysia Sheets:** Automatically generates monthly daily record grids (JAN-DEC) for each year.
        * 🛡️ **WMO Integrity Screening:** Screens missing data (NA) compliant with WMO-No. 1203 standards to prevent calculation bias.
        * 📈 **Interactive Analytics:** Offers Rainfall Intensity Heatmaps and time-series trend analysis with customizable color palettes.
        """,
        "home_wmo_box": "📄 **Official Reference: WMO-No. 1203 Guidelines** available in the system for Climate Normals calculation.",
        "btn_download_wmo": "📥 Download WMO-No. 1203 Guidelines (PDF)",
        "home_info": "👈 Please upload raw AAWS files in the sidebar to begin analysis.",
        
        # Rainfall Page
        "stations_ready": "📌 {count} Station(s) Available",
        "download_zip": "📦 DOWNLOAD ALL STATIONS (.ZIP)",
        "zip_filename": "All_Stations_Climatology_Reports.zip",
        "select_station": "Select Station for Analysis & Reports:",
        "station_name": "Station Name",
        "record_period": "Record Period",
        "total_records": "Total Records",
        "completeness_rate": "Data Completeness Rate",
        "invalid_months": "Invalid Months (Incomplete)",
        "alert_incomplete": "⚠️ **Officer Advisory:** There are **{count} month(s)** that do not meet the data completeness standard ({rule}). Total values for these months are automatically flagged as `N.A (Incomplete)` on the Excel sheet to prevent underestimation bias.",
        "subtab_form": "📄 Record Sheets & Download",
        "subtab_charts": "📈 Interactive Analytics & Charts",
        "download_excel": "📥 Download Excel Report ({station})",
        "preview_title": "#### 🔍 Raw Data Preview (First 20 Records)",
        "plot_type": "📊 Select Interactive Chart Type:",
        "color_theme": "🎨 Select Chart Color Theme:",
        "opt_heatmap": "🌧️ Daily Rainfall Matrix (Heatmap)",
        "opt_trend": "📈 Annual Total Rainfall Trend (Time-Series)",
        "opt_normals": "📊 Monthly Average Rainfall Distribution (Climatological Normals)",
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
        
        # Temperature Page
        "temp_title": "🌡️ Air Temperature Analysis Module",
        "temp_desc": "The Temperature Module is being integrated to support Maximum / Minimum Temperature parameters compliant with WMO standards.",
        
        # QC Page
        "qc_title": "📋 Data Quality Audit Report & WMO-No. 1203 Standards",
        "qc_desc": """
        Official **WMO-No. 1203 Guidelines (Calculation of Climate Normals)** state:
        * **11/5-Day Rule (Monthly Data):** A monthly value is invalid if it contains $\ge 11$ missing days or $\ge 5$ consecutive missing days.
        * **80% Completeness Criterion (30-Year Normals):** Calculating official 30-year *Climate Normals* requires at least 80% complete monthly data ($\ge 24$ years out of a 30-year period).
        """,
        "qc_select_station": "Select Station for QC Audit Logs:",
        "qc_filter_failed": "🔍 Display Incomplete / Missing Data Months Only",
        "download_qc_csv": "📥 Download QC Audit Log Report (.CSV)",
        "qc_col_year": "Year",
        "qc_col_month": "Month",
        "qc_col_na": "Missing Days (NA)",
        "qc_col_consec": "Max Consecutive (Days)",
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
# 3. SIDEBAR: TETAPAN & NAVIGASI MENU UTAMA
# ---------------------------------------------------------
st.sidebar.header(TEXTS["BM"]["sidebar_settings"])
selected_lang = st.sidebar.selectbox("🌐 Bahasa / Language", options=["Bahasa Melayu", "English"])
lang_key = "BM" if selected_lang == "Bahasa Melayu" else "EN"
t = TEXTS[lang_key]

st.sidebar.divider()

# Menu Navigasi Utama di Sebelah Kiri
st.sidebar.header(t["nav_header"])
menu_choice = st.sidebar.radio(
    "Pilihan Halaman / Page Selection:",
    options=[t["menu_home"], t["menu_rain"], t["menu_temp"], t["menu_qc"]],
    label_visibility="collapsed"
)

st.sidebar.divider()

# Sidebar: Pilihan Piawaian WMO bagi Data Hilang
qc_rule = st.sidebar.radio(
    t["qc_mode_label"],
    options=["WMO Standard (11/5 Rule)", "Strict Rule (5/3 Rule)", "No Filter (Raw Data)"],
    index=0
)

# Sidebar: Muat Naik Fail
st.sidebar.header(t["sidebar_header"])
uploaded_files = st.sidebar.file_uploader(
    t["upload_label"], 
    type=["xls", "xlsx"],
    accept_multiple_files=True
)

# ---------------------------------------------------------
# 4. CUSTOM HEADER DENGAN BACKGROUND CORAK METMALAYSIA
# ---------------------------------------------------------
st.markdown("""
<style>
    .header-card {
        background: linear-gradient(135deg, #102a45 0%, #1e4b7a 100%);
        border-radius: 12px;
        padding: 24px 30px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
    }
    .header-title {
        font-size: 28px;
        font-weight: 700;
        margin: 0;
        color: #ffffff;
    }
    .header-subtitle {
        font-size: 15px;
        color: #d0e1f9;
        margin-top: 4px;
        margin-bottom: 10px;
    }
    .header-desc {
        font-size: 14px;
        color: #f0f4f8;
        margin: 0;
        line-height: 1.4;
    }
</style>
""", unsafe_allow_html=True)

head_col1, head_col2 = st.columns([1.3, 6])
with head_col1:
    try:
        st.image("logo_met.png", width=175)
    except:
        st.write("🌤️")

with head_col2:
    st.markdown(f"""
    <div class="header-card">
        <div class="header-title">{t['title']}</div>
        <div class="header-subtitle">🏛️ {t['subtitle']}</div>
        <div class="header-desc">{t['desc']}</div>
    </div>
    """, unsafe_allow_html=True)

# ---------------------------------------------------------
# 5. FUNGSI PEMPROSESAN DATA AAWS
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
                if ':' in raw_station_text:
                    station_name = raw_station_text.split(':', 1)[1].strip()
                else:
                    station_name = raw_station_text.replace("Station", "").strip()
                    
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
            st.error(f"Error processing file {file.name}: {e}")
                
    return all_stations_data

# ---------------------------------------------------------
# 6. FUNGSI SEMAKAN DATA LENGKAP MENGIKUT WMO
# ---------------------------------------------------------
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

# ---------------------------------------------------------
# 7. FUNGSI JANA AUDIT QC DWIBALAS
# ---------------------------------------------------------
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
            
            if miss_cnt == 0:
                status_str = t_dict["qc_status_perfect"]
            elif is_valid:
                status_str = t_dict["qc_status_valid"]
            else:
                status_str = t_dict["qc_status_incomp"]
                
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

# ---------------------------------------------------------
# 8. FUNGSI PENJANAAN BORANG EXCEL
# ---------------------------------------------------------
def generate_excel_for_station(station_name, df_station, rule):
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        years = sorted(df_station['Year'].unique())
        month_names = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
        
        has_written_sheet = False
        
        for yr in years:
            df_yr = df_station[df_station['Year'] == yr]
            if df_yr.empty:
                continue
            
            pivot_num = df_yr.pivot(index='Day', columns='Month', values='Rainfall_Numeric')
            pivot_display = df_yr.pivot(index='Day', columns='Month', values='Rainfall_Display')
            
            pivot_num = pivot_num.reindex(index=range(1, 32), columns=range(1, 13))
            pivot_display = pivot_display.reindex(index=range(1, 32), columns=range(1, 13))
            
            total_rain = []
            rain_days = []
            highest_fall = []
            highest_date = []
            
            for m in range(1, 13):
                col_data = pivot_num[m]
                is_valid, _, _ = evaluate_month_qc(col_data, rule)
                
                if is_valid:
                    tot = col_data.sum(skipna=True)
                    r_days = (col_data > 0.1).sum()
                    h_fall = col_data.max(skipna=True)
                    h_date = col_data.idxmax(skipna=True)
                    
                    total_rain.append(round(tot, 1) if pd.notna(tot) else "N.A")
                    rain_days.append(r_days)
                    highest_fall.append(round(h_fall, 1) if pd.notna(h_fall) else "N.A")
                    highest_date.append(int(h_date) if pd.notna(h_date) else "-")
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
            
            sheet_title = str(yr)[:31]
            report_df.to_excel(writer, sheet_name=sheet_title)
            has_written_sheet = True
            
        if not has_written_sheet:
            pd.DataFrame({"Note": ["No Data"]}).to_excel(writer, sheet_name="No Data")
            
    output.seek(0)
    return output

# ---------------------------------------------------------
# 9. PEMPROSESAN DATA GLOBAL
# ---------------------------------------------------------
stations_data = {}
if uploaded_files:
    stations_data = process_multiple_aaws_files(uploaded_files)

# ---------------------------------------------------------
# 10. KAWALAN PAPARAN MENGIKUT PILIHAN MENU SIDEBAR
# ---------------------------------------------------------

# === HALAMAN 1: UTAMA & PANDUAN ===
if menu_choice == t["menu_home"]:
    st.subheader(t["home_title"])
    st.markdown(t["home_body"])
    
    st.info(t["home_wmo_box"])
    
    # Butang Muat Turun PDF WMO
    try:
        with open("WMO Guidelines on the Calculation of Climate Normals_en.pdf", "rb") as f:
            pdf_bytes = f.read()
        st.download_button(
            label=t["btn_download_wmo"],
            data=pdf_bytes,
            file_name="WMO_Guidelines_Calculation_Climate_Normals_No1203.pdf",
            mime="application/pdf"
        )
    except:
        st.caption("ℹ️ Letakkan fail `WMO Guidelines on the Calculation of Climate Normals_en.pdf` di folder projek untuk membolehkan muat turun PDF.")
        
    if not uploaded_files:
        st.divider()
        st.info(t["home_info"])

# === HALAMAN 2: HUJAN (RAINFALL) ===
elif menu_choice == t["menu_rain"]:
    if uploaded_files and stations_data:
        st.success(t["upload_success"].format(count=len(uploaded_files)))
        
        # Pilihan Muat Turun ZIP Kelompok
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
        
        # Pemilihan Stesen Individu
        selected_stesen = st.selectbox(t["select_station"], options=list(stations_data.keys()))
        df_stesen = stations_data[selected_stesen]
        
        min_yr = df_stesen['Year'].min()
        max_yr = df_stesen['Year'].max()
        
        # Pengiraan Ringkasan Integriti & QC
        qc_df_station = generate_qc_audit_table(df_stesen, qc_rule, t)
        total_months = len(qc_df_station)
        incomplete_months_count = (qc_df_station[t["qc_col_status"]] == t["qc_status_incomp"]).sum()
        total_missing_days = df_stesen['Rainfall_Numeric'].isna().sum()
        completeness_pct = ((len(df_stesen) - total_missing_days) / len(df_stesen)) * 100
        
        # 📊 KAD METRIK RINGKASAN INTEGRITI DATA
        m1, m2, m3, m4 = st.columns(4)
        m1.metric(t["station_name"], selected_stesen)
        m2.metric(t["record_period"], f"{min_yr} - {max_yr}")
        m3.metric(t["completeness_rate"], f"{completeness_pct:.1f}%")
        m4.metric(t["invalid_months"], f"{incomplete_months_count} / {total_months}", delta=f"-{incomplete_months_count}" if incomplete_months_count > 0 else None, delta_color="inverse")
        
        # Kotak Amaran Pintar
        if incomplete_months_count > 0:
            st.warning(t["alert_incomplete"].format(count=incomplete_months_count, rule=qc_rule))
        
        # Sub-Tab: Borang vs Graf
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
                chart_choice_label = st.selectbox(
                    t["plot_type"],
                    options=list(chart_options_map.keys())
                )
                chart_choice = chart_options_map[chart_choice_label]
                
            with ctrl_col2:
                color_choice = st.selectbox(
                    t["color_theme"],
                    options=["Blues", "Viridis", "YlGnBu", "Spectral", "Plasma", "Turbo", "Teal"]
                )
            
            # Heatmap
            if chart_choice == "Heatmap":
                years_list = sorted(df_stesen['Year'].unique())
                chosen_year = st.selectbox(t["select_year_heat"], options=years_list, index=len(years_list)-1)
                
                df_heat = df_stesen[df_stesen['Year'] == chosen_year]
                heat_pivot = df_heat.pivot(index='Day', columns='Month', values='Rainfall_Numeric')
                heat_pivot = heat_pivot.reindex(index=range(1, 32), columns=range(1, 13))
                
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
                
            # Trend
            elif chart_choice == "Trend":
                annual_df = df_stesen.groupby('Year')['Rainfall_Numeric'].sum().reset_index()
                mean_val = annual_df['Rainfall_Numeric'].mean()
                
                fig_trend = px.bar(
                    annual_df, 
                    x='Year', 
                    y='Rainfall_Numeric',
                    color='Rainfall_Numeric',
                    color_continuous_scale=color_choice,
                    labels={'Rainfall_Numeric': t["axis_rain"], 'Year': t["axis_year"]},
                    title=t["trend_title"].format(station=selected_stesen, min_yr=min_yr, max_yr=max_yr)
                )
                fig_trend.add_hline(
                    y=mean_val, 
                    line_dash="dash", 
                    line_color="red", 
                    annotation_text=f"{t['trend_avg_label']}: {mean_val:.1f} mm"
                )
                st.plotly_chart(fig_trend, use_container_width=True)
                
            # Normals
            elif chart_choice == "Normals":
                month_names = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
                month_df = df_stesen.groupby('Month')['Rainfall_Numeric'].mean().reset_index()
                month_df['Month_Name'] = month_df['Month'].apply(lambda x: month_names[x-1])
                
                fig_norm = px.line(
                    month_df,
                    x='Month_Name',
                    y='Rainfall_Numeric',
                    markers=True,
                    title=t["norm_title"].format(station=selected_stesen),
                    labels={'Rainfall_Numeric': t["axis_avg_rain"], 'Month_Name': t["axis_month"]}
                )
                fig_norm.update_traces(line_color="#1f77b4", marker=dict(size=9))
                st.plotly_chart(fig_norm, use_container_width=True)
    else:
        st.info(t["info_upload"])

# === HALAMAN 3: SUHU (TEMPERATURE) ===
elif menu_choice == t["menu_temp"]:
    st.subheader(t["temp_title"])
    st.info(t["temp_desc"])

# === HALAMAN 4: QC & WMO COMPLETENESS ===
elif menu_choice == t["menu_qc"]:
    st.subheader(t["qc_title"])
    st.markdown(t["qc_desc"])
    
    if uploaded_files and stations_data:
        st.divider()
        qc_station_choice = st.selectbox(t["qc_select_station"], options=list(stations_data.keys()), key="qc_select")
        qc_table = generate_qc_audit_table(stations_data[qc_station_choice], qc_rule, t)
        
        filter_col1, filter_col2 = st.columns([2, 1])
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
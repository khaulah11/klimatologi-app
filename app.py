import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile

# ---------------------------------------------------------
# 1. TETAPAN KAMUS BAHASA (LANGUAGE DICTIONARY)
# ---------------------------------------------------------
TEXTS = {
    "BM": {
        "page_title": "Sistem Automasi Klimatologi MetMalaysia",
        "title": "Sistem Automasi Analisis Klimatologi",
        "subtitle": "Jabatan Meteorologi Malaysia (MetMalaysia) | Pejabat Meteorologi Sabah",
        "desc": "Aplikasi ini menukar *raw data* siri masa **AAWS (Pemprosesan Berkelompok / Batch)** kepada **Format Borang Rekod Hujan Harian Piawai** secara automatik.",
        "sidebar_header": "📁 Muat Naik Fail Data",
        "upload_label": "Muat naik satu atau BERBILANG Fail Raw Data AAWS (.xls / .xlsx)",
        "upload_success": "✅ Berjaya memuat naik {count} fail AAWS!",
        "processing": "Sedang menggabungkan dan memproses semua siri masa stesen...",
        "stations_detected": "📌 {count} Stesen Dikesan Dari Semua Fail",
        "download_zip": "📦 MUAT TURUN SEMUA STESEN (FAIL .ZIP)",
        "zip_filename": "Laporan_Klimatologi_Semua_Stesen.zip",
        "select_station": "Atau pilih Stesen secara individu untuk Pratonton:",
        "station_name": "Nama Stesen",
        "record_period": "Sela Masa Rekod",
        "download_excel": "📥 Muat Turun Laporan Excel bagi {station}",
        "preview": "🔍 Pratonton Data Raw (20 Rekod Pertama)",
        "no_data": "Tiada data stesen yang sah dikesan dalam fail yang dimuat naik.",
        "info_upload": "👈 Sila muat naik satu atau lebih fail raw AAWS di menu tepi untuk bermula."
    },
    "EN": {
        "page_title": "MetMalaysia Climatology Automation System",
        "title": "Climatology Analysis Automation System",
        "subtitle": "Malaysian Meteorological Department (MetMalaysia) | Sabah Meteorological Office",
        "desc": "This application automatically processes time-series **AAWS raw data (Batch Processing)** and converts it into the **Standard Daily Rainfall Record Sheet Format**.",
        "sidebar_header": "📁 Upload Data Files",
        "upload_label": "Upload one or MULTIPLE Raw AAWS Data Files (.xls / .xlsx)",
        "upload_success": "✅ Successfully uploaded {count} AAWS file(s)!",
        "processing": "Merging and processing time-series data for all stations...",
        "stations_detected": "📌 {count} Stations Detected Across All Files",
        "download_zip": "📦 DOWNLOAD ALL STATIONS (.ZIP FILE)",
        "zip_filename": "All_Stations_Climatology_Reports.zip",
        "select_station": "Or select an individual station for preview:",
        "station_name": "Station Name",
        "record_period": "Record Period",
        "download_excel": "📥 Download Excel Report for {station}",
        "preview": "🔍 Raw Data Preview (First 20 Records)",
        "no_data": "No valid station data detected in the uploaded files.",
        "info_upload": "👈 Please upload one or more raw AAWS files in the sidebar to get started."
    }
}

# ---------------------------------------------------------
# 2. TETAPAN HALAMAN STREAMLIT & SIDEBAR TOGGLE
# ---------------------------------------------------------
st.set_page_config(
    page_title="MetMalaysia Climatology System",
    layout="wide",
    page_icon="🌤️"
)

# Sidebar Language Selector
st.sidebar.header("⚙️ Tetapan / Settings")
selected_lang = st.sidebar.selectbox("🌐 Bahasa / Language", options=["Bahasa Melayu", "English"])

lang_key = "BM" if selected_lang == "Bahasa Melayu" else "EN"
t = TEXTS[lang_key]

# ---------------------------------------------------------
# 3. HEADER ATAS: LOGO METMALAYSIA & TAJUK RASMI
# ---------------------------------------------------------
col_logo, col_title = st.columns([1.5, 6])

with col_logo:
    # Panggil terus fail logo tempatan yang Moza simpan
    try:
        st.image("logo_met.png", width=170)
    except:
        st.write("🌤️") # Backup icon jika gambar tiada

with col_title:
    st.title(t["title"])
    st.caption(f"🏛️ **{t['subtitle']}**")

st.markdown(t["desc"])
st.divider()

# Sidebar File Uploader
st.sidebar.header(t["sidebar_header"])
uploaded_files = st.sidebar.file_uploader(
    t["upload_label"], 
    type=["xls", "xlsx"],
    accept_multiple_files=True
)

# ---------------------------------------------------------
# 4. FUNGSI UTAMA: PEMPROSESAN BATCH AAWS
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
# 5. FUNGSI PENJANAAN BORANG EXCEL
# ---------------------------------------------------------
def generate_excel_for_station(station_name, df_station):
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
            
            total_rain = pivot_num.sum(axis=0, skipna=True)
            rain_days = (pivot_num > 0.1).sum(axis=0)
            highest_fall = pivot_num.max(axis=0, skipna=True)
            highest_date = pivot_num.idxmax(axis=0, skipna=True)
            
            report_df = pivot_display.copy()
            
            report_df.loc['TOTAL'] = total_rain.round(1)
            report_df.loc['No. Of Days (>0.1mm)'] = rain_days
            report_df.loc['Highest Fall'] = highest_fall.round(1)
            report_df.loc['Date of Highest'] = highest_date.fillna('-')
            
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
# 6. ALUR KERJA APLIKASI
# ---------------------------------------------------------
if uploaded_files:
    st.success(t["upload_success"].format(count=len(uploaded_files)))
    
    with st.spinner(t["processing"]):
        stations_data = process_multiple_aaws_files(uploaded_files)
        
    st.subheader(t["stations_detected"].format(count=len(stations_data)))
    
    if stations_data:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for st_name, st_df in stations_data.items():
                excel_bytes = generate_excel_for_station(st_name, st_df)
                clean_st_name = "".join([c for c in st_name if c.isalnum() or c in (' ', '_', '-')]).strip()
                clean_filename = f"Climatology_Report_{clean_st_name.replace(' ', '_')}.xlsx"
                zip_file.writestr(clean_filename, excel_bytes.getvalue())
                
        zip_buffer.seek(0)
        
        st.download_button(
            label=t["download_zip"],
            data=zip_buffer,
            file_name=t["zip_filename"],
            mime="application/zip",
            type="primary"
        )
        
        st.divider()
        
        selected_stesen = st.selectbox(
            t["select_station"], 
            options=list(stations_data.keys())
        )
        
        if selected_stesen:
            df_stesen = stations_data[selected_stesen]
            min_yr = df_stesen['Year'].min()
            max_yr = df_stesen['Year'].max()
            
            col1, col2 = st.columns(2)
            col1.metric(t["station_name"], selected_stesen)
            col2.metric(t["record_period"], f"{min_yr} - {max_yr} ({max_yr - min_yr + 1} Years)")
            
            excel_file = generate_excel_for_station(selected_stesen, df_stesen)
            clean_st_name = "".join([c for c in selected_stesen if c.isalnum() or c in (' ', '_', '-')]).strip()
            
            st.download_button(
                label=t["download_excel"].format(station=selected_stesen),
                data=excel_file,
                file_name=f"Climatology_Report_{clean_st_name.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.write(f"### {t['preview']}")
            st.dataframe(df_stesen.head(20), use_container_width=True)
    else:
        st.warning(t["no_data"])

else:
    st.info(t["info_upload"])
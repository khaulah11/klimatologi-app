import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile

# ---------------------------------------------------------
# 1. TETAPAN HALAMAN STREAMLIT
# ---------------------------------------------------------
st.set_page_config(
    page_title="Sistem Automasi Klimatologi MetMalaysia",
    layout="wide",
    page_icon="🌧️"
)

st.title("🌧️ Sistem Automasi Analisis Klimatologi MetMalaysia")
st.markdown("""
Aplikasi ini menukar *raw data* siri masa **AAWS (Pemprosesan Berkelompok / Batch)** kepada **Format Borang Rekod Hujan Harian Piawai** secara automatik.
""")

# ---------------------------------------------------------
# 2. MENU DITINGKAP TEPI (SIDEBAR) - MUAT NAIK BANYAK FAIL
# ---------------------------------------------------------
st.sidebar.header("📁 Muat Naik Fail Data")
uploaded_files = st.sidebar.file_uploader(
    "Muat naik satu atau BERBILANG Fail Raw Data AAWS (.xls / .xlsx)", 
    type=["xls", "xlsx"],
    accept_multiple_files=True  # Benarkan muat naik berkelompok
)

# ---------------------------------------------------------
# 3. FUNGSI UTAMA: PEMPROSESAN BATCH AAWS
# ---------------------------------------------------------
def process_multiple_aaws_files(files_list):
    """
    Membaca berbilang fail AAWS dan menggabungkan data mengikut Stesen.
    """
    all_stations_data = {}
    
    for file in files_list:
        xls = pd.ExcelFile(file)
        
        for sheet in xls.sheet_names:
            if sheet.lower() == 'datalist':
                continue
                
            df = pd.read_excel(xls, sheet_name=sheet)
            
            # Ekstrak Nama Stesen
            raw_station_text = str(df.iloc[2, 0])
            if ':' in raw_station_text:
                station_name = raw_station_text.split(':', 1)[1].strip()
            else:
                station_name = raw_station_text.replace("Station", "").strip()
                
            if not station_name or station_name == "nan":
                station_name = f"Stesen_{sheet}"
                
            data = df.iloc[11:].copy().iloc[:, :4]
            data.columns = ['Year', 'Month', 'Day', 'Rainfall']
            
            data['Year'] = pd.to_numeric(data['Year'], errors='coerce')
            data['Month'] = pd.to_numeric(data['Month'], errors='coerce')
            data['Day'] = pd.to_numeric(data['Day'], errors='coerce')
            
            data['Rainfall_Numeric'] = pd.to_numeric(data['Rainfall'], errors='coerce')
            data['Rainfall_Display'] = data['Rainfall']
            
            data = data.dropna(subset=['Year', 'Month', 'Day'])
            data['Year'] = data['Year'].astype(int)
            data['Month'] = data['Month'].astype(int)
            data['Day'] = data['Day'].astype(int)
            
            # Jika stesen wujud dalam fail berbeza, gabungkannya
            if station_name in all_stations_data:
                combined_df = pd.concat([all_stations_data[station_name], data], ignore_index=True)
                # Buang rekod bertindih jika ada
                combined_df = combined_df.drop_duplicates(subset=['Year', 'Month', 'Day'])
                all_stations_data[station_name] = combined_df
            else:
                all_stations_data[station_name] = data
                
    return all_stations_data

# ---------------------------------------------------------
# 4. FUNGSI PENJANAAN BORANG EXCEL
# ---------------------------------------------------------
def generate_excel_for_station(station_name, df_station):
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        years = sorted(df_station['Year'].unique())
        month_names = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
        
        for yr in years:
            df_yr = df_station[df_station['Year'] == yr]
            
            pivot_num = df_yr.pivot(index='Day', columns='Month', values='Rainfall_Numeric')
            pivot_display = df_yr.pivot(index='Day', columns='Month', values='Rainfall_Display')
            
            pivot_num = pivot_num.reindex(index=range(1, 32), columns=range(1, 13))
            pivot_display = pivot_display.reindex(index=range(1, 32), columns=range(1, 13))
            
            # Pengiraan Statistik
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
            
            report_df.to_excel(writer, sheet_name=str(yr))
            
    output.seek(0)
    return output

# ---------------------------------------------------------
# 5. ALUR KERJA APLIKASI BATCH
# ---------------------------------------------------------
if uploaded_files:
    st.success(f"✅ Berjaya memuat naik {len(uploaded_files)} fail AAWS!")
    
    with st.spinner("Sedang menggabungkan dan memproses semua siri masa stesen..."):
        stations_data = process_multiple_aaws_files(uploaded_files)
        
    st.subheader(f"📌 {len(stations_data)} Stesen Dikesan Dari Semua Fail")
    
    # Pilihan 1: Muat Turun Semua Stesen Sekaligus (ZIP File)
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for st_name, st_df in stations_data.items():
            excel_bytes = generate_excel_for_station(st_name, st_df)
            clean_filename = f"Borang_Klimatologi_{st_name.replace(' ', '_')}.xlsx"
            zip_file.writestr(clean_filename, excel_bytes.getvalue())
            
    zip_buffer.seek(0)
    
    st.download_button(
        label="📦 MUAT TURUN SEMUA STESEN (FAIL .ZIP)",
        data=zip_buffer,
        file_name="Laporan_Klimatologi_Semua_Stesen.zip",
        mime="application/zip",
        type="primary"
    )
    
    st.divider()
    
    # Pilihan 2: Pilih Stesen Tertentu Untuk Pratonton & Muat Turun Individu
    selected_stesen = st.selectbox(
        "Atau pilih Stesen secara individu untuk Pratonton:", 
        options=list(stations_data.keys())
    )
    
    if selected_stesen:
        df_stesen = stations_data[selected_stesen]
        min_yr = df_stesen['Year'].min()
        max_yr = df_stesen['Year'].max()
        
        col1, col2 = st.columns(2)
        col1.metric("Nama Stesen", selected_stesen)
        col2.metric("Sela Masa Rekod", f"{min_yr} - {max_yr} ({max_yr - min_yr + 1} Tahun)")
        
        excel_file = generate_excel_for_station(selected_stesen, df_stesen)
        
        st.download_button(
            label=f"📥 Muat Turun Laporan Excel bagi {selected_stesen}",
            data=excel_file,
            file_name=f"Borang_Klimatologi_{selected_stesen.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.write("### 🔍 Pratonton Data Raw (20 Rekod Pertama)")
        st.dataframe(df_stesen.head(20), use_container_width=True)

else:
    st.info("👈 Sila muat naik satu atau lebih fail raw AAWS di menu tepi untuk bermula.")
import streamlit as st
import pandas as pd
import numpy as np
import io

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
Aplikasi ini menukar *raw data* siri masa **AAWS** kepada **Format Borang Rekod Hujan Harian Piawai** secara automatik.
""")

# ---------------------------------------------------------
# 2. MENU DITINGKAP TEPI (SIDEBAR) - MUAT NAIK FAIL
# ---------------------------------------------------------
st.sidebar.header("📁 Muat Naik Fail Data")
uploaded_aaws = st.sidebar.file_uploader(
    "Muat naik mana-mana Fail Raw Data AAWS (.xls / .xlsx)", 
    type=["xls", "xlsx"]
)

# ---------------------------------------------------------
# 3. FUNGSI UTAMA: PEMPROSESAN DATA AAWS
# ---------------------------------------------------------
def process_aaws_file(file_aaws):
    xls = pd.ExcelFile(file_aaws)
    station_dict = {}
    
    for sheet in xls.sheet_names:
        if sheet.lower() == 'datalist':
            continue
            
        df = pd.read_excel(xls, sheet_name=sheet)
        
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
        
        data = data.dropna(subset=['Year', 'Month', 'Day'])
        data['Year'] = data['Year'].astype(int)
        data['Month'] = data['Month'].astype(int)
        data['Day'] = data['Day'].astype(int)
        
        station_dict[station_name] = data
        
    return station_dict

# ---------------------------------------------------------
# 4. FUNGSI PENJANAAN BORANG EXCEL (DENGAN STATISTIK)
# ---------------------------------------------------------
def generate_excel_for_station(station_name, df_station):
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        years = sorted(df_station['Year'].unique())
        month_names = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
        
        for yr in years:
            df_yr = df_station[df_station['Year'] == yr]
            
            pivot_num = df_yr.pivot(index='Day', columns='Month', values='Rainfall_Numeric')
            pivot_raw = df_yr.pivot(index='Day', columns='Month', values='Rainfall')
            
            pivot_num = pivot_num.reindex(index=range(1, 32), columns=range(1, 13))
            pivot_raw = pivot_raw.reindex(index=range(1, 32), columns=range(1, 13))
            
            # Statistik Klimatologi
            total_rain = pivot_num.sum(axis=0, skipna=True)
            rain_days = (pivot_num > 0.1).sum(axis=0)
            highest_fall = pivot_num.max(axis=0, skipna=True)
            highest_date = pivot_num.idxmax(axis=0, skipna=True)
            
            report_df = pivot_raw.copy()
            report_df = report_df.fillna('M')
            
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
# 5. ALUR KERJA APLIKASI
# ---------------------------------------------------------
if uploaded_aaws is not None:
    st.success("✅ Fail AAWS berjaya dimuat naik!")
    
    with st.spinner("Sedang memproses data siri masa..."):
        stations_data = process_aaws_file(uploaded_aaws)
        
    st.subheader("📌 Senarai Stesen Terkesan")
    selected_stesen = st.selectbox(
        "Pilih Stesen untuk Pratonton / Muat Turun:", 
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
        
        st.divider()
        st.write("### 🔍 Pratonton Data Raw (20 Rekod Pertama)")
        st.dataframe(df_stesen.head(20), use_container_width=True)

else:
    st.info("👈 Sila muat naik fail raw AAWS di bahagian menu tepi untuk bermula.")
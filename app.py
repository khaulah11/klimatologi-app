import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
import plotly.express as px

# ---------------------------------------------------------
# 1. KAMUS BAHASA (BILINGUAL DICTIONARY)
# ---------------------------------------------------------
TEXTS = {
    "BM": {
        "page_title": "Sistem Automasi Klimatologi MetMalaysia",
        "title": "Sistem Automasi Analisis Klimatologi",
        "subtitle": "Jabatan Meteorologi Malaysia (MetMalaysia) | Pejabat Meteorologi Sabah",
        "desc": "Aplikasi ini memproses siri masa data AAWS kepada Format Borang Rekod Hujan Piawai, Audit Integriti WMO, dan Analisis Visual Interaktif.",
        "tab_home": "🏠 Utama / Panduan",
        "tab_rain": "🌧️ Hujan (Rainfall)",
        "tab_temp": "🌡️ Suhu (Temperature)",
        "tab_qc": "📋 Semakan Kualiti & WMO",
        "sidebar_header": "📁 Muat Naik Fail Data",
        "upload_label": "Muat naik fail raw AAWS (.xls / .xlsx)",
        "upload_success": "✅ Berjaya memuat naik {count} fail AAWS!",
        "qc_mode_label": "⚙️ Piawaian Penapisan Data Hilang (WMO):",
        "download_zip": "📦 MUAT TURUN SEMUA STESEN (.ZIP)",
        "zip_filename": "Laporan_Klimatologi_Semua_Stesen.zip",
        "select_station": "Pilih Stesen untuk Analisis & Laporan:",
        "station_name": "Nama Stesen",
        "record_period": "Sela Masa Rekod",
        "download_excel": "📥 Muat Turun Borang Excel ({station})",
        "color_theme": "🎨 Pilih Tema Warna Graf:",
        "plot_type": "📊 Pilih Jenis Graf Interaktif:",
        "subtab_form": "📄 Borang & Muat Turun",
        "subtab_charts": "📈 Visualisasi & Analisis Interaktif",
        "info_upload": "👈 Sila muat naik fail raw AAWS di menu tepi untuk memulakan pemprosesan."
    },
    "EN": {
        "page_title": "MetMalaysia Climatology Automation System",
        "title": "Climatology Analysis Automation System",
        "subtitle": "Malaysian Meteorological Department (MetMalaysia) | Sabah Meteorological Office",
        "desc": "This application processes AAWS time-series data into Standard Rainfall Record Sheets, WMO Integrity Audit, and Interactive Visual Analytics.",
        "tab_home": "🏠 Home / Guide",
        "tab_rain": "🌧️ Rainfall",
        "tab_temp": "🌡️ Temperature",
        "tab_qc": "📋 QC & WMO Completeness",
        "sidebar_header": "📁 Upload Data Files",
        "upload_label": "Upload raw AAWS files (.xls / .xlsx)",
        "upload_success": "✅ Successfully uploaded {count} AAWS file(s)!",
        "qc_mode_label": "⚙️ Missing Data Standard (WMO):",
        "download_zip": "📦 DOWNLOAD ALL STATIONS (.ZIP)",
        "zip_filename": "All_Stations_Climatology_Reports.zip",
        "select_station": "Select Station for Analysis & Reports:",
        "station_name": "Station Name",
        "record_period": "Record Period",
        "download_excel": "📥 Download Excel Report ({station})",
        "color_theme": "🎨 Select Chart Color Theme:",
        "plot_type": "📊 Select Interactive Chart Type:",
        "subtab_form": "📄 Records & Download",
        "subtab_charts": "📈 Interactive Analytics & Charts",
        "info_upload": "👈 Please upload raw AAWS files in the sidebar to begin processing."
    }
}

# ---------------------------------------------------------
# 2. TETAPAN HALAMAN & TOGGLE BAHASA
# ---------------------------------------------------------
st.set_page_config(
    page_title="MetMalaysia Climatology System",
    layout="wide",
    page_icon="🌤️"
)

st.sidebar.header("⚙️ Tetapan / Settings")
selected_lang = st.sidebar.selectbox("🌐 Bahasa / Language", options=["Bahasa Melayu", "English"])
lang_key = "BM" if selected_lang == "Bahasa Melayu" else "EN"
t = TEXTS[lang_key]

# ---------------------------------------------------------
# 3. HEADER RASMI DENGAN LOGO METMALAYSIA
# ---------------------------------------------------------
col_logo, col_title = st.columns([1.5, 6])

with col_logo:
    try:
        st.image("logo_met.png", width=170)
    except:
        st.write("🌤️")

with col_title:
    st.title(t["title"])
    st.caption(f"🏛️ **{t['subtitle']}**")

st.markdown(t["desc"])
st.divider()

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
# 4. FUNGSI PEMPROSESAN DATA AAWS
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
# 5. FUNGSI SEMAKAN DATA LENGKAP & AUDIT LOG MENGIKUT WMO
# ---------------------------------------------------------
def evaluate_month_qc(series, rule):
    missing_count = series.isna().sum()
    
    # Kira hari hilang berturut-turut
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
# 6. FUNGSI JANA AUDIT QC UNTUK SESUATU STESEN
# ---------------------------------------------------------
def generate_qc_audit_table(df_station, rule):
    qc_records = []
    month_names = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
    
    for yr in sorted(df_station['Year'].unique()):
        df_yr = df_station[df_station['Year'] == yr]
        pivot_num = df_yr.pivot(index='Day', columns='Month', values='Rainfall_Numeric')
        pivot_num = pivot_num.reindex(index=range(1, 32), columns=range(1, 13))
        
        for m in range(1, 13):
            col_data = pivot_num[m]
            is_valid, miss_cnt, max_consec = evaluate_month_qc(col_data, rule)
            
            status_str = "✅ Valid" if is_valid else "❌ Incomplete"
            if miss_cnt == 0:
                status_str = "🟢 100% Complete"
                
            qc_records.append({
                "Tahun": yr,
                "Bulan": month_names[m-1],
                "Hari Hilang (NA)": miss_cnt,
                "Maks. Berturut-turut (Hari)": max_consec,
                "Status Integriti": status_str,
                "Tindakan Pengiraan": "Dikira (Abaikan NA)" if is_valid else "Ditolak (N.A Incomplete)"
            })
            
    return pd.DataFrame(qc_records)

# ---------------------------------------------------------
# 7. FUNGSI PENJANAAN BORANG EXCEL
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
# 8. STRUKTUR TAB UTAMA (HOME / RAINFALL / TEMP / QC)
# ---------------------------------------------------------
tab_home, tab_rain, tab_temp, tab_qc = st.tabs([
    t["tab_home"], 
    t["tab_rain"], 
    t["tab_temp"], 
    t["tab_qc"]
])

# === TAB 1: UTAMA / HOME ===
with tab_home:
    st.subheader("📌 Panduan Penggunaan Sistem")
    st.markdown("""
    Selamat datang ke **Sistem Automasi Analisis Klimatologi MetMalaysia**. Sistem ini dibangunkan untuk membantu proses transformasi dan verifikasi data iklim secara automatik:
    
    1. **Format Borang Standard:** Menukar data siri masa AAWS kepada Grid Rekod Hujan Harian Bulanan (JAN-DEC).
    2. **Piawaian WMO-No. 1203:** Menyemak kesempurnaan data (*Data Completeness*) sebelum pengiraan statistik dilaksanakan.
    3. **Visualisasi Interaktif:** Menjana visual matriks keamatan hujan (*Heatmap*) dan siri masa yang boleh di-*zoom* dan disesuaikan temanya.
    """)
    st.info("Sila muat naik fail raw AAWS di bar sisi kiri untuk memulakan analisis.")

# === TAB 2: RAINFALL (HUJAN) ===
with tab_rain:
    if uploaded_files:
        st.success(t["upload_success"].format(count=len(uploaded_files)))
        
        with st.spinner("Memproses data siri masa stesen..."):
            stations_data = process_multiple_aaws_files(uploaded_files)
            
        if stations_data:
            # Pilihan Muat Turun ZIP Kelompok
            col_zip1, col_zip2 = st.columns([3, 1])
            with col_zip1:
                st.subheader(f"📌 {len(stations_data)} Stesen Tersedia")
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
            qc_df_station = generate_qc_audit_table(df_stesen, qc_rule)
            total_months = len(qc_df_station)
            incomplete_months_count = (qc_df_station['Status Integriti'] == "❌ Incomplete").sum()
            total_missing_days = df_stesen['Rainfall_Numeric'].isna().sum()
            completeness_pct = ((len(df_stesen) - total_missing_days) / len(df_stesen)) * 100
            
            # 📊 KAD METRIK RINGKASAN INTEGRITI DATA
            m1, m2, m3, m4 = st.columns(4)
            m1.metric(t["station_name"], selected_stesen)
            m2.metric(t["record_period"], f"{min_yr} - {max_yr}")
            m3.metric("Tahap Kesempurnaan Data", f"{completeness_pct:.1f}%")
            m4.metric("Bulan Tidak Sah (Incomplete)", f"{incomplete_months_count} / {total_months} bln", delta=f"-{incomplete_months_count}" if incomplete_months_count > 0 else None, delta_color="inverse")
            
            # Kotak Amaran Pintar Jika Terdapat Bulan Incomplete
            if incomplete_months_count > 0:
                st.warning(f"⚠️ **Perhatian Pegawai:** Terdapat **{incomplete_months_count} bulan** yang gagal melepasi piawaian kesempurnaan data ({qc_rule}). Nilai jumlah bagi bulan-bulan tersebut ditandakan secara automatik sebagai `N.A (Incomplete)` pada borang Excel untuk mengelakkan ralat anggaran. Sila rujuk Tab **📋 Semakan Kualiti & WMO** untuk perincian.")
            
            # Sub-Tab: Borang Excel vs Visualisasi Interaktif
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
                
                st.write("#### 🔍 Pratonton Data Terkini (20 Baris Terawal)")
                st.dataframe(df_stesen.head(20), use_container_width=True)
                
            with sub_plots:
                # Kawalan Pilihan Graf & Tema Warna
                ctrl_col1, ctrl_col2 = st.columns(2)
                
                with ctrl_col1:
                    chart_choice = st.selectbox(
                        t["plot_type"],
                        options=[
                            "🌧️ Matriks Harian Hujan (Heatmap)",
                            "📈 Trend Jumlah Hujan Tahunan (Time-Series)",
                            "📊 Purata Taburan Hujan Bulanan (Climatological Normals)"
                        ]
                    )
                with ctrl_col2:
                    color_choice = st.selectbox(
                        t["color_theme"],
                        options=["Blues", "Viridis", "YlGnBu", "Spectral", "Plasma", "Turbo", "Teal"]
                    )
                
                # --- GRAF 1: HEATMAP KEAMATAN HUJAN ---
                if "Heatmap" in chart_choice:
                    years_list = sorted(df_stesen['Year'].unique())
                    chosen_year = st.selectbox("Pilih Tahun untuk Heatmap:", options=years_list, index=len(years_list)-1)
                    
                    df_heat = df_stesen[df_stesen['Year'] == chosen_year]
                    heat_pivot = df_heat.pivot(index='Day', columns='Month', values='Rainfall_Numeric')
                    heat_pivot = heat_pivot.reindex(index=range(1, 32), columns=range(1, 13))
                    
                    month_labels = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
                    
                    fig_heat = px.imshow(
                        heat_pivot,
                        labels=dict(x="Bulan", y="Hari", color="Hujan (mm)"),
                        x=month_labels,
                        y=[str(d) for d in range(1, 32)],
                        color_continuous_scale=color_choice,
                        aspect="auto",
                        title=f"Matriks Keamatan Hujan Harian bagi {selected_stesen} ({chosen_year})"
                    )
                    fig_heat.update_layout(height=650)
                    st.plotly_chart(fig_heat, use_container_width=True)
                    
                # --- GRAF 2: TREND TAHUNAN ---
                elif "Time-Series" in chart_choice:
                    annual_df = df_stesen.groupby('Year')['Rainfall_Numeric'].sum().reset_index()
                    mean_val = annual_df['Rainfall_Numeric'].mean()
                    
                    fig_trend = px.bar(
                        annual_df, 
                        x='Year', 
                        y='Rainfall_Numeric',
                        color='Rainfall_Numeric',
                        color_continuous_scale=color_choice,
                        labels={'Rainfall_Numeric': 'Jumlah Hujan (mm)', 'Year': 'Tahun'},
                        title=f"Trend Jumlah Hujan Tahunan bagi {selected_stesen} ({min_yr} - {max_yr})"
                    )
                    fig_trend.add_hline(
                        y=mean_val, 
                        line_dash="dash", 
                        line_color="red", 
                        annotation_text=f"Purata: {mean_val:.1f} mm"
                    )
                    st.plotly_chart(fig_trend, use_container_width=True)
                    
                # --- GRAF 3: PURATA BULANAN KLIMATOLOGI ---
                elif "Climatological Normals" in chart_choice:
                    month_names = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
                    month_df = df_stesen.groupby('Month')['Rainfall_Numeric'].mean().reset_index()
                    month_df['Month_Name'] = month_df['Month'].apply(lambda x: month_names[x-1])
                    
                    fig_norm = px.line(
                        month_df,
                        x='Month_Name',
                        y='Rainfall_Numeric',
                        markers=True,
                        title=f"Corak Purata Hujan Bulanan (Normal Profile) bagi {selected_stesen}",
                        labels={'Rainfall_Numeric': 'Purata Hujan (mm)', 'Month_Name': 'Bulan'}
                    )
                    fig_norm.update_traces(line_color="#1f77b4", marker=dict(size=9))
                    st.plotly_chart(fig_norm, use_container_width=True)
    else:
        st.info(t["info_upload"])

# === TAB 3: TEMPERATURE (SUHU) ===
with tab_temp:
    st.subheader("🌡️ Modul Analisis Suhu Udara (Temperature)")
    st.info("Modul Suhu Udara sedang diselaraskan untuk menyokong format Maximum / Minimum Temperature mengikut garis panduan WMO.")

# === TAB 4: QC & WMO COMPLETENESS ===
with tab_qc:
    st.subheader("📋 Laporan Audit Kualiti Data & Piawaian WMO-No. 1203")
    
    st.markdown("""
    Garis panduan rasmi **WMO-No. 1203 (Calculation of Climate Normals)** menetapkan:
    * **Peraturan 11/5 Hari (Data Bulanan):** Data bulanan dianggap tidak sah sekiranya mengandungi $\ge 11$ hari hilang atau $\ge 5$ hari hilang berturut-turut.
    * **Kriteria 80% (Normal 30-Tahun):** Pengiraan *Climate Normals* 30-tahun rasmi memerlukan sekurang-kurangnya 80% data bulanan yang lengkap ($\ge 24$ tahun daripada tempoh 30 tahun).
    """)
    
    if uploaded_files and 'stations_data' in locals() and stations_data:
        st.divider()
        qc_station_choice = st.selectbox("Pilih Stesen untuk Semakan Log QC:", options=list(stations_data.keys()), key="qc_select")
        qc_table = generate_qc_audit_table(stations_data[qc_station_choice], qc_rule)
        
        # Penapis Paparan: Tunjuk Semua atau Bulan Gagal Sahaja
        filter_col1, filter_col2 = st.columns([2, 1])
        with filter_col1:
            show_failed_only = st.checkbox("🔍 Paparkan Bulan Tidak Sah / Ada Data Hilang Sahaja", value=False)
        
        display_qc_table = qc_table[qc_table['Hari Hilang (NA)'] > 0] if show_failed_only else qc_table
        
        st.dataframe(display_qc_table, use_container_width=True)
        
        # Butang Muat Turun Log Audit QC
        csv_buffer = display_qc_table.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Muat Turun Laporan Log Audit QC (.CSV)",
            data=csv_buffer,
            file_name=f"Log_Audit_QC_{qc_station_choice}.csv",
            mime="text/csv"
        )
    else:
        st.info("Sila muat naik fail data di bar sisi kiri untuk melihat laporan audit kesempurnaan data stesen.")
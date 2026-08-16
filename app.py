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
    page_title="MetMalaysia Climatology System",
    layout="wide",
    page_icon="🌤️"
)

# ---------------------------------------------------------
# 2. KAMUS BAHASA (BILINGUAL DICTIONARY)
# ---------------------------------------------------------
TEXTS = {
    "BM": {
        "title": "Sistem Automasi Analisis Klimatologi",
        "subtitle": "Jabatan Meteorologi Malaysia (MetMalaysia) | Pejabat Meteorologi Sabah",
        "nav_home": "Utama",
        "nav_rain": "Hujan",
        "nav_temp": "Suhu",
        "nav_qc": "Semakan QC & WMO",
        "sidebar_header": "Data & Tetapan",
        "upload_label": "Muat naik fail siri masa AAWS (.xls / .xlsx)",
        "upload_success": "✅ {count} fail berjaya dimuat naik.",
        "qc_mode_label": "Piawaian Data Hilang (WMO):",
        "home_title": "Panduan & Standard Pengiraan",
        "home_desc": "Sistem automasi bagi penukaran data siri masa AAWS kepada format Borang Rekod Hujan Piawai serta penapisan integriti data mengikut panduan WMO-No. 1203.",
        "home_points": """
        * **Format Piawai:** Penjanaan automatik matriks 31 hari $\\times$ 12 bulan (JAN–DEC) bagi setiap tahun stesen.
        * **Tapisan Integriti WMO:** Pengesanan hari hilang (NA) mengikut piawaian saintifik sebelum jumlah/purata dikira.
        * **Visualisasi:** Analisis matriks keamatan hujan harian (Heatmap) dan siri masa tahunan secara interaktif.
        """,
        "btn_download_wmo": "📥 Muat Turun Dokumen WMO-No. 1203 (PDF)",
        "stations_ready": "{count} Stesen Dikesan",
        "download_zip": "📦 Muat Turun Semua (.ZIP)",
        "zip_filename": "Laporan_Klimatologi_Semua_Stesen.zip",
        "select_station": "Pilih Stesen:",
        "station_name": "Stesen",
        "record_period": "Tempoh Rekod",
        "total_records": "Jumlah Baris",
        "completeness_rate": "Kesempurnaan Data",
        "invalid_months": "Bulan Tidak Lengkap",
        "alert_incomplete": "⚠️ **Perhatian:** Terdapat **{count} bulan** tidak memenuhi piawaian data lengkap ({rule}). Statistik bagi bulan berkenaan ditandakan sebagai `N.A (Incomplete)`.",
        "subtab_form": "Borang Rekod & Muat Turun",
        "subtab_charts": "Visualisasi Data",
        "download_excel": "📥 Muat Turun Excel ({station})",
        "plot_type": "Pilih Jenis Graf:",
        "color_theme": "Skim Warna:",
        "opt_heatmap": "Matriks Harian Hujan (Heatmap)",
        "opt_trend": "Trend Hujan Tahunan (Time-Series)",
        "opt_normals": "Purata Profil Bulanan (Normals)",
        "select_year_heat": "Pilih Tahun:",
        "heat_title": "Matriks Keamatan Hujan Harian — {station} ({year})",
        "axis_month": "Bulan",
        "axis_day": "Hari",
        "axis_rain": "Hujan (mm)",
        "axis_year": "Tahun",
        "axis_avg_rain": "Purata Hujan (mm)",
        "trend_title": "Trend Hujan Tahunan — {station} ({min_yr} - {max_yr})",
        "trend_avg_label": "Purata",
        "norm_title": "Profil Purata Hujan Bulanan — {station}",
        "temp_title": "Modul Analisis Suhu Udara",
        "temp_desc": "Modul suhu akan diaktifkan apabila format data Maximum/Minimum dibekalkan.",
        "qc_title": "Log Audit Integriti Data (WMO-No. 1203)",
        "qc_filter_failed": "Tapis: Paparkan bulan tidak lengkap / ada NA sahaja",
        "download_qc_csv": "📥 Muat Turun Log Audit (.CSV)",
        "qc_col_year": "Tahun",
        "qc_col_month": "Bulan",
        "qc_col_na": "Hari Hilang (NA)",
        "qc_col_consec": "Maks. Berturut-turut",
        "qc_col_status": "Status Integriti",
        "qc_col_action": "Tindakan",
        "qc_status_valid": "Sah (Valid)",
        "qc_status_incomp": "Tidak Lengkap (Incomplete)",
        "qc_status_perfect": "100% Lengkap",
        "qc_act_calc": "Dikira (Abaikan NA)",
        "qc_act_reject": "Ditolak (N.A Incomplete)",
        "info_upload": "Sila muat naik fail data AAWS di bar sisi kiri untuk memulakan analisis."
    },
    "EN": {
        "title": "Climatology Analysis Automation System",
        "subtitle": "Malaysian Meteorological Department (MetMalaysia) | Sabah Meteorological Office",
        "nav_home": "Home",
        "nav_rain": "Rainfall",
        "nav_temp": "Temperature",
        "nav_qc": "QC & WMO Audit",
        "sidebar_header": "Data & Controls",
        "upload_label": "Upload AAWS time-series files (.xls / .xlsx)",
        "upload_success": "✅ {count} file(s) successfully uploaded.",
        "qc_mode_label": "Missing Data Standard (WMO):",
        "home_title": "User Guide & Calculation Standards",
        "home_desc": "Automated system for converting AAWS time-series data into standard climatological record sheets and screening data integrity per WMO-No. 1203 guidelines.",
        "home_points": """
        * **Standard Grid Format:** Automated 31-day $\\times$ 12-month (JAN-DEC) matrix generation per station-year.
        * **WMO Integrity Screening:** Evaluates missing days (NA) against scientific thresholds before statistical aggregation.
        * **Interactive Visuals:** Interactive daily intensity heatmaps and annual time-series charts.
        """,
        "btn_download_wmo": "📥 Download WMO-No. 1203 Document (PDF)",
        "stations_ready": "{count} Station(s) Found",
        "download_zip": "📦 Download All (.ZIP)",
        "zip_filename": "All_Stations_Climatology_Reports.zip",
        "select_station": "Select Station:",
        "station_name": "Station",
        "record_period": "Record Period",
        "total_records": "Total Rows",
        "completeness_rate": "Data Completeness",
        "invalid_months": "Incomplete Months",
        "alert_incomplete": "⚠️ **Advisory:** There are **{count} month(s)** failing data completeness criteria ({rule}). Monthly statistics are marked as `N.A (Incomplete)`.",
        "subtab_form": "Record Sheets & Download",
        "subtab_charts": "Data Visuals",
        "download_excel": "📥 Download Excel ({station})",
        "plot_type": "Select Chart Type:",
        "color_theme": "Color Theme:",
        "opt_heatmap": "Daily Rainfall Matrix (Heatmap)",
        "opt_trend": "Annual Rainfall Trend (Time-Series)",
        "opt_normals": "Monthly Average Profile (Normals)",
        "select_year_heat": "Select Year:",
        "heat_title": "Daily Rainfall Intensity Matrix — {station} ({year})",
        "axis_month": "Month",
        "axis_day": "Day",
        "axis_rain": "Rainfall (mm)",
        "axis_year": "Year",
        "axis_avg_rain": "Average Rainfall (mm)",
        "trend_title": "Annual Rainfall Trend — {station} ({min_yr} - {max_yr})",
        "trend_avg_label": "Average",
        "norm_title": "Monthly Average Rainfall Profile — {station}",
        "temp_title": "Air Temperature Analysis Module",
        "temp_desc": "The temperature module will be enabled once Maximum/Minimum data formats are supplied.",
        "qc_title": "Data Integrity Audit Log (WMO-No. 1203)",
        "qc_filter_failed": "Filter: Show incomplete / missing data months only",
        "download_qc_csv": "📥 Download Audit Log (.CSV)",
        "qc_col_year": "Year",
        "qc_col_month": "Month",
        "qc_col_na": "Missing Days (NA)",
        "qc_col_consec": "Max Consecutive",
        "qc_col_status": "Integrity Status",
        "qc_col_action": "Action",
        "qc_status_valid": "Valid",
        "qc_status_incomp": "Incomplete",
        "qc_status_perfect": "100% Complete",
        "qc_act_calc": "Calculated (Exclude NA)",
        "qc_act_reject": "Rejected (N.A Incomplete)",
        "info_upload": "Please upload AAWS data files in the left sidebar to start analysis."
    }
}

# ---------------------------------------------------------
# 3. SIDEBAR: KAWALAN & FAIL
# ---------------------------------------------------------
with st.sidebar:
    selected_lang = st.selectbox("Language / Bahasa", options=["Bahasa Melayu", "English"])
    lang_key = "BM" if selected_lang == "Bahasa Melayu" else "EN"
    t = TEXTS[lang_key]
    
    st.divider()
    st.markdown(f"### {t['sidebar_header']}")
    
    qc_rule = st.radio(
        t["qc_mode_label"],
        options=["WMO Standard (11/5 Rule)", "Strict Rule (5/3 Rule)", "No Filter (Raw Data)"],
        index=0
    )
    
    uploaded_files = st.file_uploader(
        t["upload_label"], 
        type=["xls", "xlsx"],
        accept_multiple_files=True
    )

# ---------------------------------------------------------
# 4. PENGEPALA MINIMALIS (CLEAN HEADER)
# ---------------------------------------------------------
header_col1, header_col2 = st.columns([1, 6])
with header_col1:
    try:
        st.image("logo_met.png", width=140)
    except:
        st.write("🌤️")

with header_col2:
    st.markdown(f"### **{t['title']}**")
    st.caption(f"🏛️ {t['subtitle']}")

st.divider()

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
            st.error(f"Ralat memproses {file.name}: {e}")
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
# 6. TAB NAVIGASI MINIMALIS
# ---------------------------------------------------------
tab_home, tab_rain, tab_temp, tab_qc = st.tabs([
    t["nav_home"], 
    t["nav_rain"], 
    t["nav_temp"], 
    t["nav_qc"]
])

# === TAB 1: UTAMA ===
with tab_home:
    st.markdown(f"#### {t['home_title']}")
    st.write(t["home_desc"])
    st.markdown(t["home_points"])
    
    st.write("")
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
        st.caption("ℹ️ Fail rujukan WMO-No. 1203 PDF tersedia dalam folder projek.")[cite: 1]
        
    if not uploaded_files:
        st.info(t["info_upload"])

# === TAB 2: HUJAN ===
with tab_rain:
    if uploaded_files and stations_data:
        col_top1, col_top2 = st.columns([3, 1])
        with col_top1:
            st.caption(t["stations_ready"].format(count=len(stations_data)))
        with col_top2:
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
        m4.metric(t["invalid_months"], f"{incomplete_months_count} / {total_months}")
        
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
            st.dataframe(df_stesen.head(15), use_container_width=True)
            
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
                color_choice = st.selectbox(t["color_theme"], options=["Blues", "Viridis", "YlGnBu", "Spectral", "Plasma", "Teal"])
                
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
                fig_heat.update_layout(height=600, margin=dict(l=20, r=20, t=40, b=20))
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
                fig_trend.update_layout(margin=dict(l=20, r=20, t=40, b=20))
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
                fig_norm.update_traces(line_color="#1f77b4", marker=dict(size=8))
                fig_norm.update_layout(margin=dict(l=20, r=20, t=40, b=20))
                st.plotly_chart(fig_norm, use_container_width=True)
    else:
        st.info(t["info_upload"])

# === TAB 3: SUHU ===
with tab_temp:
    st.markdown(f"#### {t['temp_title']}")
    st.info(t["temp_desc"])

# === TAB 4: QC ===
with tab_qc:
    st.markdown(f"#### {t['qc_title']}")
    if uploaded_files and stations_data:
        qc_station_choice = st.selectbox(t["select_station"], options=list(stations_data.keys()), key="qc_select")
        qc_table = generate_qc_audit_table(stations_data[qc_station_choice], qc_rule, t)
        
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
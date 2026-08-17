import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
import plotly.express as px
import plotly.graph_objects as go

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
        "title": "Sistem Automasi Analisis Klimatologi",
        "subtitle": "Jabatan Meteorologi Malaysia (MetMalaysia) | Pejabat Meteorologi Sabah",
        "nav_analysis": "📊 Analisis Parameter",
        "nav_qc": "📋 Semakan Kualiti & WMO",
        "sidebar_header": "📁 Data & Tetapan",
        "manual_header": "📖 Manual Penggunaan & Rujukan",
        "manual_desc": """
        **Panduan Ringkas Penggunaan:**
        1. **Muat Naik Fail:** Masukkan fail raw AAWS (`.xls` / `.xlsx`) di menu bar sisi.
        2. **Pilih Parameter:** Gunakan butang di Zon Kawalan untuk bertukar antara **Hujan** dan **Suhu**.
        3. **Audit WMO:** Semakan integriti data hilang (`NA`) automatik mengikut panduan **WMO-No. 1203**.
        4. **Eksport Data & Visual:** Jana borang Excel piawai atau muat turun graf `.HTML` interaktif.
        """,
        "btn_download_wmo": "📥 Muat Turun WMO-No. 1203 (PDF)",
        "upload_label": "Muat naik fail siri masa AAWS (.xls / .xlsx)",
        "qc_mode_label": "⚙️ Piawaian Data Hilang (WMO):",
        "download_zip": "📦 Muat Turun Semua ({param}) [.ZIP]",
        "zip_filename": "Laporan_Klimatologi_{param}_Semua_Stesen.zip",
        "select_param": "Parameter Cerapan:",
        "param_rain": "🌧️ Hujan (Rainfall)",
        "param_temp": "🌡️ Suhu Udara (Temperature)",
        "no_param_data": "⚠️ **Tiada data {param_name} dikesan.** Parameter dikesan dalam fail: **{detected_params}**. Sila tukar pilihan atau muat naik fail yang sepadan.",
        "select_station": "Pilih Stesen:",
        "station_name": "Stesen Cerapan",
        "record_period": "Tempoh Siri Masa",
        "completeness_rate": "Tahap Kesempurnaan",
        "invalid_months": "Status Integriti Bulanan",
        "alert_incomplete": "⚠️ **Perhatian Pegawai:** Terdapat **{count} bulan** tidak memenuhi piawaian kesempurnaan data ({rule}). Nilai bagi bulan berkenaan ditandakan sebagai `N.A (Incomplete)`.",
        "subtab_form": "📄 Borang Rekod Piawai & Eksport",
        "subtab_charts": "📈 Visualisasi Data & Analisis Julat",
        "download_excel": "📥 Muat Turun Excel ({station})",
        "download_chart_html": "📥 Muat Turun Graf Interaktif ({chart}) [.HTML]",
        "plot_type": "Pilih Perspektif Visual:",
        "color_theme": "Skim Warna Graf:",
        "opt_heatmap": "🌧️ Matriks Harian (Heatmap)",
        "opt_trend": "📈 Trend Siri Masa Tahunan",
        "opt_normals": "📊 Profil Normal Bulanan & Julat Min-Maks (Range Envelope)",
        "opt_anomaly": "📉 Anomali Iklim Siri Masa (Climate Anomaly)",
        "select_year_heat": "Pilih Tahun Kalendar:",
        "heat_title": "Matriks Harian {param} — {station} ({year})",
        "axis_month": "Bulan",
        "axis_day": "Hari",
        "axis_year": "Tahun",
        "trend_title": "Trend Tahunan {param} — {station} ({min_yr} - {max_yr})",
        "trend_avg_label": "Purata Normal",
        "norm_title": "Profil Purata Bulanan & Julat Ekstrem (Min-Max) — {station}",
        "anomaly_title": "Siri Masa Anomali {param} — {station} ({min_yr} - {max_yr})",
        "axis_control_header": "⚙️ Tetapan Skala Paksi Graf (Custom Axis Range)",
        "axis_enable_custom": "Seragamkan / Kunci Skala Paksi Graf (Standardize Axis)",
        "axis_min_label": "Nilai Had Bawah (Min):",
        "axis_max_label": "Nilai Had Atas (Max):",
        "qc_title": "Log Audit Integriti Data (WMO-No. 1203)",
        "qc_filter_failed": "🔍 Paparkan bulan tidak lengkap / ada data hilang (NA) sahaja",
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
        "info_upload": "👈 Sila muat naik fail raw AAWS di menu bar sisi kiri untuk memulakan analisis."
    },
    "EN": {
        "title": "Climatology Analysis Automation System",
        "subtitle": "Malaysian Meteorological Department (MetMalaysia) | Sabah Meteorological Office",
        "nav_analysis": "📊 Parameter Analysis",
        "nav_qc": "📋 QC & WMO Audit",
        "sidebar_header": "📁 Data & Settings",
        "manual_header": "📖 User Manual & Standards",
        "manual_desc": """
        **Quick User Guide:**
        1. **Upload Files:** Upload raw AAWS files (`.xls` / `.xlsx`) via the sidebar menu.
        2. **Select Parameter:** Switch dynamically between **Rainfall** and **Air Temperature**.
        3. **WMO Screening:** Missing data (`NA`) is audited automatically in compliance with **WMO-No. 1203**.
        4. **Export Sheets & Visuals:** Generate Excel sheets or export interactive standalone `.HTML` charts.
        """,
        "btn_download_wmo": "📥 Download WMO-No. 1203 (PDF)",
        "upload_label": "Upload AAWS time-series files (.xls / .xlsx)",
        "qc_mode_label": "⚙️ Missing Data Standard (WMO):",
        "download_zip": "📦 Download All ({param}) [.ZIP]",
        "zip_filename": "Climatology_Report_{param}_All_Stations.zip",
        "select_param": "Observation Parameter:",
        "param_rain": "🌧️ Rainfall",
        "param_temp": "🌡️ Air Temperature",
        "no_param_data": "⚠️ **No {param_name} data detected.** Detected parameter in file: **{detected_params}**. Please switch toggle or upload matching data.",
        "select_station": "Select Station:",
        "station_name": "Observation Station",
        "record_period": "Time-Series Period",
        "completeness_rate": "Completeness Level",
        "invalid_months": "Monthly Integrity Status",
        "alert_incomplete": "⚠️ **Officer Advisory:** There are **{count} month(s)** failing completeness standards ({rule}). Values are flagged as `N.A (Incomplete)`.",
        "subtab_form": "📄 Standard Sheets & Export",
        "subtab_charts": "📈 Interactive Visuals & Range Analytics",
        "download_excel": "📥 Download Excel ({station})",
        "download_chart_html": "📥 Download Interactive Chart ({chart}) [.HTML]",
        "plot_type": "Select Visual Perspective:",
        "color_theme": "Chart Color Theme:",
        "opt_heatmap": "🌧️ Daily Matrix (Heatmap)",
        "opt_trend": "📈 Annual Time-Series Trend",
        "opt_normals": "📊 Monthly Normals & Range Envelope (Min-Max)",
        "opt_anomaly": "📉 Time-Series Climate Anomaly",
        "select_year_heat": "Select Calendar Year:",
        "heat_title": "Daily Matrix for {param} — {station} ({year})",
        "axis_month": "Month",
        "axis_day": "Day",
        "axis_year": "Year",
        "trend_title": "Annual Trend for {param} — {station} ({min_yr} - {max_yr})",
        "trend_avg_label": "Normal Mean",
        "norm_title": "Monthly Average Profile & Range Envelope (Min-Max) — {station}",
        "anomaly_title": "{param} Anomaly Time-Series — {station} ({min_yr} - {max_yr})",
        "axis_control_header": "⚙️ Custom Axis Scaling & Standardization",
        "axis_enable_custom": "Standardize / Lock Graph Axis Range",
        "axis_min_label": "Lower Bound (Min):",
        "axis_max_label": "Upper Bound (Max):",
        "qc_title": "Data Integrity Audit Log (WMO-No. 1203)",
        "qc_filter_failed": "🔍 Show incomplete / missing data months only",
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
        "info_upload": "👈 Please upload raw AAWS files in the left sidebar to start analysis."
    }
}

# ---------------------------------------------------------
# 3. SIDEBAR: KAWALAN PENGGUNA & FAIL
# ---------------------------------------------------------
with st.sidebar:
    selected_lang = st.selectbox("🌐 Bahasa / Language", options=["Bahasa Melayu", "English"])
    lang_key = "BM" if selected_lang == "Bahasa Melayu" else "EN"
    t = TEXTS[lang_key]
    
    with st.expander(t["manual_header"], expanded=False):
        st.markdown(t["manual_desc"])
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
        except Exception:
            st.caption("ℹ️ Fail rujukan WMO-No. 1203 sedia ada.")
            
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
# 4. PENGEPALA APLIKASI
# ---------------------------------------------------------
header_col1, header_col2 = st.columns([1.1, 6])
with header_col1:
    try:
        st.image("logo_met.png", width=135)
    except Exception:
        st.write("🌤️")

with header_col2:
    st.markdown(f"### **{t['title']}**")
    st.caption(f"🏛️ {t['subtitle']}")

st.divider()

# ---------------------------------------------------------
# 5. ENJIN PEMPROSESAN DATA & WMO
# ---------------------------------------------------------
def detect_parameter_type(header_text):
    text_lower = header_text.lower()
    temp_keywords = ['temperature', 'suhu', 'temp', 'celsius', '°c', 'deg c', 'degc']
    rain_keywords = ['rainfall', 'hujan', 'rain', 'presipitasi', 'rf', 'mm']
    is_temp = any(k in text_lower for k in temp_keywords)
    is_rain = any(k in text_lower for k in rain_keywords)
    if is_temp and not is_rain:
        return "Temperature"
    elif is_rain and not is_temp:
        return "Rainfall"
    elif is_temp:
        return "Temperature"
    return "Rainfall"

def process_multiple_aaws_files(files_list):
    all_data = {"Rainfall": {}, "Temperature": {}}
    for file in files_list:
        try:
            xls = pd.ExcelFile(file)
            for sheet in xls.sheet_names:
                if sheet.lower() == 'datalist':
                    continue
                df = pd.read_excel(xls, sheet_name=sheet)
                header_dump = " ".join([str(val) for val in df.iloc[:11, :].values.flatten()])
                detected_param = detect_parameter_type(header_dump)
                
                raw_station_text = str(df.iloc[2, 0]) if df.shape[0] > 2 else "Station"
                station_name = raw_station_text.split(':', 1)[1].strip() if ':' in raw_station_text else raw_station_text.replace("Station", "").strip()
                if not station_name or station_name == "nan":
                    station_name = f"Station_{sheet}"
                    
                data = df.iloc[11:].copy().iloc[:, :4]
                data.columns = ['Year', 'Month', 'Day', 'Value']
                data['Year'] = pd.to_numeric(data['Year'], errors='coerce')
                data['Month'] = pd.to_numeric(data['Month'], errors='coerce')
                data['Day'] = pd.to_numeric(data['Day'], errors='coerce')
                data['Value_Numeric'] = pd.to_numeric(data['Value'], errors='coerce')
                data['Value_Display'] = data['Value']
                
                data = data.dropna(subset=['Year', 'Month', 'Day'])
                data = data[(data['Year'] > 1900) & (data['Year'] < 2100)]
                data['Year'] = data['Year'].astype(int)
                data['Month'] = data['Month'].astype(int)
                data['Day'] = data['Day'].astype(int)
                
                if not data.empty:
                    target_dict = all_data[detected_param]
                    if station_name in target_dict:
                        combined_df = pd.concat([target_dict[station_name], data], ignore_index=True)
                        combined_df = combined_df.drop_duplicates(subset=['Year', 'Month', 'Day'])
                        target_dict[station_name] = combined_df
                    else:
                        target_dict[station_name] = data
        except Exception as e:
            st.error(f"Ralat memproses {file.name}: {e}")
    return all_data

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
        pivot_num = df_yr.pivot(index='Day', columns='Month', values='Value_Numeric').reindex(index=range(1, 32), columns=range(1, 13))
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

def generate_excel_for_station(station_name, df_station, rule, param_type):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        years = sorted(df_station['Year'].unique())
        month_names = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
        has_written = False
        
        for yr in years:
            df_yr = df_station[df_station['Year'] == yr]
            if df_yr.empty:
                continue
            pivot_num = df_yr.pivot(index='Day', columns='Month', values='Value_Numeric').reindex(index=range(1, 32), columns=range(1, 13))
            pivot_display = df_yr.pivot(index='Day', columns='Month', values='Value_Display').reindex(index=range(1, 32), columns=range(1, 13))
            
            stat_row1, stat_row2, stat_row3, stat_row4 = [], [], [], []
            
            for m in range(1, 13):
                col_data = pivot_num[m]
                is_valid, _, _ = evaluate_month_qc(col_data, rule)
                
                if is_valid:
                    if param_type == "Rainfall":
                        tot = col_data.sum(skipna=True)
                        stat_row1.append(round(tot, 1) if pd.notna(tot) else "N.A")
                        stat_row2.append((col_data > 0.1).sum())
                        stat_row3.append(round(col_data.max(skipna=True), 1) if pd.notna(col_data.max(skipna=True)) else "N.A")
                        stat_row4.append(int(col_data.idxmax(skipna=True)) if pd.notna(col_data.idxmax(skipna=True)) else "-")
                    else:
                        mean_temp = col_data.mean(skipna=True)
                        stat_row1.append(round(mean_temp, 1) if pd.notna(mean_temp) else "N.A")
                        stat_row2.append(round(col_data.max(skipna=True), 1) if pd.notna(col_data.max(skipna=True)) else "N.A")
                        stat_row3.append(round(col_data.min(skipna=True), 1) if pd.notna(col_data.min(skipna=True)) else "N.A")
                        rng = col_data.max(skipna=True) - col_data.min(skipna=True)
                        stat_row4.append(round(rng, 1) if pd.notna(rng) else "N.A")
                else:
                    stat_row1.append("N.A (Incomplete)")
                    stat_row2.append("N.A")
                    stat_row3.append("N.A")
                    stat_row4.append("-")
                    
            report_df = pivot_display.copy()
            if param_type == "Rainfall":
                report_df.loc['TOTAL (mm)'] = stat_row1
                report_df.loc['No. Of Days (>0.1mm)'] = stat_row2
                report_df.loc['Highest Fall (mm)'] = stat_row3
                report_df.loc['Date of Highest'] = stat_row4
            else:
                report_df.loc['MEAN TEMP (°C)'] = stat_row1
                report_df.loc['MAX TEMP (°C)'] = stat_row2
                report_df.loc['MIN TEMP (°C)'] = stat_row3
                report_df.loc['TEMP RANGE (°C)'] = stat_row4
                
            report_df.columns = month_names
            report_df.index.name = "DATE"
            report_df.to_excel(writer, sheet_name=str(yr)[:31])
            has_written = True
            
        if not has_written:
            pd.DataFrame({"Note": ["No Data"]}).to_excel(writer, sheet_name="No Data")
    output.seek(0)
    return output

all_data = {"Rainfall": {}, "Temperature": {}}
if uploaded_files:
    all_data = process_multiple_aaws_files(uploaded_files)

# ---------------------------------------------------------
# 6. TAB NAVIGASI KERJA
# ---------------------------------------------------------
tab_analysis, tab_qc = st.tabs([
    t["nav_analysis"], 
    t["nav_qc"]
])

# =========================================================
# TAB 1: ANALISIS PARAMETER (3-ZON LINEAR ARCHITECTURE)
# =========================================================
with tab_analysis:
    if uploaded_files:
        # -------------------------------------------------
        # ZON 1: KAWALAN UTAMA (CONTROL BAR)
        # -------------------------------------------------
        ctrl_col1, ctrl_col2, ctrl_col3 = st.columns([1.8, 1.8, 1.4])
        
        with ctrl_col1:
            chosen_param_label = st.radio(
                t["select_param"],
                options=[t["param_rain"], t["param_temp"]],
                horizontal=True
            )
            param_mode = "Rainfall" if chosen_param_label == t["param_rain"] else "Temperature"
            unit_str = "mm" if param_mode == "Rainfall" else "°C"
            
        stations_data = all_data.get(param_mode, {})
        
        if not stations_data:
            detected = [k for k, v in all_data.items() if len(v) > 0]
            detected_str = ", ".join(detected) if detected else "Tiada"
            st.warning(t["no_param_data"].format(param_name=param_mode, detected_params=detected_str))
        else:
            with ctrl_col2:
                selected_stesen = st.selectbox(t["select_station"], options=list(stations_data.keys()))
                
            with ctrl_col3:
                st.write("")  # Alignment spacer
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for st_name, st_df in stations_data.items():
                        excel_bytes = generate_excel_for_station(st_name, st_df, qc_rule, param_mode)
                        clean_st_name = "".join([c for c in st_name if c.isalnum() or c in (' ', '_', '-')]).strip()
                        zip_file.writestr(f"{param_mode}_{clean_st_name.replace(' ', '_')}.xlsx", excel_bytes.getvalue())
                zip_buffer.seek(0)
                st.download_button(
                    label=t["download_zip"].format(param=param_mode),
                    data=zip_buffer,
                    file_name=t["zip_filename"].format(param=param_mode),
                    mime="application/zip",
                    use_container_width=True,
                    type="primary"
                )
                
            # -------------------------------------------------
            # ZON 2: METRIK STATUS PANTAS (KPI HERO CARDS)
            # -------------------------------------------------
            df_stesen = stations_data[selected_stesen]
            min_yr, max_yr = df_stesen['Year'].min(), df_stesen['Year'].max()
            qc_df_station = generate_qc_audit_table(df_stesen, qc_rule, t)
            total_months = len(qc_df_station)
            incomplete_months_count = (qc_df_station[t["qc_col_status"]] == t["qc_status_incomp"]).sum()
            total_missing_days = df_stesen['Value_Numeric'].isna().sum()
            completeness_pct = ((len(df_stesen) - total_missing_days) / len(df_stesen)) * 100
            
            kpi_box = st.container(border=True)
            with kpi_box:
                m1, m2, m3, m4 = st.columns(4)
                m1.metric(t["station_name"], selected_stesen)
                m2.metric(t["record_period"], f"{min_yr} – {max_yr}")
                m3.metric(t["completeness_rate"], f"{completeness_pct:.1f}%")
                m4.metric(t["invalid_months"], f"{incomplete_months_count} / {total_months} Incomplete")
                
            if incomplete_months_count > 0:
                st.warning(t["alert_incomplete"].format(count=incomplete_months_count, rule=qc_rule))
                
            # -------------------------------------------------
            # ZON 3: RUANG KERJA UTAMA (DUAL-VIEW WORKSPACE)
            # -------------------------------------------------
            sub_form, sub_plots = st.tabs([t["subtab_form"], t["subtab_charts"]])
            
            # SUB-VIEW 1: BORANG EXCEL
            with sub_form:
                excel_file = generate_excel_for_station(selected_stesen, df_stesen, qc_rule, param_mode)
                clean_st_name = "".join([c for c in selected_stesen if c.isalnum() or c in (' ', '_', '-')]).strip()
                st.download_button(
                    label=t["download_excel"].format(station=selected_stesen),
                    data=excel_file,
                    file_name=f"Borang_{param_mode}_{clean_st_name.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.dataframe(df_stesen.head(15), use_container_width=True)
                
            # SUB-VIEW 2: VISUALISASI DATA & ANALISIS JULAT
            with sub_plots:
                # Kawalan Pilihan Visual & Skim Warna
                plot_ctrl1, plot_ctrl2 = st.columns([2, 1.2])
                chart_options_map = {
                    t["opt_normals"]: "Normals_Range",
                    t["opt_anomaly"]: "Anomaly",
                    t["opt_trend"]: "Trend",
                    t["opt_heatmap"]: "Heatmap"
                }
                with plot_ctrl1:
                    chart_choice_label = st.selectbox(t["plot_type"], options=list(chart_options_map.keys()))
                    chart_choice = chart_options_map[chart_choice_label]
                with plot_ctrl2:
                    default_idx = 0 if param_mode == "Rainfall" else 1
                    color_choice = st.selectbox(t["color_theme"], options=["Blues", "Thermal", "Viridis", "YlGnBu", "Spectral", "Plasma", "Teal"], index=default_idx)
                    
                # Kawalan Skala Paksi Tersuai
                with st.expander(t["axis_control_header"], expanded=False):
                    use_custom_axis = st.checkbox(t["axis_enable_custom"], value=False)
                    if use_custom_axis:
                        ax_col1, ax_col2 = st.columns(2)
                        with ax_col1:
                            axis_min_val = st.number_input(
                                t["axis_min_label"], 
                                value=-500.0 if chart_choice == "Anomaly" else (15.0 if param_mode == "Temperature" else 0.0), 
                                step=10.0 if param_mode == "Rainfall" else 1.0
                            )
                        with ax_col2:
                            axis_max_val = st.number_input(
                                t["axis_max_label"], 
                                value=500.0 if chart_choice == "Anomaly" else (3500.0 if param_mode == "Rainfall" else 40.0), 
                                step=50.0 if param_mode == "Rainfall" else 1.0
                            )
                
                fig_config = {
                    'toImageButtonOptions': {
                        'format': 'png',
                        'filename': f'{param_mode}_{selected_stesen}_{chart_choice}',
                        'height': 800,
                        'width': 1400,
                        'scale': 2.5
                    }
                }
                
                # KANVAS 1: PROFIL NORMAL & CARTA JULAT (RIBBON ENVELOPE)
                if chart_choice == "Normals_Range":
                    month_names = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
                    month_stats = df_stesen.groupby('Month')['Value_Numeric'].agg(['mean', 'min', 'max']).reset_index()
                    month_stats['Month_Name'] = month_stats['Month'].apply(lambda x: month_names[x-1])
                    
                    fig_range = go.Figure()
                    fig_range.add_trace(go.Scatter(
                        x=month_stats['Month_Name'], y=month_stats['max'],
                        mode='lines', line=dict(width=0), showlegend=False, hoverinfo='skip'
                    ))
                    fig_range.add_trace(go.Scatter(
                        x=month_stats['Month_Name'], y=month_stats['min'],
                        mode='lines', line=dict(width=0), fill='tonexty',
                        fillcolor='rgba(31, 119, 180, 0.2)' if param_mode == 'Rainfall' else 'rgba(214, 39, 40, 0.2)',
                        name='Julat Ekstrem (Min-Max Range)',
                        hovertemplate='Julat Minimum: %{y:.1f}' + f' {unit_str}<extra></extra>'
                    ))
                    main_line_color = '#1f77b4' if param_mode == 'Rainfall' else '#d62728'
                    fig_range.add_trace(go.Scatter(
                        x=month_stats['Month_Name'], y=month_stats['mean'],
                        mode='lines+markers', name=f'Purata Normal ({param_mode})',
                        line=dict(color=main_line_color, width=3), marker=dict(size=8),
                        hovertemplate='Bulan: %{x}<br>Purata: %{y:.1f}' + f' {unit_str}<extra></extra>'
                    ))
                    fig_range.update_layout(
                        title=t["norm_title"].format(param=param_mode, station=selected_stesen),
                        xaxis_title=t["axis_month"], yaxis_title=f"{param_mode} ({unit_str})",
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                        margin=dict(l=20, r=20, t=50, b=20)
                    )
                    if use_custom_axis:
                        fig_range.update_layout(yaxis_range=[axis_min_val, axis_max_val])
                        
                    st.plotly_chart(fig_range, use_container_width=True, config=fig_config)
                    
                    html_range = fig_range.to_html(include_plotlyjs='cdn').encode('utf-8')
                    st.download_button(
                        label=t["download_chart_html"].format(chart="Normals_Range"),
                        data=html_range,
                        file_name=f"Normals_Range_{param_mode}_{selected_stesen}.html",
                        mime="text/html"
                    )
                    
                    peak_m = month_stats.loc[month_stats['mean'].idxmax()]
                    dry_m = month_stats.loc[month_stats['mean'].idxmin()]
                    
                    st.info(f"""
                    📌 **Rumusan Profil Bulanan & Carta Julat (Min–Maks):**
                    * **Kawasan Berlorek (Ribbon):** Menunjukkan julat variabiliti antara bacaan minimum dan maksimum yang pernah direkodkan bagi setiap bulan sepanjang siri masa.
                    * **Bulan Kemuncak (Peak Month):** Bulan **{peak_m['Month_Name']}** dengan purata tertinggi `{peak_m['mean']:.1f} {unit_str}` (Maks: `{peak_m['max']:.1f} {unit_str}`).
                    * **Bulan Terendah:** Bulan **{dry_m['Month_Name']}** dengan purata `{dry_m['mean']:.1f} {unit_str}` (Min: `{dry_m['min']:.1f} {unit_str}`).
                    """)
                    
                # KANVAS 2: ANOMALI IKLIM (DIVERGING BAR)
                elif chart_choice == "Anomaly":
                    if param_mode == "Rainfall":
                        annual_df = df_stesen.groupby('Year')['Value_Numeric'].sum().reset_index()
                    else:
                        annual_df = df_stesen.groupby('Year')['Value_Numeric'].mean().reset_index()
                        
                    norm_mean = annual_df['Value_Numeric'].mean()
                    annual_df['Anomaly'] = annual_df['Value_Numeric'] - norm_mean
                    
                    if param_mode == "Rainfall":
                        annual_df['Status'] = annual_df['Anomaly'].apply(lambda x: 'Lebih Normal (Wet)' if x >= 0 else 'Kurang Normal (Dry)')
                        color_map = {'Lebih Normal (Wet)': '#1f77b4', 'Kurang Normal (Dry)': '#d62728'}
                    else:
                        annual_df['Status'] = annual_df['Anomaly'].apply(lambda x: 'Lebih Panas (Warm)' if x >= 0 else 'Lebih Sejuk (Cool)')
                        color_map = {'Lebih Panas (Warm)': '#d62728', 'Lebih Sejuk (Cool)': '#1f77b4'}
                    
                    fig_anom = px.bar(
                        annual_df, x='Year', y='Anomaly', color='Status',
                        color_discrete_map=color_map,
                        labels={'Anomaly': f'Anomali ({unit_str})', 'Year': t['axis_year']},
                        title=t["anomaly_title"].format(param=param_mode, station=selected_stesen, min_yr=min_yr, max_yr=max_yr)
                    )
                    fig_anom.add_hline(y=0, line_color="black", line_width=1.5)
                    if use_custom_axis:
                        fig_anom.update_layout(yaxis_range=[axis_min_val, axis_max_val])
                    fig_anom.update_layout(margin=dict(l=20, r=20, t=40, b=20))
                    st.plotly_chart(fig_anom, use_container_width=True, config=fig_config)
                    
                    html_anom = fig_anom.to_html(include_plotlyjs='cdn').encode('utf-8')
                    st.download_button(
                        label=t["download_chart_html"].format(chart="Anomaly"),
                        data=html_anom,
                        file_name=f"Anomaly_{param_mode}_{selected_stesen}.html",
                        mime="text/html"
                    )
                    
                    highest_anom = annual_df.loc[annual_df['Anomaly'].idxmax()]
                    lowest_anom = annual_df.loc[annual_df['Anomaly'].idxmin()]
                    
                    st.info(f"""
                    📌 **Rumusan Anomali Iklim (Rujukan Normal: `{norm_mean:.1f} {unit_str}`):**
                    * Garisan 0 mewakili purata normal iklim. Bar menunjukkan sisihan (*departure*) daripada purata jangka panjang.
                    * **Anomali Positif Tertinggi:** Tahun **{int(highest_anom['Year'])}** (`+{highest_anom['Anomaly']:.1f} {unit_str}`).
                    * **Anomali Negatif Terendah:** Tahun **{int(lowest_anom['Year'])}** (`{lowest_anom['Anomaly']:.1f} {unit_str}`).
                    """)
                    
                # KANVAS 3: TREND SIRI MASA TAHUNAN
                elif chart_choice == "Trend":
                    if param_mode == "Rainfall":
                        annual_df = df_stesen.groupby('Year')['Value_Numeric'].sum().reset_index()
                        val_label = f"Jumlah Hujan ({unit_str})"
                    else:
                        annual_df = df_stesen.groupby('Year')['Value_Numeric'].mean().reset_index()
                        val_label = f"Purata Suhu ({unit_str})"
                        
                    mean_val = annual_df['Value_Numeric'].mean()
                    max_yr_row = annual_df.loc[annual_df['Value_Numeric'].idxmax()]
                    min_yr_row = annual_df.loc[annual_df['Value_Numeric'].idxmin()]
                    
                    fig_trend = px.bar(
                        annual_df, x='Year', y='Value_Numeric', color='Value_Numeric',
                        color_continuous_scale=color_choice,
                        labels={'Value_Numeric': val_label, 'Year': t["axis_year"]},
                        title=t["trend_title"].format(param=param_mode, station=selected_stesen, min_yr=min_yr, max_yr=max_yr)
                    )
                    fig_trend.add_hline(y=mean_val, line_dash="dash", line_color="red", annotation_text=f"{t['trend_avg_label']}: {mean_val:.1f} {unit_str}")
                    if use_custom_axis:
                        fig_trend.update_layout(yaxis_range=[axis_min_val, axis_max_val])
                    fig_trend.update_layout(margin=dict(l=20, r=20, t=40, b=20))
                    st.plotly_chart(fig_trend, use_container_width=True, config=fig_config)
                    
                    html_trend = fig_trend.to_html(include_plotlyjs='cdn').encode('utf-8')
                    st.download_button(
                        label=t["download_chart_html"].format(chart="Trend"),
                        data=html_trend,
                        file_name=f"Trend_{param_mode}_{selected_stesen}.html",
                        mime="text/html"
                    )
                    
                    st.info(f"""
                    📌 **Rumusan Trend Siri Masa ({min_yr}–{max_yr}):**
                    * **Purata Normal Jangka Panjang:** `{mean_val:.1f} {unit_str}` (garisan putus merah).
                    * **Tahun Tertinggi / Ekstrem:** Tahun **{int(max_yr_row['Year'])}** dengan nilai `{max_yr_row['Value_Numeric']:.1f} {unit_str}`.
                    * **Tahun Terendah:** Tahun **{int(min_yr_row['Year'])}** dengan nilai `{min_yr_row['Value_Numeric']:.1f} {unit_str}`.
                    """)
                    
                # KANVAS 4: MATRIKS KEAMATAN HARIAN (HEATMAP)
                elif chart_choice == "Heatmap":
                    years_list = sorted(df_stesen['Year'].unique())
                    chosen_year = st.selectbox(t["select_year_heat"], options=years_list, index=len(years_list)-1)
                    df_heat = df_stesen[df_stesen['Year'] == chosen_year]
                    heat_pivot = df_heat.pivot(index='Day', columns='Month', values='Value_Numeric').reindex(index=range(1, 32), columns=range(1, 13))
                    month_labels = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
                    
                    fig_heat = px.imshow(
                        heat_pivot,
                        labels=dict(x=t["axis_month"], y=t["axis_day"], color=f"{param_mode} ({unit_str})"),
                        x=month_labels, y=[str(d) for d in range(1, 32)],
                        color_continuous_scale=color_choice, aspect="auto",
                        range_color=[axis_min_val, axis_max_val] if use_custom_axis else None,
                        title=t["heat_title"].format(param=param_mode, station=selected_stesen, year=chosen_year)
                    )
                    fig_heat.update_layout(height=580, margin=dict(l=20, r=20, t=40, b=20))
                    st.plotly_chart(fig_heat, use_container_width=True, config=fig_config)
                    
                    html_heat = fig_heat.to_html(include_plotlyjs='cdn').encode('utf-8')
                    st.download_button(
                        label=t["download_chart_html"].format(chart="Heatmap"),
                        data=html_heat,
                        file_name=f"Heatmap_{param_mode}_{selected_stesen}_{chosen_year}.html",
                        mime="text/html"
                    )
                    
                    max_day_val = df_heat['Value_Numeric'].max()
                    max_day_date = df_heat.loc[df_heat['Value_Numeric'].idxmax()] if pd.notna(max_day_val) else None
                    date_str = f"{int(max_day_date['Day'])}/{int(max_day_date['Month'])}/{chosen_year}" if max_day_date is not None else "-"
                    
                    st.info(f"""
                    📌 **Rumusan Matriks Harian ({chosen_year}):**
                    * Graf ini memaparkan corak keamatan harian sepanjang 12 bulan. Petak berwarna lebih gelap menandakan bacaan yang lebih tinggi.
                    * **Bacaan Harian Tertinggi:** `{max_day_val:.1f} {unit_str}` dicatatkan pada tarikh **{date_str}**.
                    * Petak kosong/putih menunjukkan hari tanpa cerapan hujan atau suhu minimum.
                    """)
    else:
        st.info(t["info_upload"])

# =========================================================
# TAB 2: SEMAKAN KUALITI & WMO AUDIT
# =========================================================
with tab_qc:
    st.markdown(f"#### {t['qc_title']}")
    if uploaded_files:
        qc_param_choice = st.radio(
            t["select_param"],
            options=[t["param_rain"], t["param_temp"]],
            horizontal=True,
            key="qc_param_radio"
        )
        param_mode_qc = "Rainfall" if qc_param_choice == t["param_rain"] else "Temperature"
        stations_qc = all_data.get(param_mode_qc, {})
        
        if not stations_qc:
            detected = [k for k, v in all_data.items() if len(v) > 0]
            detected_str = ", ".join(detected) if detected else "Tiada"
            st.warning(t["no_param_data"].format(param_name=param_mode_qc, detected_params=detected_str))
        else:
            qc_station_choice = st.selectbox(t["select_station"], options=list(stations_qc.keys()), key="qc_select")
            qc_table = generate_qc_audit_table(stations_qc[qc_station_choice], qc_rule, t)
            
            show_failed_only = st.checkbox(t["qc_filter_failed"], value=False)
            display_qc_table = qc_table[qc_table[t["qc_col_na"]] > 0] if show_failed_only else qc_table
            
            st.dataframe(display_qc_table, use_container_width=True)
            csv_buffer = display_qc_table.to_csv(index=False).encode('utf-8')
            st.download_button(
                label=t["download_qc_csv"],
                data=csv_buffer,
                file_name=f"Log_Audit_QC_{param_mode_qc}_{qc_station_choice}.csv",
                mime="text/csv"
            )
    else:
        st.info(t["info_upload"])
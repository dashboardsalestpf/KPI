import streamlit as st
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import altair as alt
import json
import time
import io

# ----- Connect to Google Sheet -----
@st.cache_resource
def connect_gsheet(x):
    creds_dict = st.secrets["projectkpidashboard"]  # read secret
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key("1bxVq20N1G9UIyek6BVMgpEMOvQJib-j7Jbsfj82KZtE")
    sheet = spreadsheet.worksheet(x)
    return sheet

# ----- Load Google Sheet into DataFrame -----
def load_data(x):
    sheet = connect_gsheet(x)
    data = sheet.get_all_records()
    df = pd.DataFrame(data)
    return df

@st.cache_data
def loading_data():
    df = load_data("Database")
    df2 = load_data("Database2")
    df3 = load_data("Database3")
    df4 = load_data("Form Responses 1")
    df5 = load_data("Database4")
    df6 = load_data("Form Responses 2")
    df7 = load_data("Database5")
    return df, df2, df3, df4, df5, df6, df7

# ----- Calculate KPI 1 (Alokasi AR) -----
def calculate_kpi_ar(df, Month, Year, User):
    df = df[(df['Month'] == Month) & (df['Year'] == Year) & (df['user'] == User)]
    target = df['DocNum'].count()
    realisasi = df['Poin'].sum()
    percentage = (realisasi / target) * 100 if target != 0 else 0
    poin = 50
    final = poin * (percentage / 100)
    
    return pd.DataFrame({
        "Target": [target],
        "Realisasi": [realisasi],
        "%": [round(percentage, 2)],
        "Poin": [poin],
        "Final": [round(final, 2)]
    }, index=["Alokasi AR Tepat Waktu (Daily) H+1 Tanggal Uang Masuk"])

# ----- Calculate KPI 2 (Cancelled Incoming) -----
def calculate_kpi_cancel(df, Month, Year, User):
    df = df[(df['Month'] == Month) & (df['Year'] == Year) & (df['user'] == User)]
    target = 0
    realisasi = df[df['Canceled'] == "Y"]['DocNum'].count()
    if realisasi <= 20:
        faktor_pengurang = 0
    elif realisasi <= 30:
        faktor_pengurang = -5
    elif realisasi <= 40:
        faktor_pengurang = -10
    elif realisasi <= 50:
        faktor_pengurang = -15
    else:
        faktor_pengurang = -20
    poin = 0
    final = poin + faktor_pengurang
    percentage = 100 - ((faktor_pengurang / -20) * 100)
    
    return pd.DataFrame({
        "Target": [target],
        "Realisasi": [realisasi],
        "%": [round(percentage, 2)],
        "Poin": [poin],
        "Final": [round(final, 2)]
    }, index=["Cancel Incoming (Monthly) Pengurangan Setiap Adanya Cancel Incoming"])

def calculate_kpi_tagih_invoice(df2, Month, Year):
    df2 = df2[(df2['Month'] == Month) & (df2['Year'] == Year)]
    target = df2['Document Number'].count()
    realisasi = df2['Poin'].sum()
    percentage = (realisasi / target) * 100 if target != 0 else 0
    poin = 20
    final = poin * (percentage / 100)
    
    return pd.DataFrame({
        "Target": [target],
        "Realisasi": [realisasi],
        "%": [round(percentage, 2)],
        "Poin": [poin],
        "Final": [round(final, 2)]
    }, index=["Keberhasilan Penagihan (Khusus Tempo) %Invoice Jt >14 Hari Setiap Bulannya"])

def calculate_kpi_closing_bank(df4, Month, Year):
    df4 = df4[(df4['Month'] == Month) & (df4['Year (YYYY)'] == Year)]
    # Convert Timestamp column to datetime
    df4['Timestamp'] = pd.to_datetime(df4['Timestamp'], errors='coerce', format="%m/%d/%Y %H:%M:%S")
    # Target deadline
    target = pd.to_datetime(f"10 {Month} {Year}", format="%d %B %Y")
    target = target + pd.DateOffset(months=1)
    # Take the latest closing timestamp for the selected month
    if df4['Timestamp'].notna().any():
        realisasi = df4['Timestamp'].min()
    else:
        realisasi = "Belum Upload"
    if realisasi != "Belum Upload":
        selisih = (pd.to_datetime(realisasi).date() - pd.to_datetime(target).date()).days
        percentage = 100 if selisih <= 0 else 0
    else:
        selisih = "Belum Upload"
        percentage = 0
    poin = 10
    final = poin * (percentage / 100)
    
    return pd.DataFrame({
        "Target": [target],
        "Realisasi": [realisasi],
        "%": [round(percentage, 2)],
        "Poin": [poin],
        "Final": [round(final, 2)],
    }, index=["Closing Bank (Credit) Tepat Waktu (Monthly) Setiap Tanggal 10 Bulan Berikutnya"])

def calculate_kpi_filing_ke_accounting(df6, Month, Year):
    df6 = df6[(df6['Month'] == Month) & (df6['Year (YYYY)'] == Year)]
    # Convert Timestamp column to datetime
    df6['Timestamp'] = pd.to_datetime(df6['Timestamp'], errors='coerce', format="%m/%d/%Y %H:%M:%S")
    # Target deadline
    target = pd.to_datetime(f"1 {Month} {Year}", format="%d %B %Y")
    target = target - pd.DateOffset(months=1) + pd.offsets.MonthEnd(0)
    # Take the latest closing timestamp for the selected month
    if df6['Timestamp'].notna().any():
        realisasi = df6['Timestamp'].max()
    else:
        realisasi = "Belum Upload"
    if realisasi != "Belum Upload":
        selisih = (pd.to_datetime(realisasi).date() - pd.to_datetime(target).date()).days
        percentage = 100 if selisih <= 0 else 0
    else:
        selisih = "Belum Upload"
        percentage = 0
    poin = 10
    final = poin * (percentage / 100)
    
    return pd.DataFrame({
        "Target": [target],
        "Realisasi": [realisasi],
        "%": [round(percentage, 2)],
        "Poin": [poin],
        "Final": [round(final, 2)],
    }, index=["Serah Terima Dokumen Filing Ke Accounting 1 Bulan Setelah Periode Berjalan"])

def calculate_kpi_performance(df7, Month, Year, User):
    df7 = df7[(df7['Month'] == Month) & (df7['Year'] == Year) & (df7['User'] == User)]
    target = 0
    poin = 10
    realisasi = target - df7["Poin"].count()
    final = poin + realisasi
    percentage = (final / poin) * 100 if final >= 0 else 0
    
    return pd.DataFrame({
        "Target": [target],
        "Realisasi": [realisasi],
        "%": [round(percentage, 2)],
        "Poin": [poin],
        "Final": [round(final, 2)]
    }, index=["Pelanggaran Sop Kerja Poin Pengurangan Nilai Kpi Setiap Pelanggaran Yang Timbul"])

def calculate_total_kpi(kpi1, kpi2, kpi3, kpi4, kpi5, kpi6, Month, Year):
    return pd.DataFrame({
        "Poin": [kpi1["Poin"].iloc[0] + kpi2["Poin"].iloc[0] + kpi3["Poin"].iloc[0] + kpi4["Poin"].iloc[0] + kpi5["Poin"].iloc[0] + kpi6["Poin"].iloc[0]],
        "Final": [kpi1["Final"].iloc[0] + kpi2["Final"].iloc[0] + kpi3["Final"].iloc[0] + kpi4["Final"].iloc[0] + kpi5["Final"].iloc[0] + kpi6["Final"].iloc[0]]
    }, index=[f"TOTAL KPI {Month} {Year}"])

# ----- Streamlit App -----
def main_app():
    # Set wide layout
    st.set_page_config(page_title="KPI Indicator", layout="wide")
    st.title("📊 KPI Indicator")
    if st.button("Refresh Data"):
        time.sleep(1)
        st.cache_data.clear()
        df, df2, df3, df4, df5, df6, df7 = loading_data()
        st.rerun()
    df, df2, df3, df4, df5, df6, df7 = loading_data()
    Year = st.selectbox("Select a year", df['Year'].sort_values().unique())
    months = df[df['Year'] == Year]['Month'].unique()
    month_order = ["January", "February", "March", "April", "May", "June","July", "August", "September", "October", "November", "December"]
    months = [m for m in month_order if m in months]  # keeps calendar order
    Month = st.selectbox("Select a month", months)
    # % progress bars
    st.subheader("KPI Indicator per User")
    user = df[(df['Year'] == Year) & (df['Month'] == Month)]['user'].unique()
    User = st.selectbox("Select a user", user)

    # Calculate KPIs
    kpi1 = calculate_kpi_ar(df, Month, Year, User)
    kpi2 = calculate_kpi_cancel(df, Month, Year, User)
    kpi3 = calculate_kpi_tagih_invoice(df2, Month, Year)
    kpi4 = calculate_kpi_closing_bank(df4, Month, Year)
    kpi5 = calculate_kpi_filing_ke_accounting(df6, Month, Year)
    kpi6 = calculate_kpi_performance(df7, Month, Year, User)
    total = calculate_total_kpi(kpi1, kpi2, kpi3, kpi4, kpi5, kpi6, Month, Year)
    kpi_table = pd.concat([kpi1, kpi2, kpi3, kpi4, kpi5, kpi6, total])

    # Display KPIs
    kpi_bar = pd.concat([kpi1, kpi2, kpi6])
    for idx, row in kpi_bar.iterrows():
        col1, col2, col3 = st.columns([5, 5, 2])  # adjust column widths
        with col1:
            st.write(f"**{idx}**")
        with col2:
            progress_value = int(min(max(row['%'], 0), 100))  # cap between 0-100
            st.progress(progress_value)
        with col3:
            st.write(f"Final: {row['Final']}")
    st.subheader("KPI Indicator by Division")
    kpi_bar2 = pd.concat([kpi3, kpi4, kpi5])
    for idx, row in kpi_bar2.iterrows():
        col1, col2, col3 = st.columns([5, 5, 2])  # adjust column widths
        with col1:
            st.write(f"**{idx}**")
        with col2:
            progress_value = int(min(max(row['%'], 0), 100))  # cap between 0-100
            st.progress(progress_value)
        with col3:
            st.write(f"Final: {row['Final']}")
    # Show combined KPI table
    st.dataframe(kpi_table)
    # Create an in-memory output file for Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        kpi_table.to_excel(writer, index=True, sheet_name="KPI")
    # Rewind the buffer
    output.seek(0)
    # Download button
    st.download_button(
        label="Download KPI Table (Excel)",
        data=output,
        file_name=f"KPI_{Month}_{Year}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

main_app()
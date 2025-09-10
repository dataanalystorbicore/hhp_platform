import streamlit as st
import pandas as pd
import gspread
from prepare_visit_report import cleaning
from upload_done_calls import write_dataframe_to_gsheet
from Database_update import update_database
from io import BytesIO
import openpyxl



# Page configuration
st.set_page_config(page_title="Visit Manager", layout="wide")
service_account_info = st.secrets["gcp_service_account"]
gc = gspread.service_account_from_dict(service_account_info)
# Title
st.title("📍 Data Controller Dashboard")

# Spacer
st.write("")
st.write("")


def download_attendance():
     
    attendance = gc.open("attendancereport")
    attendance_worksheet = attendance.sheet1
    data = attendance_worksheet.get_all_values()
    data_attendance = pd.DataFrame(data[1:], columns=['Date','BBE CODE','BBE NAME','Wilaya','Attendance'])

    return data_attendance


# Two columns for the two buttons
col1, col2 , col3 , col4 = st.columns(4)


with col1:
    st.subheader("Prepare List of Calls")
    visits_report = st.file_uploader("Upload visits Report", type=["xlsx"])

    if(visits_report):
        
        cleaning(visits_report)

with col2:

    st.subheader("Upload Done Calls")
    done_calls = st.file_uploader("Upload Done Calls", type=["xlsx"])

    if(done_calls):
         
        excel_data = pd.read_excel(done_calls)
        
        write_dataframe_to_gsheet(excel_data)


with col3:

    st.subheader("Daily Attendance")


    if st.button("Fill Attendance"):
        st.switch_page("pages/attendance.py")
        
    df = download_attendance()

    tmp_buffer = BytesIO()
    with pd.ExcelWriter(tmp_buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    tmp_buffer.seek(0)


    wb = openpyxl.load_workbook(tmp_buffer)
    ws = wb.active

    final_buffer = BytesIO()
    wb.save(final_buffer)
    final_buffer.seek(0)

    st.download_button(
        label="📥 Download Attendance Report",
        data=final_buffer,
        file_name="Attendance_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


with col4:

    st.subheader("Update Database")
    pos_report = st.file_uploader("Upload POS Extraction", type=["xlsx"])

    if(pos_report):
         
        update_database(pos_report)
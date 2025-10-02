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

def donwnload_database():
    database = gc.open("database")
    database_worksheet = database.sheet1
    data = database_worksheet.get_all_values()
    data_pos = pd.DataFrame(data[1:], columns=data[0])

    return data_pos


def download_bbeinfo():
    bbe_info = gc.open("bbe_info")
    bbe_worksheet = bbe_info.sheet1
    bbe_values = bbe_worksheet.get_all_values()
    df_bbe = pd.DataFrame(bbe_values[1:], columns=bbe_values[0])

    return df_bbe


def download_attendance():
     
    attendance = gc.open("attendancereport")
    attendance_worksheet = attendance.sheet1
    data = attendance_worksheet.get_all_values()
    data_attendance = pd.DataFrame(data[1:], columns=['Date','BBE CODE','BBE NAME','Wilaya','Attendance'])

    return data_attendance


# Two columns for the two buttons
col1, col2 , col3 , col4, col5 = st.columns(5)


with col1:
    st.subheader("Prepare List of Calls")
    BBE = st.selectbox("Select BBE", download_bbeinfo()["NAME"].unique(),index=None)
    visits_report = st.file_uploader("Upload visits Report", type=["xlsx"])

    if(visits_report):


        if(BBE != None):

            BBE_CODE = (download_bbeinfo().loc[download_bbeinfo()["NAME"] == BBE, "BBE_CODE"].iloc[0])
            cleaning(visits_report, BBE_CODE)
        
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
         
        new_database = update_database(pos_report)

        tmp_buffer = BytesIO()
        with pd.ExcelWriter(tmp_buffer, engine="openpyxl") as writer:
            new_database.to_excel(writer, index=False, sheet_name="Sheet1")
        tmp_buffer.seek(0)


        wb = openpyxl.load_workbook(tmp_buffer)
        ws = wb.active

        final_buffer = BytesIO()
        wb.save(final_buffer)
        final_buffer.seek(0)

        st.download_button(
            label="📥 Download Updated Database",
            data=final_buffer,
            file_name="Base De Données.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:

        new_database = donwnload_database()

        tmp_buffer = BytesIO()
        with pd.ExcelWriter(tmp_buffer, engine="openpyxl") as writer:
            new_database.to_excel(writer, index=False, sheet_name="Sheet1")
        tmp_buffer.seek(0)


        wb = openpyxl.load_workbook(tmp_buffer)
        ws = wb.active

        final_buffer = BytesIO()
        wb.save(final_buffer)
        final_buffer.seek(0)

        st.download_button(
            label="📥 Download Updated Database",
            data=final_buffer,
            file_name="Base De Données.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )



with col5:

    st.subheader("Edit BBE INFO")


    if st.button("GO to Page"):
        st.switch_page("pages/bbe_info.py")

    df_bbe = download_bbeinfo()
    
    tmp_buffer_bbe = BytesIO()
    with pd.ExcelWriter(tmp_buffer_bbe, engine="openpyxl") as writer:
        df_bbe.to_excel(writer, index=False, sheet_name="Sheet1")
        tmp_buffer.seek(0)


    wb = openpyxl.load_workbook(tmp_buffer_bbe)
    ws = wb.active

    final_buffer_bbe = BytesIO()
    wb.save(final_buffer_bbe)
    final_buffer_bbe.seek(0)

    st.download_button(
        label="📥 Download BBE INFO",
        data=final_buffer_bbe,
        file_name="BBE INFO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ) 



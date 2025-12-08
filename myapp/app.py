import streamlit as st
import pandas as pd
import gspread
from prepare_visit_report import cleaning
from upload_done_calls import write_dataframe_to_gsheet
from Database_update import update_database
from io import BytesIO
import openpyxl
from openpyxl.styles import Font
from difflib import SequenceMatcher



service_account_info = st.secrets["gcp_service_account"]
gc = gspread.service_account_from_dict(service_account_info)


# Page configuration
st.set_page_config(page_title="Visit Manager", layout="wide")
st.title("📍 Daily Tasks Preparator")

# Spacer
st.write("")
st.write("")


def similarity(a, b):
    return SequenceMatcher(None, a, b).ratio()


def clean_with_sequence_matcher(series, target, threshold=0.8):
    return series.apply(
        lambda x: target if similarity(str(x).lower(), target.lower()) >= threshold else x
    )



def prepare_visits_report(df):

    df_bbe = download_bbeinfo()

    df["REMARK"] = df["REMARK"].astype(str).str.lower()

    df["REMARK"] = clean_with_sequence_matcher(df["REMARK"], "wrong number", threshold=0.8)

    df["REMARK"] = clean_with_sequence_matcher(df["REMARK"], "ok", threshold=0.8)

    df["REMARK"] = df["REMARK"].replace('ok','call completed')

    allowed = ["call completed", "wrong number"]

    df.loc[~df["REMARK"].str.lower().isin(allowed), "REMARK"] = "no answer"

    df["REMARK"] = df["REMARK"].astype(str).str.upper()

    df_bbe.rename(columns={'NAME': 'BBE_NAME'}, inplace=True)

    df = df.merge(df_bbe[['BBE (ORB)','BBE_NAME']],on = 'BBE (ORB)', how = 'left')

    df = df[['DATE','Region','BBE (ORB)', 'BBE_NAME','Wilaya', 'Commune', 'Pos Type',
       'Name','Nom_complet_proprio', 'PropriétairePhone',
       'Nom_complet_gerant', 'GérantPhone', 'POS Adress',
       'Merchandiser visit', 'Sales request for the week', 'POP S25',
       'Relationship with merchandiser', 'REMARK']]

    return df


def download_calendar():
    Calendrier = gc.open("Official_Calendar")
    Calendrier_worksheet = Calendrier.sheet1
    Calendrier_values = Calendrier_worksheet.get_all_values()
    df_Calendrier = pd.DataFrame(Calendrier_values[1:], columns=Calendrier_values[0])

    return df_Calendrier


def download_database():
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


def download_visitsreport():
     
    visits = gc.open("visitsreport")
    visits_worksheet = visits.sheet1
    data = visits_worksheet.get_all_values()
    data_visits = pd.DataFrame(data[1:], columns=[
                    'Region','BBE (ORB)','Wilaya','Commune','Pos Type',
                    'Site ID','Name','Nom_complet_proprio','PropriétairePhone',
                    'Nom_complet_gerant','GérantPhone','POS Adress','DATE',
                    'Merchandiser visit','Sales request for the week','POP S25',
                    'Relationship with merchandiser','REMARK'])

    return data_visits


def download_function(df,label,key,filename):

    tmp_buffer = BytesIO()
    with pd.ExcelWriter(tmp_buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    tmp_buffer.seek(0)


    wb = openpyxl.load_workbook(tmp_buffer)
    ws = wb.active
    font = Font(name='Calibri', size=12)

    for row in ws.iter_rows():
        for cell in row:
            cell.font = font

    final_buffer = BytesIO()
    wb.save(final_buffer)
    final_buffer.seek(0)

    st.download_button(
        label=label,
        data=final_buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key = key
    )


with st.container():

    # Two columns for the two buttons
    col1, col2 , col3 , col4, col5= st.columns(5)


    with col1:
        st.subheader("Prepare List of Calls")
        visits_report = st.file_uploader("Upload visits Report", type=["xlsx"])
        BBE = st.selectbox("Select BBE", download_bbeinfo()["NAME"].unique(),index=None)

        if(visits_report):


            if(BBE != None):

                BBE_CODE = (download_bbeinfo().loc[download_bbeinfo()["NAME"] == BBE, "BBE_CODE"].iloc[0])
                cleaning(visits_report, BBE_CODE)
            
            cleaning(visits_report)


    with col2:

        st.subheader("Upload Done Calls")
        done_calls = st.file_uploader("Upload Done Calls", type=["xlsx"])
        df = download_visitsreport()
        df_calendar = download_calendar()

        df["DATE"] = df["DATE"].astype("str")

        df_calendar["DATE"] = pd.to_datetime(df_calendar["DATE"], errors="coerce").dt.strftime("%Y-%m-%d")
        
        df = df.merge(df_calendar[['DATE','Week']], on = 'DATE',how = 'left')
        Week = st.selectbox("Select Week", df["Week"].unique(),index=None)
        df = df[df["Week"] == Week]

        df = prepare_visits_report(df)


        if(done_calls):
            
            excel_data = pd.read_excel(done_calls)
            
            write_dataframe_to_gsheet(excel_data)


        label="📥 Download Visits Report"

        key = "Download_Visits_report"

        file_name="Visits_report.xlsx"

        download_function(df,label,key,file_name)


    with col3:

        st.subheader("Update Database")
        pos_report = st.file_uploader("Upload POS Extraction", type=["xlsx"])

        if(pos_report):
            
            new_database = update_database(pos_report)

            label="📥 Download Updated Database"

            key = "Download_database"

            file_name="Database.xlsx"

            download_function(new_database,label,key,file_name)


        else:

            new_database = download_database()

            label="📥 Download Updated Database"

            key = "Download_database_2"

            file_name="Database.xlsx"

            download_function(new_database,label,key,file_name)


    with col4:

        st.subheader("Daily Attendance")


        if st.button("Fill Attendance"):
            st.switch_page("pages/attendance.py")
            
        df = download_attendance()

        label="📥 Download Attendance Report"

        key = "Download_daily_attendance"

        file_name="Attendance_report.xlsx"

        download_function(df,label,key,file_name)



    with col5:

        st.subheader("Edit BBE INFO")


        if st.button("GO to Page"):
            st.switch_page("pages/bbe_info.py")

        df_bbe = download_bbeinfo()

        label="📥 Download BBE INFO"

        key = "Download_bbe_info"

        file_name="BBE_INFO.xlsx"

        download_function(df_bbe,label,key,file_name)

        



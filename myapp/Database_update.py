import pandas as pd
import gspread
import streamlit as st

service_account_info = st.secrets["gcp_service_account"]
gc = gspread.service_account_from_dict(service_account_info)

database = gc.open("database")
database_worksheet = database.sheet1
data = database_worksheet.get_all_values()
data_pos = pd.DataFrame(data[1:], columns=data[0])

def update_database(pos_report):


    df = pd.read_excel(pos_report)

    st.write(df)


    return
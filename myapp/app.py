import streamlit as st
import pandas as pd
import gspread
from prepare_visit_report import cleaning
from upload_done_calls import write_dataframe_to_gsheet


# Page configuration
st.set_page_config(page_title="Visit Manager", layout="wide")
service_account_info = st.secrets["gcp_service_account"]
gc = gspread.service_account_from_dict(service_account_info)
# Title
st.title("📍 Site Visit Dashboard")

# Spacer
st.write("")
st.write("")

# Two columns for the two buttons
col1, col2 = st.columns(2)


with col1:
    st.subheader("Prepared List of Calls")
    visits_report = st.file_uploader("Upload visits Report", type=["xlsx"])

    if(visits_report):
        
        cleaning(visits_report)

with col2:

    st.subheader("Upload Done Calls")
    done_calls = st.file_uploader("Upload Done Calls", type=["xlsx"])

    if(done_calls):
         
        excel_data = pd.read_excel(done_calls)
        
        write_dataframe_to_gsheet(excel_data)


        

import streamlit as st
import gspread

service_account_info = st.secrets["gcp_service_account"]
gc = gspread.service_account_from_dict(service_account_info)
sh = gc.open("visitsreport")
worksheet = sh.sheet1

def write_dataframe_to_gsheet(df):
    
    try:
        sh = gc.open("visitsreport")
        worksheet = sh.sheet1
        # Convert DataFrame to a list of lists, excluding the header
        df = df.fillna(" ")
        data_to_append = df.values.tolist()
        
        # Append the new data to the end of the worksheet
        worksheet.append_rows(data_to_append, value_input_option='USER_ENTERED')
        st.success(f"Data successfully written to Google Sheet '{sh}' in worksheet '{worksheet}'.")
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"Worksheet '{worksheet}' not found.")
    except Exception as e:
        st.error(f"An error occurred: {e}")

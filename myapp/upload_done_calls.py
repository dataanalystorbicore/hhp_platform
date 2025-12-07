import streamlit as st
import gspread
import pandas as pd
from datetime import datetime


service_account_info = st.secrets["gcp_service_account"]
gc = gspread.service_account_from_dict(service_account_info)


sh = gc.open("visitsreport")
worksheet = sh.sheet1

def download_visitreport():
    visits_values = worksheet.get_all_values()
    df_visits = pd.DataFrame(visits_values[1:], columns=visits_values[0])

    return df_visits

def write_dataframe_to_gsheet(df):

    required_columns = [
    'Region','BBE (ORB)','Wilaya','Commune','Pos Type',
    'Site ID','Name','Nom_complet_proprio','PropriétairePhone',
    'Nom_complet_gerant','GérantPhone','POS Adress','DATE',
    'Merchandiser visit','Sales request for the week','POP S25',
    'Relationship with merchandiser','REMARK'
    ]


    df["DATE"] = pd.to_datetime(
        df["DATE"], 
        format="%d-%m", 
        errors="coerce"
    ).apply(lambda d: d.replace(year=datetime.now().year) if pd.notnull(d) else d)

    df["DATE"] = df["DATE"].astype(str)


    missing_cols = [col for col in required_columns if col not in df.columns]

    unique_dates = pd.Series(df["DATE"].unique())

    if missing_cols:
        st.error(f"Some columns are missing: {missing_cols}")

    elif unique_dates.isin(download_visitreport()["DATE"].unique()).any():
        st.error("The visits you're trying to upload already exist")

    else:

        try:
            sh = gc.open("visitsreport")
            worksheet = sh.sheet1
            # Convert DataFrame to a list of lists, excluding the header
            df = df.fillna(" ")
            data_to_append = df.values.tolist()
            
            # Append the new data to the end of the worksheet
            worksheet.append_rows(data_to_append, value_input_option='USER_ENTERED')
            st.success(f"Data successfully Added in google Sheet Visitsreport")
        except gspread.exceptions.WorksheetNotFound:
            st.error(f"Worksheet '{worksheet}' not found.")
        except Exception as e:
            st.error(f"An error occurred: {e}")
import streamlit as st
import pandas as pd
import gspread
import openpyxl
from openpyxl.styles import PatternFill, Font
from io import BytesIO


# --- Google Sheets setup
service_account_info = st.secrets["gcp_service_account"]
gc = gspread.service_account_from_dict(service_account_info)

# reads old visits report
visitreport = gc.open("visitsreport")
visitreport_worksheet = visitreport.sheet1


# reads database
database = gc.open("database")
database_worksheet = database.sheet1


def download_bbeinfo():
    bbe_info = gc.open("bbe_info")
    bbe_worksheet = bbe_info.sheet1
    bbe_values = bbe_worksheet.get_all_values()
    df_bbe = pd.DataFrame(bbe_values[1:], columns=bbe_values[0])

    return df_bbe


# sample rows for wilayas
def select_and_sample_rows_for_wilayas(df, BBE = None):

    if BBE != None:
        if (BBE in (df['BBE_CODE'].unique())):
            df_filtered = df[df['BBE_CODE'] == BBE]
        else:
            st.write("BBE didn't work yesterday")
    else:
        st.write('No BBE is selected')

    wilaya_rows = df.drop_duplicates(subset=['Wilaya'])
    target_rows = 50
    num_to_sample = target_rows - len(wilaya_rows)

    if num_to_sample <= 0:
        return wilaya_rows.reset_index(drop=True)

    remaining_rows = df[~df.index.isin(wilaya_rows.index)]

    if len(remaining_rows) < num_to_sample:
        sampled_rows = remaining_rows
    else:
        sampled_rows = remaining_rows.sample(n=num_to_sample, replace=False, random_state=1)

    if "df_filtered" in locals() or "df_filtered" in globals():

        final_df = pd.concat([wilaya_rows, sampled_rows, df_filtered]).reset_index(drop=True)
    else:
        final_df = pd.concat([wilaya_rows, sampled_rows]).reset_index(drop=True)

    final_df = final_df.sample(frac=1).reset_index(drop=True)

    return final_df



def cleaning(uploaded_file, BBE_CODE = None):

    if uploaded_file:
        try:
            st.write("File uploaded, processing...")

            df_bbe = download_bbeinfo()

            # --- Load visits
            data_visits = pd.read_excel(uploaded_file)

            data = database_worksheet.get_all_values()
            data_pos = pd.DataFrame(data[1:], columns=data[0])

            data_pos = data_pos[~(data_pos["PropriétairePhone"] == " ")]

            if (data_visits['Username'] == 'test').any():
                data_visits = data_visits[data_visits['Username'] != 'test']

            if (data_visits['Closed'] == 'YES').any():
                data_visits = data_visits[data_visits['Closed'] != 'YES']

            data_visits = data_visits.merge(df_bbe[['BBE_CODE', 'Username']], on = 'Username', how = 'left')

            data_visits["Pos id"] = (
                data_visits["Region"].astype(str) +
                data_visits["District"].astype(str).str.zfill(2) +
                data_visits["Territory"].astype(str) +
                data_visits["Code"].astype(str).str.zfill(5)
            ).astype(int)

            if data_visits['Site ID'].isnull().any():
                lookup = dict(zip(data_pos['Pos id'].astype(int), data_pos['Site ID']))
                data_visits['Site ID'] = data_visits['Site ID'].fillna(data_visits['Pos id'].map(lookup))

            data_pos["Nom_complet_proprio"] = data_pos["PropriétaireLastname"] + "_" + data_pos["PropriétaireFirstname"]
            data_pos["Nom_complet_gerant"] = data_pos["GérantLastname"] + "_" + data_pos["GérantFirstname"]

            data_visits = data_visits.merge(
                data_pos[['Site ID','PropriétairePhone','GérantPhone',
                          'Nom_complet_proprio','Nom_complet_gerant',
                          'Pos Type','Wilaya(48)']],
                on='Site ID', how='left'
            )

            data_visits = data_visits.drop_duplicates(subset=['Site ID'])
            data_visits['BBE_CODE'] = data_visits['BBE_CODE'].astype(str)
            data_visits = data_visits[['Region','BBE_CODE','Wilaya','Wilaya(48)',
                                       'Commune','Pos Type','Site ID','Name',
                                       'Nom_complet_proprio','PropriétairePhone',
                                       'Nom_complet_gerant','GérantPhone','Address']]

            for col in ['DATE','Merchandiser visit','Sales request for the week','POP S25',
                'Relationship with merchandiser','REMARK']:
                data_visits[col] = None     

            # Exclude already saved sites
            report_data = visitreport_worksheet.get_all_values()
            done_calls = pd.DataFrame(report_data[1:],columns=[
                'Region','BBE_CODE','Wilaya','Wilaya(48)','Commune','Pos Type',
                'Site ID','Name','Nom_complet_proprio','PropriétairePhone',
                'Nom_complet_gerant','GérantPhone','Address','DATE',
                'Merchandiser visit','Sales request for the week','POP S25',
                'Relationship with merchandiser','REMARK'])

            data_visits = data_visits.rename(columns = {"Address" : "POS Adress"})


            if not done_calls.empty:
                saved_sites = done_calls['Site ID'].unique()
                data_visits = data_visits[~data_visits['Site ID'].isin(saved_sites)]

            data_visits = data_visits.dropna(axis = 0, subset = ["PropriétairePhone","GérantPhone"])

            data_visits = select_and_sample_rows_for_wilayas(data_visits,BBE_CODE)

            st.success("New Cleaned visits were added successfully")

            # --- STEP 1: Write DataFrame to temp buffer
            tmp_buffer = BytesIO()
            with pd.ExcelWriter(tmp_buffer, engine="openpyxl") as writer:
                data_visits.to_excel(writer, index=False, sheet_name="Sheet1")
            tmp_buffer.seek(0)

            # --- STEP 2: Open with openpyxl
            wb = openpyxl.load_workbook(tmp_buffer)
            ws = wb.active

            green_fill = PatternFill(start_color="FF00B0F0", end_color="FF00B0F0", fill_type="solid")
            blue_fill = PatternFill(start_color="FF757171", end_color="FF757171", fill_type="solid")
            red_fill = PatternFill(start_color="FF7030A0", end_color="FF7030A0", fill_type="solid")
            white_bold_font = Font(color="FFFFFFFF", bold=True)
            font = Font(name='Calibri', size=12)

            for col_index, cell in enumerate(ws[1], 1):
                cell.font = white_bold_font
                if col_index <= 6:
                    cell.fill = green_fill
                elif 7 <= col_index <= 12:
                    cell.fill = blue_fill
                elif 13 <= col_index <= 20:
                    cell.fill = red_fill

            for row in ws.iter_rows(min_row=2):  # start from second row
                for cell in row:
                    cell.font = font


            # --- STEP 3: Save to NEW buffer
            final_buffer = BytesIO()
            wb.save(final_buffer)
            final_buffer.seek(0)

            # --- STEP 4: Download
            st.download_button(
                label="📥 Download Call Report",
                data=final_buffer,
                file_name="Phone Calls.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Failed to process file: {e}")

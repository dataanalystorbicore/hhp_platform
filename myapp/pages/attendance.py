import streamlit as st
import pandas as pd
from datetime import date
import gspread

service_account_info = st.secrets["gcp_service_account"]
gc = gspread.service_account_from_dict(service_account_info)
sh = gc.open("attendancereport")
worksheet = sh.sheet1


st.title(f"🗓️ Daily Presence for {date.today()}")

st.title("")



people = {
    "ORB002#R#44": {"wilaya": "AÏN DEFLA", "name": "HADJ SEDOUK FETHALLAH"},
    "ORB003#R#16": {"wilaya": "ALGER", "name": "WALID MENACER"},
    "ORB004#R#16": {"wilaya": "ALGER", "name": "LAICHI AMINE ABDELKADER"},
    "ORB009#R#16": {"wilaya": "ALGER", "name": "YAKOUB BOURENANI"},
    "ORB007#R#16": {"wilaya": "ALGER", "name": "BOURENANI KHALED"},
    "ORB036#WS#16": {"wilaya": "ALGER", "name": "SALMI ABDELGHANI"},
    "ORB043#WS#16": {"wilaya": "ALGER", "name": "ZAKARIA BOUHAFS"},
    "ORB039#R#23": {"wilaya": "ANNABA", "name": "HOUHOU AHLEM"},
    "ORB011#R#05": {"wilaya": "BATNA", "name": "MERBAI YASSER"},
    "ORB060#R#05": {"wilaya": "BATNA", "name": "BILAL SGHIRI"},
    "ORB013#R#06": {"wilaya": "BÉJAÏA", "name": "DJINNI REDA"},
    "ORB042#R#09": {"wilaya": "BLIDA", "name": "BENDALI BILEL"},
    "ORB015#R#34": {"wilaya": "BORDJ BOU ARRERIDJ", "name": "ABDELLAH TLIDJAN"},
    "ORB016#R#35": {"wilaya": "BOUMERDÈS", "name": "BELHADJOURI MEHDI"},
    "ORB017#R#02": {"wilaya": "CHLEF", "name": "FELLAH YOUCEF"},
    "ORB018#R#25": {"wilaya": "CONSTANTINE", "name": "KAHOUL SAMI"},
    "ORB019#R#25": {"wilaya": "CONSTANTINE", "name": "LAOUAR NAZIM"},
    "ORB020#R#17": {"wilaya": "DJELFA", "name": "TITTOUH LAMINE"},
    "ORB022#R#29": {"wilaya": "MASCARA", "name": "NADIR"},
    "ORB024#R#27": {"wilaya": "MOSTAGANEM", "name": "BENABOU AHMED"},
    "ORB025#R#31": {"wilaya": "ORAN", "name": "HADFI WANIS"},
    "ORB038#WS#31": {"wilaya": "ORAN", "name": "KHEDIM MAHDJOUBA"},
    "ORB037#WS#19": {"wilaya": "SÉTIF (el eulma )", "name": "TAHAR ROGAI"},
    "ORB030#R#12": {"wilaya": "TÉBESSA", "name": "MERABTI YACINE"},
    "ORB032#R#42": {"wilaya": "TIPAZA", "name": "NOUFEL LAZZOULI"},
    "ORB041#R#15": {"wilaya": "TIZI OUZOU", "name": "DAHMOUN RACHID"},
    "ORB034#R#13": {"wilaya": "TLEMCEN", "name": "BERRAHMOUNE SAMIR"}
}

attendance_data = []

for key, value in people.items():

        col1, col2, col3 = st.columns([1, 1, 2])

        with col1:
            st.write(value["name"])

        with col2:
            st.write(value["wilaya"])

        with col3:
            status = st.selectbox(
                "Attendance",
                ["Present", "Absent", "Congé"],
                index=0,
                key=f"attendance_{key}"  # unique key for each row
            )


        attendance_data.append({
        "date": date.today().isoformat(),
        "code": key,
        "wilaya": value["wilaya"],
        "name": value["name"],
        "status": status})
        
    
        st.markdown("<hr style='border:1px solid #ddd;'>", unsafe_allow_html=True)



df = pd.DataFrame(attendance_data)

if st.button("Save Attendance"):
    
    try:
        sh = gc.open("attendancereport")
        worksheet = sh.sheet1
        data_to_append = df.values.tolist()
        
        worksheet.append_rows(data_to_append, value_input_option='USER_ENTERED')
        st.success(f"Data successfully written to Google Sheet '{sh}' in worksheet '{worksheet}'.")
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"Worksheet '{worksheet}' not found.")
    except Exception as e:
        st.error(f"An error occurred: {e}")




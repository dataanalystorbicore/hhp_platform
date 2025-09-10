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
    "ORB032#R#42": {"wilaya": "TIPAZA", "name": "NASSIM"},
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
        "name": value["name"],
        "wilaya": value["wilaya"],
        "status": status})
        
    
        st.markdown("<hr style='border:1px solid #ddd;'>", unsafe_allow_html=True)



df = pd.DataFrame(attendance_data)

if st.button("Save Attendance"):
    try:
        sh = gc.open("attendancereport")
        worksheet = sh.sheet1

        # Get all rows (list of lists)
        all_values = worksheet.get_all_values()

        # Build dataframe manually (if not empty)
        if all_values:
            df_existing = pd.DataFrame(all_values, columns=["date", "code", "name", "wilaya", "status"])
        else:
            df_existing = pd.DataFrame(columns=["date", "code", "name", "wilaya", "status"])

        # Build unique keys for both
        df_existing["unique_key"] = df_existing["date"].astype(str) + df_existing["code"].astype(str)
        df["unique_key"] = df["date"].astype(str) + df["code"].astype(str)

        for _, row in df.iterrows():
            if row["unique_key"] in df_existing["unique_key"].values:
                # Find row index in sheet (1-indexed because no header row!)
                idx = df_existing.index[df_existing["unique_key"] == row["unique_key"]][0] + 1
                worksheet.update(
                    f"A{idx}:E{idx}",
                    [[row["date"], row["code"], row["name"], row["wilaya"], row["status"]]]
                )
            else:
                # Append new row
                worksheet.append_row(
                    [row["date"], row["code"], row["name"], row["wilaya"], row["status"]],
                    value_input_option="USER_ENTERED"
                )

        st.success("✅ Attendance saved without duplicates!")

    except Exception as e:
        st.error(f"An error occurred: {e}")


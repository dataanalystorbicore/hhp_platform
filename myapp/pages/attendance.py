import streamlit as st
import pandas as pd
from datetime import date
import gspread

service_account_info = st.secrets["gcp_service_account"]
gc = gspread.service_account_from_dict(service_account_info)

sh = gc.open("attendancereport")
worksheet = sh.sheet1


# reading bbe_info informations
bbe_info = gc.open("bbe_info")
bbe_worksheet = bbe_info.sheet1
bbe_values = bbe_worksheet.get_all_values()
df_bbe = pd.DataFrame(bbe_values[1:], columns=bbe_values[0])



st.title(f"🗓️ Daily Presence for {date.today()}")

st.title("")



attendance_data = []

for idx, row in df_bbe.iterrows():
    col1, col2, col3 = st.columns([1, 1, 2])

    with col1:
        st.markdown(f"<p style='margin:0'>{row['NAME']}</p>", unsafe_allow_html=True)

    with col2:
        st.markdown(f"<p style='margin:0'>{row['WILAYA']}</p>", unsafe_allow_html=True)

    with col3:
        status = st.selectbox(
            "Attendance",
            ["Present", "Absent", "Congé"],
            index=0,
            key=f"attendance_{idx}"
        )

    attendance_data.append({
        "date": date.today().isoformat(),
        "code": idx,
        "name": row["NAME"],
        "wilaya": row["WILAYA"],
        "status": status
    })

    # make the divider thinner
    st.markdown("<hr style='border:0.5px solid #ccc; margin:5px 0'>", unsafe_allow_html=True)


df = pd.DataFrame(attendance_data)

if st.button("Save Attendance"):
    try:
        
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


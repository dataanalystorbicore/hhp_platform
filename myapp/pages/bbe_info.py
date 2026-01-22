import streamlit as st
import pandas as pd
from datetime import date
import gspread


service_account_info = st.secrets["gcp_service_account"]
gc = gspread.service_account_from_dict(service_account_info)

# reading bbe_info informations
bbe_info = gc.open("bbe_info")
bbe_worksheet = bbe_info.sheet1
bbe_values = bbe_worksheet.get_all_values()
df_bbe = pd.DataFrame(bbe_values[1:], columns=bbe_values[0])


@st.dialog("Edit your BBE")
def BBE(idx,name,phone,wilaya):
    
    Name = st.text_input("NAME",f"{name}")
    Phone = st.text_input("PHONE",f"{phone}")
    Wilaya = st.text_input("WILAYA",f"{wilaya}")
    sheet_row = idx + 2
    
    if st.button("Submit"):
        
        bbe_worksheet.update(f"D{sheet_row}:F{sheet_row}",[[Wilaya,Name,Phone]])

        st.success(f"✅ BBE INFO of {Name} UPDATED Succesfully")
        
        st.rerun()

        


st.title(f" BBE INFO EDIT PAGE")

st.title("")

for idx, row in df_bbe.iterrows():
        
        with st.container():

            col1, col2, col3,col4,col5,col6,col7 = st.columns([1, 1 , 3 , 2, 3 , 2, 2])

            with col1:
                st.write(row["REGION"])

            with col2:
                st.write(row["WILAYA_ID"])

            with col3:
                st.write(row["WILAYA"])

            with col4:
                st.write(row["BBE_CODE"])

            with col5:
                st.write(row["NAME"])

            with col6:
                st.write(row["PHONE"])

            with col7:

                if st.button("Edit",key=f"edit_button_{idx}"):
                    BBE(idx,f"{row['NAME']}",f"{row['PHONE']}",f"{row['WILAYA']}")
        
            st.markdown("<hr style='border:1px solid #ddd;'>", unsafe_allow_html=True)



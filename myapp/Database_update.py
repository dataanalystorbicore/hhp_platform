import pandas as pd
import gspread
import streamlit as st

service_account_info = st.secrets["gcp_service_account"]
gc = gspread.service_account_from_dict(service_account_info)

# read the old database
database = gc.open("database")
database_worksheet = database.sheet1
data = database_worksheet.get_all_values()
data_pos = pd.DataFrame(data[1:], columns=data[0])


def update_database(pos_report):


    # the first sheet of the uploaded file db
    df = pd.read_excel(pos_report)
    
    

    df = df[["System ID","Create Time","Update Time","Delete Time","Last Visit","Region","District","Territory","Code","Site ID",
             "Name","Wilaya","Commune","Address","Latitude","Longitude","Photo"]]
    
    # the pos id for the first sheet
    
    df['Pos id'] = (
        df["Region"].fillna(0).astype(int).astype(str) +
        df["District"].fillna(0).astype(int).astype(str).str.zfill(2) +
        df["Territory"].fillna(0).astype(int).astype(str) +
        df["Code"].fillna(0).astype(int).astype(str).str.zfill(5)).astype(str)

    # reading the second sheet 
    df2 = pd.read_excel(pos_report, sheet_name = "Contacts")

    # creating the pos id for the second sheet
    df2['Pos id'] = (
        df2["Region"].fillna(0).astype(int).astype(str) +
        df2["District"].fillna(0).astype(int).astype(str).str.zfill(2) +
        df2["Territory"].fillna(0).astype(int).astype(str) +
        df2["Code"].fillna(0).astype(int).astype(str).str.zfill(5)).astype(str)

    

    df2 = df2[["Pos id","Type","Firstname","Lastname","Phone"]]

    # separating values like gerant, proprietaire, vendeur
    df2["Type"] = df2["Type"].str.split(", ")
    df2 = df2.explode("Type").reset_index(drop=True)


    cols =  df2["Type"].unique()

    df_new_database = df2.copy()

    for col in cols:
        
        # adjusting owners informations in columns
        df_filtered = df_new_database[df_new_database["Type"]== col]
        df_filtered = df_filtered[["Pos id","Firstname","Lastname","Phone"]]
        df_filtered = df_filtered.rename(columns={"Firstname": col+"Firstname", "Lastname": col+"Lastname","Phone": col+"Phone"})
        df_new_database = df_new_database.merge(df_filtered, on = "Pos id", how = "left")
        

    # create column region_name

    df_new_database = df_new_database.drop_duplicates(subset="Pos id")

    Regions = {1:"Center", 2:"East", 3:"West"}
    df["Region_name"] = df["Region"].map(Regions)



    # fill empty site id in new database from existing ones in old database
    if df['Site ID'].isnull().any():
        lookup = dict(zip(data_pos['Pos id'].astype(str), data_pos['Site ID']))
        df['Site ID'] = df['Site ID'].fillna(df['Pos id'].map(lookup))

    # merging the two sheets
    df = df.merge(df_new_database,on = "Pos id", how = "left")


    # detecting new stores in the new database
    missing_in_df = df[~df["Pos id"].isin(data_pos["Pos id"])]

    unique_posid = missing_in_df["Pos id"].unique()

    df = df.sort_values(by="Create Time")

    # Get the last valid site ID once before the loop
    last_site_id = df.loc[df["Site ID"].notna(), "Site ID"].iloc[-1]
    last_num = int(last_site_id[1:])
    
    if not missing_in_df.empty:

        st.success("New Stores Are Detected")

        for pos in unique_posid:
            
            last_num += 1
            site_id = f"C{last_num:09d}"
            if site_id in (df["Site ID"].values):
                last_num += 1
            else:
                df.loc[df["Pos id"] == pos, "Site ID"] = site_id


        df = df.merge(data_pos[['Pos id','Wilaya(48)','APP ID','BBE CODE','Pos Type','Rooting','Coverage','Visit Day','Status']], on = 'Pos id', how = 'left')

        df = df[['Region_name','Pos id', 'Status' ,'Site ID', 'Name', 'Wilaya','Wilaya(48)', 'APP ID', 'BBE CODE', 'Pos Type','Commune',
        'Address', 'Rooting', 'Coverage', 'Visit Day','Latitude', 'Longitude','Photo','PropriétaireFirstname','PropriétaireLastname',
        'PropriétairePhone', 'GérantFirstname','GérantLastname', 'GérantPhone', 'VendeurFirstname', 'VendeurLastname','VendeurPhone']]

        df['Status'] = df['Status'].fillna("Actif")

        df["Pos id"] = df["Pos id"].astype(int)

        df = df.fillna(" ")

        database_worksheet.clear()

        values = [df.columns.tolist()] + df.values.tolist()

        database_worksheet.update(values)

        st.success("New Database is Written to google Sheet")
            

    else:

        st.error("No New Stores Were Detected")

        return data_pos
            
    return df


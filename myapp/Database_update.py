import pandas as pd
import gspread
import streamlit as st
from dictionnary import base,commune_to_postcode
import numpy as np


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
    
    df = df.astype(str)

    df = df[["System ID","Create Time","Update Time","Delete Time","Last Visit","Region","District","Territory","Code","Name","Wilaya","Commune","Address",
             'Latitude','Longitude', "Area"]]

    df["Address"] = df["Address"].str.upper()

    
    # the DISTRICT ID for the first sheet
    
    df['DISTRICT ID'] = (
        df["Region"].fillna(0).astype(int).astype(str) +
        df["District"].fillna(0).astype(int).astype(str).str.zfill(2) +
        df["Territory"].fillna(0).astype(int).astype(str) +
        df["Code"].fillna(0).astype(int).astype(str).str.zfill(5)).astype(str)

    # reading the second sheet 
    df2 = pd.read_excel(pos_report, sheet_name = "Contacts")

    # creating the DISTRICT ID for the second sheet
    df2['DISTRICT ID'] = (
        df2["Region"].fillna(0).astype(int).astype(str) +
        df2["District"].fillna(0).astype(int).astype(str).str.zfill(2) +
        df2["Territory"].fillna(0).astype(int).astype(str) +
        df2["Code"].fillna(0).astype(int).astype(str).str.zfill(5)).astype(str)


    df2 = df2[["DISTRICT ID","Type","Firstname","Lastname","Phone"]]

    # separating values like gerant, proprietaire, vendeur
    df2["Type"] = df2["Type"].str.split(", ")
    df2 = df2.explode("Type").reset_index(drop=True)
    #df2["Phone"] = df2["Phone"].replace(' ',0)
    df2["Phone"] = df2["Phone"].fillna(0)


    cols =  df2["Type"].unique()

    df2 = df2.astype(str)

    df_new_database = df2.copy()

    for col in cols:
        
        # adjusting owners informations in columns
        df_filtered = df_new_database[df_new_database["Type"]== col]
        df_filtered = df_filtered[["DISTRICT ID","Firstname","Lastname","Phone"]]
        df_filtered = df_filtered.rename(columns={"Firstname": col+"Firstname", "Lastname": col+"Lastname","Phone": col+"Phone"})
        df_new_database = df_new_database.merge(df_filtered, on = "DISTRICT ID", how = "left")
        

    # create column region_name

    df_new_database = df_new_database.drop_duplicates(subset="DISTRICT ID")

    Regions = {'1':"Center", '2':"East", '3':"West"}
    df["Region_name"] = df["Region"].map(Regions)


    mapping_postcode = {k: v for k, v in commune_to_postcode.items()}

    df["POSTCODE"] = df["Commune"].map(mapping_postcode)

    # build a mapping dict from your base
    mapping = {k: v["numero"] for k, v in base.items()}

    # map the DataFrame column (assumes exact name matches)
    df["Wilaya(48)"] = df["Wilaya"].map(mapping)

    df["Wilaya(48)_str"] = (
        df["Wilaya(48)"]
        .fillna(0)
        .astype(int)
        .astype(str)
        .str.zfill(2)
    )

    # Precompute padded district
    df["District_str"] = (
        df["District"]
        .fillna(0)
        .astype(int)
        .astype(str)
        .str.zfill(3)
    )

    # APP ID
    df["APP CREDENTIALS"] = "R" + df["Region"] + "ORB" + df["District_str"]

    # Wholesale district list
    ws_districts = ["043", "036", "037", "038"]

    
    # BBE CODE
    df["BBE (ORB)"] = np.where(
        df["District_str"].isin(ws_districts),
        "ORB" + df["District_str"] + "#"+"WS" +"#" + df["Wilaya(48)_str"],
        "ORB" + df["District_str"] + "#"+"R" +"#" + df["Wilaya(48)_str"]
    )

    # Pos Type
    df["Pos Type"] = np.where(
        df["District_str"].isin(ws_districts),
        "WHOLESALE",
        "RETAIL"
    )

    # Cluster
    df["Cluster"] = np.where(
        df["District_str"].isin(ws_districts),
        1,
        0
    )

    # merging the two sheets
    df = df.merge(df_new_database,on = "DISTRICT ID", how = "left")

    df['Latitude'] = pd.to_numeric(df['Latitude'], errors='coerce')
    df['Longitude'] = pd.to_numeric(df['Longitude'], errors='coerce')

    df['Latitude'] = df['Latitude'].fillna(0)
    df['Longitude'] = df['Longitude'].fillna(0)

    df["PropriétairePhone"] = df["PropriétairePhone"].fillna(0)
    df["GérantPhone"] = df["GérantPhone"].fillna(0)

    st.write(data_pos.columns)

    df = df.merge(data_pos[['DISTRICT ID','Latitude','Longitude']],on = "DISTRICT ID", how = 'left',suffixes=('_x', '_y'))

    df = df.merge(data_pos[['DISTRICT ID','Site ID','Coverage','Visit Day','POS Stat','Grade','Photo',
                            'CHANNEL','SUB-CHANNEL','DATA TYPE','TYPE DE PDV','SITUATION GEOGRAPHIQUE']] , on = "DISTRICT ID", how = 'left')
    
    df = df.sort_values(by="Create Time")
    
    df['Latitude_y'] = df['Latitude_y'].astype(str).str.replace(',', '.', regex=False)
    df['Latitude_y'] = pd.to_numeric(df['Latitude_y'], errors='coerce')

    df['Longitude_y'] = df['Longitude_y'].astype(str).str.replace(',', '.', regex=False)
    df['Longitude_y'] = pd.to_numeric(df['Longitude_y'], errors='coerce')

    df['Latitude_y'] = df['Latitude_y'].fillna(0)
    df['Longitude_y'] = df['Longitude_y'].fillna(0)

    df['Latitude'] = np.where(
    df['Latitude_x'] != 0,
    df['Latitude_x'],
    df['Latitude_y'])

    df['Longitude'] = np.where(
    df['Longitude_x'] != 0,
    df['Longitude_x'],
    df['Longitude_y'])

    df.drop(columns=['Latitude_x','Longitude_x','Latitude_y','Longitude_y'],inplace=True)

    df = df[['Region_name','Wilaya','Commune','POSTCODE','Cluster', 'BBE (ORB)','APP CREDENTIALS','DISTRICT ID','POS Stat','Site ID','Grade','Pos Type','Name',
            'PropriétaireFirstname','PropriétaireLastname','PropriétairePhone', 'GérantFirstname','GérantLastname', 'GérantPhone', 'Address','Photo','CHANNEL',
            'SUB-CHANNEL','DATA TYPE','TYPE DE PDV','SITUATION GEOGRAPHIQUE','Area','Latitude','Longitude','Coverage', 'Visit Day']]

    # detecting new stores in the new database
    missing_in_df = df[~df["DISTRICT ID"].isin(data_pos["DISTRICT ID"])]

    df["DISTRICT ID"] = df["DISTRICT ID"].astype("int64")

    unique_posid = missing_in_df["DISTRICT ID"].unique()

    # Get the last valid site ID once before the loop
    last_site_id = df.loc[df["Site ID"].notna(), "Site ID"].iloc[-1]
    last_num = int(last_site_id[1:])

    if not missing_in_df.empty:

        st.success("New Stores Are Detected")

        for pos in unique_posid:

            pos = int(pos)
            
            last_num += 1
            site_id = f"C{last_num:09d}"
            if site_id in (df["Site ID"].values):
                last_num += 1
            else:
                df.loc[df["DISTRICT ID"] == pos, "Site ID"] = site_id

            df = df.fillna(" ")

            df['POS Stat'] = df['POS Stat'].replace({" ":"Actif"})

            df['Coverage'] = df['Coverage'].replace({" ": 1})

        df["Coverage"] = df["Coverage"].astype("int64")

        database_worksheet.clear()

        values = [df.columns.tolist()] + df.values.tolist()

        database_worksheet.update(values)

        st.success("New Database is Written to google Sheet")
            

    else:

        st.error("No New Stores Were Detected")

        return data_pos
            
    return df


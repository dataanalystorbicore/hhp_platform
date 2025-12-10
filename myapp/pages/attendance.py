import streamlit as st
import pandas as pd
from datetime import date
import gspread


service_account_info = st.secrets["gcp_service_account"]
gc = gspread.service_account_from_dict(service_account_info)


st.set_page_config(layout="wide")

attendance_sheet = gc.open("attendancereport").sheet1
bbe_sheet = gc.open("bbe_info").sheet1

# === Load BBE data ===
bbe_values = bbe_sheet.get_all_values()
df_bbe = pd.DataFrame(bbe_values[1:], columns=bbe_values[0])
names = df_bbe["NAME"].tolist()

# --- Helper callbacks to enforce mutual exclusion ---
def present_changed(name):
    # if present checkbox is checked, uncheck absent and leave for that name
    kp = f"present_{name}"
    if st.session_state.get(kp):
        st.session_state[f"absent_{name}"] = False
        st.session_state[f"leave_{name}"] = False

def absent_changed(name):
    ka = f"absent_{name}"
    if st.session_state.get(ka):
        st.session_state[f"present_{name}"] = False
        st.session_state[f"leave_{name}"] = False

def leave_changed(name):
    kl = f"leave_{name}"
    if st.session_state.get(kl):
        st.session_state[f"present_{name}"] = False
        st.session_state[f"absent_{name}"] = False

# Initialize keys if they don't exist (streamlit will create keys on first widget render,
# but it's convenient to ensure they exist so we can set them programmatically)
for name in names:
    for prefix in ("present_", "absent_", "leave_"):
        key = f"{prefix}{name}"
        if key not in st.session_state:
            st.session_state[key] = False

# === UI ===
st.title("🗓️ Daily Attendance")

selected_date = st.date_input("Select attendance date", date.today())

st.markdown("---")

# === Select All and Clear All buttons ===
col_control_a, col_control_b = st.columns([1, 1])
with col_control_a:
    if st.button("✅ Select All as Present"):
        for name in names:
            st.session_state[f"present_{name}"] = True
            st.session_state[f"absent_{name}"] = False
            st.session_state[f"leave_{name}"] = False

with col_control_b:
    if st.button("🧹 Clear All"):
        for name in names:
            st.session_state[f"present_{name}"] = False
            st.session_state[f"absent_{name}"] = False
            st.session_state[f"leave_{name}"] = False

st.markdown("---")

# === Three plain columns (no fixed boxes) ===
col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("### ✅ Presence")
    for name in names:
        # do NOT pass value=... — let Streamlit persist via key/session_state
        st.checkbox(name, key=f"present_{name}", on_change=lambda n=name: present_changed(n))

with col2:
    st.markdown("### ❌ Absence")
    for name in names:
        st.checkbox(name, key=f"absent_{name}", on_change=lambda n=name: absent_changed(n))

with col3:
    st.markdown("### 🏖️ Leave")
    for name in names:
        st.checkbox(name, key=f"leave_{name}", on_change=lambda n=name: leave_changed(n))

# === Build attendance dataframe ===
attendance_rows = []
for name in names:
    if st.session_state.get(f"present_{name}"):
        status = "Present"
    elif st.session_state.get(f"absent_{name}"):
        status = "Absent"
    elif st.session_state.get(f"leave_{name}"):
        status = "Congé"
    else:
        status = None

    if status:
        wilaya = df_bbe.loc[df_bbe["NAME"] == name, "WILAYA"].values[0]
        attendance_rows.append({
            "date": selected_date.isoformat(),
            "name": name,
            "wilaya": wilaya,
            "status": status
        })

df_to_save = pd.DataFrame(attendance_rows)

st.markdown("---")
st.write(f"Selected date: **{selected_date.isoformat()}**")
st.write(f"Total marked: **{len(df_to_save)}**")

# === Save logic: remove same-date rows then append ===
if st.button("💾 Save Attendance"):
    try:
        all_values = attendance_sheet.get_all_values()
        if all_values and len(all_values) > 0:
            # if header exists, assume first row is header
            if len(all_values) > 1:
                df_existing = pd.DataFrame(all_values[1:], columns=all_values[0])
            else:
                df_existing = pd.DataFrame(columns=all_values[0])
        else:
            df_existing = pd.DataFrame(columns=["date", "code", "name", "wilaya", "status"])

        # ensure expected columns
        expected_cols = ["date", "code", "name", "wilaya", "status"]
        for c in expected_cols:
            if c not in df_existing.columns:
                df_existing[c] = ""

        # remove existing rows for selected date
        df_existing_filtered = df_existing[df_existing["date"] != selected_date.isoformat()].copy()

        # assign codes for new rows
        start_code = 1
        try:
            existing_codes = pd.to_numeric(df_existing_filtered["code"], errors="coerce")
            if existing_codes.notna().any():
                start_code = int(existing_codes.max()) + 1
        except Exception:
            start_code = len(df_existing_filtered) + 1

        if not df_to_save.empty:
            df_to_save = df_to_save.reset_index(drop=True)
            df_to_save["code"] = [start_code + i for i in range(len(df_to_save))]
            # align columns
            df_to_save = df_to_save[expected_cols]

        final_df = pd.concat([df_existing_filtered, df_to_save], ignore_index=True, sort=False)
        final_df = final_df[expected_cols]

        attendance_sheet.clear()
        attendance_sheet.update([final_df.columns.values.tolist()] + final_df.values.tolist(),
                                value_input_option="USER_ENTERED")

        st.success(f"✅ Attendance for {selected_date.isoformat()} saved successfully! ({len(df_to_save)} rows)")
    except Exception as e:
        st.error(f"❌ Error saving to Google Sheets: {e}")

import streamlit as st
import pandas as pd
from datetime import datetime, date
import gspread
from google.oauth2.service_account import Credentials
import os
import json
from dotenv import load_dotenv

load_dotenv()

st.set_page_config(page_title="Permit to Work", page_icon="📝")

emirate = st.session_state.get("emirate", None)
if not emirate:
    st.warning("Please go to Home and select your Emirate first.")
    st.stop()

color = "🔵" if emirate == "Abu Dhabi" else "🔴"
st.title(f"📝 Permit to Work — {color} {emirate}")

# Permit types per emirate
permit_types_ad = [
    "Hot Work Permit (CoP 21.0 + 28.0)",
    "Confined Space Entry Permit (CoP 21.0 + 27.0)",
    "Excavation Permit (CoP 21.0 + 29.0)",
    "Working at Height Permit (CoP 21.0 + 23.0)",
    "Electrical Isolation / LOTO Permit (CoP 21.0 + 15.0)",
    "Lifting Operations Permit (CoP 21.0 + 34.0)",
    "Radiography / NDT Permit (CoP 21.0)"
]

permit_types_dubai = [
    "Hot Work Permit (DM Code Ch.15)",
    "Confined Space Entry Permit (DM Code Ch.19)",
    "Excavation Permit (DM Code Ch.9)",
    "Electrical Isolation / LOTO Permit (DM Code Ch.17)",
    "Radiography / NDT Permit",
    "Working at Height Permit",
    "Lifting Operations Permit (DM Code Ch.22)"
]

permit_types = permit_types_ad if emirate == "Abu Dhabi" else permit_types_dubai

st.markdown(f"**Regulatory Reference:** {'ADOSH CoP 21.0 — Permit to Work Systems' if emirate == 'Abu Dhabi' else 'Dubai Municipality Code — Ch.2.4.4.11'}")

# Form
with st.form("ptw_form"):
    st.subheader("Permit Details")
    
    col1, col2 = st.columns(2)
    with col1:
        permit_type = st.selectbox("Permit Type", permit_types)
        project = st.text_input("Project Name")
        location = st.text_input("Exact Work Location on Site")
        work_date = st.date_input("Work Date", value=date.today())
    
    with col2:
        permit_number = st.text_input("Permit Number", value=f"PTW-{emirate[:2].upper()}-{datetime.now().strftime('%Y%m%d-%H%M')}")
        start_time = st.time_input("Planned Start Time")
        end_time = st.time_input("Planned End Time")
        workers_count = st.number_input("Number of Workers", min_value=1, value=3)
    
    st.subheader("Personnel")
    col3, col4 = st.columns(2)
    with col3:
        permit_issuer = st.text_input("Permit Issuer Name & Title")
        work_supervisor = st.text_input("Work Supervisor Name")
    with col4:
        permit_receiver = st.text_input("Permit Receiver Name")
        safety_officer = st.text_input("HSE Officer Name")
    
    st.subheader("Work Description")
    work_description = st.text_area("Describe the work to be carried out", height=100)
    
    st.subheader("Hazards Identified")
    hazards = st.text_area("List main hazards for this work", height=80)
    
    st.subheader("Controls in Place (tick all that apply)")
    col5, col6 = st.columns(2)
    with col5:
        ppe = st.checkbox("✅ PPE provided and worn")
        isolation = st.checkbox("✅ Equipment isolated (LOTO)")
        atmosphere = st.checkbox("✅ Atmosphere tested (confined space/excavation)")
        ventilation = st.checkbox("✅ Ventilation in place")
    with col6:
        fire_watch = st.checkbox("✅ Fire watch in place (hot work)")
        barriers = st.checkbox("✅ Barriers and signage installed")
        first_aid = st.checkbox("✅ First aid available on site")
        emergency_plan = st.checkbox("✅ Emergency plan communicated to all")
    
    st.subheader("Emergency Information")
    emergency_col1, emergency_col2 = st.columns(2)
    with emergency_col1:
        if emirate == "Abu Dhabi":
            st.info("🚨 **Abu Dhabi Emergency Numbers:**\nAmbulance: 998 | Police: 999 | Civil Defence: 997")
        else:
            st.info("🚨 **Dubai Emergency Numbers:**\nPolice: 999 | Civil Defence: 997 | DM: 800900 | Ambulance: 998")
    with emergency_col2:
        nearest_hospital = st.text_input("Nearest Hospital / Clinic")
        muster_point = st.text_input("Site Muster Point Location")
    
    submit = st.form_submit_button("📤 Issue Permit", type="primary", use_container_width=True)

if submit:
    # Save to Google Sheets
    try:
        scope = ["https://spreadsheets.google.com/feeds", 
                 "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_file(
            os.getenv("GOOGLE_CREDENTIALS_PATH", "google_credentials.json"),
            scopes=scope
        )
        client_gs = gspread.authorize(creds)
        sheet = client_gs.open(os.getenv("GOOGLE_SHEET_NAME", "ADOSH_HSE_Records"))
        
        # Try to get PTW worksheet, create if doesn't exist
        try:
            ws = sheet.worksheet("Permits_to_Work")
        except:
            ws = sheet.add_worksheet(title="Permits_to_Work", rows=1000, cols=20)
            ws.append_row(["Permit No", "Emirate", "Type", "Project", "Date", 
                          "Issuer", "Receiver", "Status", "Timestamp"])
        
        ws.append_row([
            permit_number, emirate, permit_type, project,
            str(work_date), permit_issuer, permit_receiver, 
            "ISSUED", datetime.now().isoformat()
        ])
        
        st.success(f"✅ Permit {permit_number} issued and saved to records!")
        
    except Exception as e:
        st.warning(f"Permit generated but could not save to Google Sheets: {str(e)}")
        st.success(f"✅ Permit {permit_number} generated successfully!")
    
    # Display permit summary
    st.markdown("---")
    st.markdown(f"""
    ## {color} PERMIT TO WORK — {permit_number}
    **Emirate:** {emirate} | **Type:** {permit_type}
    **Project:** {project} | **Location:** {location}
    **Date:** {work_date} | **Time:** {start_time} — {end_time}
    
    **Permit Issuer:** {permit_issuer}  
    **Permit Receiver:** {permit_receiver}
    **Supervisor:** {work_supervisor}
    
    **Work Description:** {work_description}
    
    ---
    ⚠️ *This permit is valid for one shift only. Must be renewed if work continues. 
    Display at work location at all times.*
    """)
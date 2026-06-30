import streamlit as st
import anthropic
import os
from dotenv import load_dotenv
from datetime import date

load_dotenv()

st.set_page_config(page_title="Toolbox Talks", page_icon="🗣️")

emirate = st.session_state.get("emirate", None)
if not emirate:
    st.warning("Please go to Home and select your Emirate first.")
    st.stop()

color = "🔵" if emirate == "Abu Dhabi" else "🔴"
st.title(f"🗣️ Toolbox Talk Generator — {color} {emirate}")

topics_ad = [
    "Heat Stress & Safety in the Heat",
    "Working at Heights",
    "Scaffolding Safety",
    "Excavation Safety",
    "Permit to Work System",
    "Personal Protective Equipment",
    "Electrical Safety & LOTO",
    "Confined Space Entry",
    "Hot Work Safety",
    "Lifting Operations Safety",
    "Manual Handling",
    "Fire Prevention and Response",
    "Traffic Safety on Site",
    "Housekeeping & Slip/Trip/Fall",
    "Chemical Safety & MSDS",
    "Noise and Vibration",
    "Incident Reporting",
    "Stop Work Authority",
    "Last Minute Risk Assessment (Take 5)",
    "Summer Midday Work Ban"
]

topics_dubai = [
    "Heat Stress & Summer Safety",
    "Excavation Safety (Dubai Code Ch.9)",
    "Scaffolding Safety (Dubai Code Ch.8)",
    "Electrical Safety (Dubai Code Ch.16)",
    "Confined Space Entry (Dubai Code Ch.19)",
    "Lockout/Tagout (Dubai Code Ch.17)",
    "Fire Safety (Dubai Code Ch.5)",
    "PPE Requirements (Dubai Code Ch.4)",
    "Material Handling & Storage (Dubai Code Ch.7)",
    "Crane & Lifting Safety (Dubai Code Ch.22)",
    "Working at Heights",
    "Accident Reporting Procedures",
    "Permit to Work System",
    "Welding & Cutting Safety (Dubai Code Ch.15)",
    "Chemical Safety (Dubai Code Ch.18)",
    "Site Induction Requirements",
    "Signs and Barricades (Dubai Code Ch.6)",
    "Underground Services",
    "Demolition Safety (Dubai Code Ch.13)",
    "Steel Erection Safety (Dubai Code Ch.11)"
]

topics = topics_ad if emirate == "Abu Dhabi" else topics_dubai

col1, col2 = st.columns(2)
with col1:
    topic = st.selectbox("Select Topic", topics)
    project_name = st.text_input("Project Name", placeholder="e.g. Bloom Living Almeria")
    conducted_by = st.text_input("Conducted By", placeholder="e.g. Aftab Qamar — Senior HSSE Engineer")

with col2:
    workers_count = st.number_input("Number of Attendees", min_value=1, max_value=200, value=20)
    tbt_date = st.date_input("Date", value=date.today())
    language = st.selectbox("Language", ["English", "Arabic", "English + Arabic"])

if st.button("⚡ Generate Toolbox Talk", type="primary", use_container_width=True):
    with st.spinner(f"Generating {topic} toolbox talk for {emirate}..."):
        try:
            client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
            
            ref = "ADOSH CoP" if emirate == "Abu Dhabi" else "Dubai Municipality Code of Construction Safety Practice"
            authority = "ADOSH (ADOSH-SF Version 4.0, July 2024)" if emirate == "Abu Dhabi" else "Dubai Municipality"
            
            prompt = f"""Generate a complete, professional Toolbox Talk on: {topic}
            
Emirate: {emirate} ONLY
Regulatory Authority: {authority}
Reference: {ref}
Project: {project_name}
Date: {tbt_date}
Conducted by: {conducted_by}
Attendees: {workers_count} workers

Format the toolbox talk EXACTLY as follows:

# TOOLBOX TALK — {topic.upper()}
**{color} {emirate} SPECIFIC | Ref: [specific regulation reference]**
**Project:** {project_name} | **Date:** {tbt_date} | **By:** {conducted_by}

---

## 🎯 KEY MESSAGE (1 sentence, bold, memorable)

## ⚠️ KEY HAZARDS
(3-5 bullet points — specific hazards for this topic)

## 🛡️ WHAT YOU MUST DO — CONTROLS
(5-8 actionable bullet points with specific {emirate} regulatory requirements)

## 🚨 EMERGENCY PROCEDURE
(What to do if something goes wrong — include emergency numbers)

## ❓ DISCUSSION QUESTIONS (3 questions for the supervisor to ask workers)

## ✅ KEY TAKEAWAY (1 sentence)

---
**Attendance Record:** (table with Name | Employee ID | Signature | Date)

End with: ⚠️ This toolbox talk references {emirate} regulations. Verify against current official publications.
"""
            
            response = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=2000,
                messages=[{"role": "user", "content": prompt}]
            )
            
            tbt_content = response.content[0].text
            st.markdown(tbt_content)
            
            # Download button
            st.download_button(
                label="⬇️ Download Toolbox Talk",
                data=tbt_content,
                file_name=f"TBT_{topic.replace(' ', '_')}_{tbt_date}_{emirate}.md",
                mime="text/markdown"
            )
            
        except Exception as e:
            st.error(f"Error: {str(e)}")
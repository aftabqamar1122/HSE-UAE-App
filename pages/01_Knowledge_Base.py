import streamlit as st
import json

st.set_page_config(page_title="HSE Knowledge Base", page_icon="📚")

emirate = st.session_state.get("emirate", None)
if not emirate:
    st.warning("Please go to Home and select your Emirate first.")
    st.stop()

color = "🔵" if emirate == "Abu Dhabi" else "🔴"
st.title(f"📚 HSE Knowledge Base — {color} {emirate}")

# ── ABU DHABI CONTENT ──────────────────────────────────
if emirate == "Abu Dhabi":
    st.markdown("**Source: ADOSH-SF Version 4.0 | ADPHC**")
    
    # Search
    search = st.text_input("🔍 Search CoPs by keyword or activity", 
                            placeholder="e.g. excavation, heat, scaffolding, lifting...")
    
    cops = {
        "CoP 2.0": ("Personal Protective Equipment (PPE)", "Critical", 
                     "Selection, provision, and use of PPE. Employer must provide at no cost. Hierarchy of controls applies first."),
        "CoP 11.0": ("Safety in the Heat", "Critical",
                      "WBGT monitoring. Midday ban 12:30-15:00, 15 Jun-15 Sep. Acclimatisation 2 weeks. 250ml water/hour minimum."),
        "CoP 21.0": ("Permit to Work Systems", "Critical",
                      "Mandatory for: Hot Work, Confined Space, Excavation, Working at Heights, LOTO, Lifting, Radiography. Valid one shift only."),
        "CoP 23.0": ("Working at Heights", "Critical",
                      "Hierarchy: Avoid > Prevent (guardrails) > Mitigate (nets/arrest). Full body harness mandatory. Edge protection: 950mm rail + 150mm toe board."),
        "CoP 26.0": ("Scaffolding", "Critical",
                      "Competent erectors only. Inspection every 7 days + after adverse weather. Green/Red tag system. SWL marked."),
        "CoP 29.0": ("Excavation Work", "Critical",
                      "Permit required >0.5m. NOC from utilities. CAT scanner mandatory. Max 45° repose. Daily inspections. Spoil ≥0.6m from edge."),
        "CoP 34.0": ("Safe Use of Lifting Equipment", "Critical",
                      "Critical lift >75% SWL or >10T. Annual crane inspection (3rd party). 6-monthly accessories. Exclusion zone under loads."),
        "CoP 44.0": ("Traffic Management and Logistics", "Critical",
                      "Traffic Management Plan required. TR-531 reference. Designated logistics routes. Banksman for reversing vehicles."),
        "CoP 53.1": ("OSH Construction Management Plan", "Critical",
                      "14 mandatory sections. Submitted to consultant before work starts. ISO 45001:2018 aligned."),
        "CoP 54.0": ("Waste Management", "Critical",
                      "TADWEER waste manifests required. Segregation: general/hazardous/recyclable/inert. Licensed carriers only. No burning on site."),
        "CoP 15.0": ("Electrical Safety", "High",
                      "LOTO before all electrical work. Qualified electricians only. Grounding required. No metal ladders near electrical."),
        "CoP 27.0": ("Confined Spaces", "Critical",
                      "Written safety plan. Permit required. O₂: 19.5-23.5%. Combustibles <10% LEL. Mechanical ventilation. Attendant at entry point."),
        "CoP 28.0": ("Hot Work Operations", "Critical",
                      "Hot Work Permit mandatory. Fire watch during and 30 min after. Fire extinguisher at hot work location. Remove/protect combustibles."),
        "CoP 40.0": ("False Work (Formwork)", "Critical",
                      "Structural engineer design for complex formwork. Striking schedule based on concrete strength. Inspection before concrete pour."),
        "CoP 42.0": ("Pre-Cast Construction", "Critical",
                      "Lifting plan required for all precast. Rated lifting inserts. Crane capacity checks. Temporary propping."),
    }
    
    # Filter by search
    filtered = {k: v for k, v in cops.items() 
                if not search or search.lower() in k.lower() or 
                search.lower() in v[0].lower() or 
                search.lower() in v[2].lower()}
    
    # Display by criticality
    critical = {k: v for k, v in filtered.items() if v[1] == "Critical"}
    high = {k: v for k, v in filtered.items() if v[1] == "High"}
    
    if critical:
        st.markdown("### 🔴 Critical CoPs — Construction")
        for cop_num, (title, level, summary) in critical.items():
            with st.expander(f"**{cop_num} — {title}**"):
                st.markdown(f"**Summary:** {summary}")
                st.markdown(f"**Reference:** ADOSH-SF v4.0 | {cop_num}")
                st.markdown(f"**Source:** [ADPHC Website](https://www.adphc.gov.ae/en/Legislation/Code-of-Practices)")
    
    if high:
        st.markdown("### 🟡 High Priority CoPs")
        for cop_num, (title, level, summary) in high.items():
            with st.expander(f"**{cop_num} — {title}**"):
                st.markdown(f"**Summary:** {summary}")

# ── DUBAI CONTENT ──────────────────────────────────────
else:
    st.markdown("**Source: Dubai Municipality Code of Construction Safety Practice**")
    
    search = st.text_input("🔍 Search chapters by keyword", 
                            placeholder="e.g. excavation, electrical, confined space, scaffolding...")
    
    chapters = {
        "Chapter 1": ("Preliminary — Definitions & Scope", 
                      "Defines: Contractor, Engineer, Inspector, Construction Works, Scaffolds, Occupational Disease, Serious Accident, LTI, Minor Injury."),
        "Chapter 2": ("General Health and Safety Provision",
                      "Joint Engineer+Contractor responsibility. Safety Plan mandatory. Safety staff ratios (Tables 1-4). Accident reporting 72hrs to DM. Records 5 years. Safety meetings every 15 days."),
        "Chapter 3": ("Occupational Health and Environmental Control",
                      "First aid ratios (Table 1). First aid box contents (Table 2). First aid room ≥250 employees (Table 3). Eye wash ≤30m. Medical examinations."),
        "Chapter 4": ("Personal Protective Equipment",
                      "Employer provides ALL PPE at no cost. Must be appropriate to specific hazards. Supervisor to ensure effective use."),
        "Chapter 5": ("Fire Protection and Prevention",
                      "Fire prevention plan. Employee training on extinguisher use. Hot work permits. Never use water on electrical fires — CO₂ or dry chemical only."),
        "Chapter 6": ("Signs, Signals and Barricades",
                      "Mandatory at all worksites. Illuminated from dusk to dawn at excavations. Warning signs in Arabic and English."),
        "Chapter 8": ("Scaffolding Safety",
                      "Only competent persons may erect/dismantle/inspect scaffolds. Good quality materials. Thorough inspection before use."),
        "Chapter 9": ("Excavation",
                      "Shoring >1.25m depth. Max 40° angle of repose. NOC from all authorities before start. Spoil ≥60cm from edge. Access every 15m. Ladder extends 90cm above surface."),
        "Chapter 16": ("Electrical Hazards",
                       "All installations per DEWA. LOTO before maintenance. No jewellery/rings near circuits. CO₂/dry chemical only for electrical fires. Never touch electric shock victim — disconnect power first."),
        "Chapter 17": ("Control of Hazardous Energy (LOTO)",
                       "Full 9-step LOTO procedure. Individual locks per worker. Verify isolation before work. Release stored energy. Restore in reverse order."),
        "Chapter 19": ("Confined Spaces",
                        "Identify and list all confined spaces. Warning signs at openings. O₂: 19.5-23.5%. Combustibles <10% LEL (0% for welding). 20 air changes/hour ventilation. Attendant at entry at all times."),
        "Chapter 22": ("Cranes",
                        "Annual inspection by DM-approved body. Safety certificate required. Only DM-authorised operators. No person under suspended load."),
    }
    
    filtered = {k: v for k, v in chapters.items() 
                if not search or search.lower() in k.lower() or 
                search.lower() in v[0].lower() or 
                search.lower() in v[1].lower()}
    
    for ch_num, (title, summary) in filtered.items():
        with st.expander(f"**{ch_num} — {title}**"):
            st.markdown(f"**Summary:** {summary}")
            st.markdown(f"**Source:** Dubai Municipality Code of Construction Safety Practice")
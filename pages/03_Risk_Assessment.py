import streamlit as st
import anthropic
import os
from dotenv import load_dotenv
from datetime import date

load_dotenv()

st.set_page_config(page_title="Risk Assessment", page_icon="📋")

emirate = st.session_state.get("emirate", None)
if not emirate:
    st.warning("Please go to Home and select your Emirate first.")
    st.stop()

color = "🔵" if emirate == "Abu Dhabi" else "🔴"
st.title(f"📋 Risk Assessment Generator — {color} {emirate}")

if emirate == "Abu Dhabi":
    st.markdown("**Reference: ADOSH-SF v4.0 | ENGC 14-Column Format**")
else:
    st.markdown("**Reference: Dubai Municipality Code of Construction Safety Practice**")

st.divider()

# ── Input Form ────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    project_name = st.text_input(
        "Project Name",
        placeholder="e.g. Bloom Living Almeria"
    )
    company_name = st.text_input(
        "Company Name",
        value="ENGC"
    )
    prepared_by = st.text_input(
        "Prepared By",
        placeholder="e.g. Aftab Qamar — Senior HSSE Engineer"
    )

with col2:
    activity = st.text_input(
        "Work Activity / Task",
        placeholder="e.g. Excavation Work, Scaffolding, Lifting Operations"
    )
    ra_date = st.date_input("Date", value=date.today())
    location = st.text_input(
        "Work Location",
        placeholder="e.g. Plot C13, Zayed City, Abu Dhabi"
    )

st.divider()

# ── Activity Details ──────────────────────────────────────
st.subheader("Activity Details")

work_description = st.text_area(
    "Describe the work activity in detail",
    placeholder="e.g. Excavation for foundation works to a depth of 2.5m...",
    height=100
)

workers_involved = st.text_input(
    "Who might be harmed?",
    placeholder="e.g. Excavation crew, site workers, public nearby"
)

# ── Generate Button ───────────────────────────────────────
if st.button(
    "⚡ Generate Risk Assessment",
    type="primary",
    use_container_width=True
):
    if not activity:
        st.error("Please enter a work activity first.")
    else:
        with st.spinner(
            f"Generating risk assessment for {activity} "
            f"in {emirate}..."
        ):
            try:
                client = anthropic.Anthropic(
                    api_key=os.getenv("ANTHROPIC_API_KEY")
                )

                if emirate == "Abu Dhabi":
                    reg_ref = (
                        "ADOSH-SF Version 4.0 (July 2024) and relevant "
                        "ADOSH Codes of Practice"
                    )
                    format_note = (
                        "Use the ENGC 14-column format: "
                        "S/N | Activity Element | Hazards | "
                        "Risk & Consequences | Who Harmed | "
                        "P | S | R(PxS) | Initial Risk Level | "
                        "Controls | P | S | R(PxS) | Residual Risk Level"
                    )
                else:
                    reg_ref = (
                        "Dubai Municipality Code of Construction "
                        "Safety Practice"
                    )
                    format_note = (
                        "Use standard risk assessment format with "
                        "hazard, likelihood, severity, risk rating, "
                        "and control measures"
                    )

                prompt = f"""Generate a detailed, professional Risk 
Assessment for the following:

EMIRATE: {emirate} ONLY
REGULATORY REFERENCE: {reg_ref}
PROJECT: {project_name}
COMPANY: {company_name}
ACTIVITY: {activity}
LOCATION: {location}
DATE: {ra_date}
PREPARED BY: {prepared_by}
WORK DESCRIPTION: {work_description}
WHO MIGHT BE HARMED: {workers_involved}

FORMAT INSTRUCTIONS:
{format_note}

Risk Rating Scale:
- P (Probability): 1=Rare, 2=Unlikely, 3=Possible, 
  4=Likely, 5=Almost Certain
- S (Severity): 1=Negligible, 2=Minor, 3=Moderate, 
  4=Major, 5=Catastrophic
- R = P x S
- 1-4 = LOW (Green)
- 5-9 = MODERATE (Yellow)  
- 10-15 = HIGH (Orange)
- 16-25 = EXTREME (Red)

Generate at least 5 risk assessment rows covering the main 
hazards for this activity. For each row include:
1. The specific hazard
2. Risk before controls (Initial Risk)
3. Specific control measures referencing {emirate} regulations
4. Risk after controls (Residual Risk)

End with a summary statement and reminder to verify against 
current {emirate} regulations.

Label everything clearly as {color} {emirate} SPECIFIC."""

                response = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=3000,
                    messages=[{"role": "user", "content": prompt}]
                )

                ra_content = response.content[0].text

                st.success("✅ Risk Assessment Generated!")
                st.divider()
                st.markdown(ra_content)

                # Download button
                st.download_button(
                    label="⬇️ Download Risk Assessment",
                    data=ra_content,
                    file_name=(
                        f"RA_{activity.replace(' ', '_')}"
                        f"_{ra_date}_{emirate}.md"
                    ),
                    mime="text/markdown"
                )

            except Exception as e:
                st.error(f"Error generating risk assessment: {str(e)}")
                st.info(
                    "Please check your API key in the app secrets."
                )
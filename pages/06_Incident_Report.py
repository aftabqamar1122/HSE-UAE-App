import streamlit as st
import anthropic
import os
from dotenv import load_dotenv
from datetime import date, datetime, time
from io import BytesIO
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

load_dotenv()

st.set_page_config(
    page_title="Incident Report", page_icon="🚨"
)

emirate = st.session_state.get("emirate", None)
if not emirate:
    st.warning(
        "Please go to Home and select your Emirate first."
    )
    st.stop()

color = "🔵" if emirate == "Abu Dhabi" else "🔴"
st.title(f"🚨 Incident Report — {color} {emirate}")

if emirate == "Abu Dhabi":
    st.markdown(
        "**Reference: ADOSH-SF Version 4.0 | ADPHC**"
    )
    st.info(
        "⚠️ All incidents must be reported to ADOSH "
        "within the required timeframe per ADOSH-SF v4.0"
    )
else:
    st.markdown(
        "**Reference: Dubai Municipality Code — "
        "Article 2.4.5**"
    )
    st.error(
        "🚨 DUBAI REQUIREMENT: Written report to Dubai "
        "Municipality within **72 HOURS** of occurrence.\n\n"
        "Records must be kept for **5 YEARS**.\n\n"
        "Emergency: Police 999 | Civil Defence 997 | "
        "DM 800900"
    )

st.divider()

# ── HELPERS ───────────────────────────────────────────────

def set_cell_color(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def set_borders(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = OxmlElement('w:tcBorders')
    for side in ['top', 'left', 'bottom', 'right']:
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), '000000')
        borders.append(b)
    tcPr.append(borders)


def add_section_heading(doc, text, color_hex="C0C0C0"):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0, 0, 0)
    # Add background shading to paragraph
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    pPr.append(shd)
    return p


def add_field_row(tbl, label, value):
    row = tbl.add_row()
    lbl_cell = row.cells[0]
    val_cell = row.cells[1]
    lbl_cell.text = label
    lbl_cell.paragraphs[0].runs[0].bold = True
    lbl_cell.paragraphs[0].runs[0].font.size = Pt(9)
    set_cell_color(lbl_cell, "F2F2F2")
    set_borders(lbl_cell)
    val_cell.text = str(value) if value else "—"
    val_cell.paragraphs[0].runs[0].font.size = Pt(9)
    set_borders(val_cell)


def build_incident_docx(data, ai_analysis, emirate):
    doc = Document()

    # Page setup A4 Portrait
    sec = doc.sections[0]
    sec.page_width    = Cm(21.0)
    sec.page_height   = Cm(29.7)
    sec.left_margin   = Cm(2.0)
    sec.right_margin  = Cm(2.0)
    sec.top_margin    = Cm(2.0)
    sec.bottom_margin = Cm(2.0)

    # ── TITLE ──────────────────────────────────────────
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run(
        f"{data['company']} — INCIDENT REPORT"
    )
    r.bold = True
    r.font.size = Pt(16)
    r.font.color.rgb = RGBColor(192, 0, 0)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = sub.add_run(
        f"Emirate: {emirate}   |   "
        f"Report No: {data['report_no']}   |   "
        f"Date: {data['incident_date']}"
    )
    r2.font.size = Pt(10)
    r2.bold = True

    doc.add_paragraph()

    # ── SECTION 1: INCIDENT CLASSIFICATION ────────────
    add_section_heading(
        doc,
        "SECTION 1 — INCIDENT CLASSIFICATION",
        "FF0000"
    )
    tbl1 = doc.add_table(rows=0, cols=2)
    tbl1.style = 'Table Grid'
    tbl1.columns[0].width = Cm(6)
    tbl1.columns[1].width = Cm(11)

    add_field_row(tbl1, "Incident Type",
                  data['incident_type'])
    add_field_row(tbl1, "Injury Classification",
                  data['injury_class'])
    add_field_row(tbl1, "Report Number",
                  data['report_no'])
    add_field_row(tbl1, "Date of Incident",
                  str(data['incident_date']))
    add_field_row(tbl1, "Time of Incident",
                  str(data['incident_time']))
    add_field_row(tbl1, "Day of Week",
                  data['day_of_week'])

    doc.add_paragraph()

    # ── SECTION 2: PROJECT DETAILS ─────────────────────
    add_section_heading(
        doc,
        "SECTION 2 — PROJECT DETAILS",
        "C0C0C0"
    )
    tbl2 = doc.add_table(rows=0, cols=2)
    tbl2.style = 'Table Grid'

    add_field_row(tbl2, "Project Name",
                  data['project_name'])
    add_field_row(tbl2, "Company", data['company'])
    add_field_row(tbl2, "Location / Site",
                  data['location'])
    add_field_row(tbl2, "Reported By",
                  data['reported_by'])
    add_field_row(tbl2, "HSE Officer",
                  data['hse_officer'])
    add_field_row(tbl2, "Report Date",
                  str(data['report_date']))

    doc.add_paragraph()

    # ── SECTION 3: INJURED PERSON ──────────────────────
    add_section_heading(
        doc,
        "SECTION 3 — INJURED / AFFECTED PERSON DETAILS",
        "C0C0C0"
    )
    tbl3 = doc.add_table(rows=0, cols=2)
    tbl3.style = 'Table Grid'

    add_field_row(tbl3, "Name", data['injured_name'])
    add_field_row(tbl3, "Age", data['injured_age'])
    add_field_row(tbl3, "Nationality",
                  data['injured_nationality'])
    add_field_row(tbl3, "Trade / Designation",
                  data['injured_trade'])
    add_field_row(tbl3, "Employer / Contractor",
                  data['injured_employer'])
    add_field_row(tbl3, "Experience in Trade",
                  data['injured_experience'])
    add_field_row(tbl3, "Nature of Injury",
                  data['nature_of_injury'])
    add_field_row(tbl3, "Body Part Affected",
                  data['body_part'])
    add_field_row(tbl3, "Medical Treatment",
                  data['medical_treatment'])
    add_field_row(tbl3, "Hospital / Clinic",
                  data['hospital'])

    doc.add_paragraph()

    # ── SECTION 4: INCIDENT DESCRIPTION ───────────────
    add_section_heading(
        doc,
        "SECTION 4 — INCIDENT DESCRIPTION",
        "C0C0C0"
    )
    p_desc = doc.add_paragraph()
    p_desc.add_run(
        data['incident_description']
    ).font.size = Pt(10)
    doc.add_paragraph()

    # ── SECTION 5: IMMEDIATE CAUSES ───────────────────
    add_section_heading(
        doc,
        "SECTION 5 — IMMEDIATE CAUSES",
        "FFC000"
    )
    tbl5 = doc.add_table(rows=0, cols=2)
    tbl5.style = 'Table Grid'
    add_field_row(tbl5, "Unsafe Act",
                  data['unsafe_act'])
    add_field_row(tbl5, "Unsafe Condition",
                  data['unsafe_condition'])
    add_field_row(tbl5, "Activity at Time",
                  data['activity_at_time'])
    add_field_row(tbl5, "Equipment Involved",
                  data['equipment_involved'])

    doc.add_paragraph()

    # ── SECTION 6: ROOT CAUSES ────────────────────────
    add_section_heading(
        doc,
        "SECTION 6 — ROOT CAUSES",
        "FFC000"
    )
    tbl6 = doc.add_table(rows=0, cols=2)
    tbl6.style = 'Table Grid'
    add_field_row(tbl6, "Personal Factors",
                  data['root_personal'])
    add_field_row(tbl6, "Job / Task Factors",
                  data['root_job'])
    add_field_row(tbl6, "Management Factors",
                  data['root_management'])

    doc.add_paragraph()

    # ── SECTION 7: AI ANALYSIS ─────────────────────────
    add_section_heading(
        doc,
        "SECTION 7 — AI INVESTIGATION ANALYSIS & "
        "CORRECTIVE ACTIONS",
        "002060"
    )
    p_ai = doc.add_paragraph()
    r_ai = p_ai.add_run(ai_analysis)
    r_ai.font.size = Pt(9)

    doc.add_paragraph()

    # ── SECTION 8: REGULATORY REPORTING ───────────────
    add_section_heading(
        doc,
        "SECTION 8 — REGULATORY REPORTING REQUIREMENTS",
        "C0C0C0"
    )
    tbl8 = doc.add_table(rows=0, cols=2)
    tbl8.style = 'Table Grid'

    if emirate == "Abu Dhabi":
        add_field_row(
            tbl8, "Reporting Authority", "ADOSH / ADPHC"
        )
        add_field_row(
            tbl8, "Reference",
            "ADOSH-SF Version 4.0 (July 2024)"
        )
        add_field_row(
            tbl8, "Notification",
            "Immediate verbal notification to ADOSH "
            "for serious incidents"
        )
    else:
        add_field_row(
            tbl8, "Reporting Authority",
            "Dubai Municipality (DM) — Specialised Dept"
        )
        add_field_row(
            tbl8, "Reference",
            "DM Code Art. 2.4.5"
        )
        add_field_row(
            tbl8, "Written Report Deadline",
            "Within 72 HOURS of occurrence"
        )
        add_field_row(
            tbl8, "Emergency Numbers",
            "Police: 999 | Civil Defence: 997 | "
            "DM: 800900 | Ambulance: 998"
        )
        add_field_row(
            tbl8, "Records Retention",
            "5 YEARS from date of occurrence"
        )
        add_field_row(
            tbl8, "Police Notification",
            "Required for serious accidents and LTIs "
            "involving hospital transfer"
        )

    doc.add_paragraph()

    # ── SECTION 9: SIGNATURES ─────────────────────────
    add_section_heading(
        doc,
        "SECTION 9 — SIGNATURES",
        "C0C0C0"
    )
    sig_tbl = doc.add_table(rows=2, cols=3)
    sig_tbl.style = 'Table Grid'
    headers = [
        "Reported By", "HSE Officer", "Project Manager"
    ]
    for i, h in enumerate(headers):
        cell = sig_tbl.rows[0].cells[i]
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(9)
        set_cell_color(cell, "C0C0C0")
        set_borders(cell)
    for i in range(3):
        cell = sig_tbl.rows[1].cells[i]
        cell.text = "\n\nSignature: _______________\nDate:"
        cell.paragraphs[0].runs[0].font.size = Pt(9)
        set_borders(cell)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ── INCIDENT TYPE DEFINITIONS ─────────────────────────────
if emirate == "Abu Dhabi":
    incident_types = [
        "Near Miss",
        "First Aid Injury (FAI)",
        "Medical Treatment Case (MTC)",
        "Restricted Work Case (RWC)",
        "Lost Time Injury (LTI)",
        "Permanent Partial Disability (PPD)",
        "Permanent Total Disability (PTD)",
        "Fatality",
        "Property Damage",
        "Environmental Incident",
        "Fire / Explosion",
        "Dangerous Occurrence",
    ]
    injury_classes = [
        "Near Miss — No injury",
        "First Aid Only",
        "Medical Treatment (no lost time)",
        "Restricted Work (modified duties)",
        "Lost Time Injury (>1 day away from work)",
        "Serious Injury",
        "Fatality",
        "Property / Equipment Damage Only",
    ]
else:
    incident_types = [
        "Near Miss",
        "Minor Injury (First Aid)",
        "Lost Time Injury (LTI — >3 days)",
        "Serious / Major Injury",
        "Fatality",
        "Property Damage",
        "Fire / Explosion",
        "Structural Collapse",
        "Environmental Incident",
        "Dangerous Occurrence",
    ]
    injury_classes = [
        "Minor Injury (First Aid only)",
        "Lost Time Injury (>3 days absence)",
        "Serious Injury — fracture/amputation/burns",
        "Fatality",
        "Property Damage Only",
        "No Injury (Near Miss)",
    ]

# ── FORM ──────────────────────────────────────────────────
with st.form("incident_form"):

    # SECTION 1: Classification
    st.subheader("🔴 Section 1 — Incident Classification")
    col1, col2 = st.columns(2)
    with col1:
        incident_type = st.selectbox(
            "Incident Type", incident_types
        )
        incident_date = st.date_input(
            "Date of Incident", value=date.today()
        )
        report_no = st.text_input(
            "Report Number",
            value=f"IR-{emirate[:2].upper()}-"
                  f"{datetime.now().strftime('%Y%m%d-%H%M')}"
        )
    with col2:
        injury_class = st.selectbox(
            "Injury Classification", injury_classes
        )
        incident_time = st.time_input(
            "Time of Incident",
            value=time(8, 0)
        )
        day_of_week = st.selectbox(
            "Day of Week",
            ["Monday","Tuesday","Wednesday","Thursday",
             "Friday","Saturday","Sunday"]
        )

    st.divider()

    # SECTION 2: Project Details
    st.subheader("📋 Section 2 — Project Details")
    col3, col4 = st.columns(2)
    with col3:
        project_name = st.text_input(
            "Project Name",
            placeholder="e.g. Bloom Living Almeria"
        )
        company = st.text_input(
            "Company", value="ENGC"
        )
        location = st.text_input(
            "Location / Site",
            placeholder="e.g. Plot C13, Zayed City"
        )
    with col4:
        reported_by = st.text_input(
            "Reported By",
            placeholder="Name and designation"
        )
        hse_officer = st.text_input(
            "HSE Officer",
            placeholder="e.g. Aftab Qamar"
        )
        report_date = st.date_input(
            "Report Date", value=date.today()
        )

    st.divider()

    # SECTION 3: Injured Person
    st.subheader(
        "👷 Section 3 — Injured / Affected Person"
    )
    col5, col6 = st.columns(2)
    with col5:
        injured_name = st.text_input(
            "Full Name of Injured Person"
        )
        injured_age = st.text_input("Age")
        injured_nationality = st.text_input(
            "Nationality"
        )
        injured_trade = st.text_input(
            "Trade / Designation",
            placeholder="e.g. Carpenter, Electrician"
        )
        injured_employer = st.text_input(
            "Employer / Sub-contractor"
        )
    with col6:
        injured_experience = st.text_input(
            "Experience in Trade",
            placeholder="e.g. 3 years"
        )
        nature_of_injury = st.text_input(
            "Nature of Injury",
            placeholder="e.g. Laceration to right hand"
        )
        body_part = st.selectbox(
            "Body Part Affected",
            ["Head", "Eye", "Neck", "Shoulder",
             "Arm", "Hand", "Finger", "Back",
             "Chest", "Abdomen", "Leg", "Knee",
             "Foot", "Multiple", "None"]
        )
        medical_treatment = st.selectbox(
            "Medical Treatment Given",
            ["First Aid on Site",
             "Referred to Clinic",
             "Referred to Hospital",
             "Admitted to Hospital",
             "None Required"]
        )
        hospital = st.text_input(
            "Hospital / Clinic Name (if applicable)"
        )

    st.divider()

    # SECTION 4: Description
    st.subheader("📝 Section 4 — Incident Description")
    incident_description = st.text_area(
        "Describe what happened in detail "
        "(sequence of events leading to the incident)",
        height=150,
        placeholder=(
            "e.g. At approximately 09:30hrs, while "
            "the worker was carrying out excavation "
            "works at a depth of 1.8m, the trench "
            "wall collapsed without warning..."
        )
    )

    st.divider()

    # SECTION 5: Immediate Causes
    st.subheader("⚡ Section 5 — Immediate Causes")
    col7, col8 = st.columns(2)
    with col7:
        unsafe_act = st.text_area(
            "Unsafe Act (What did the person do wrong?)",
            height=80,
            placeholder=(
                "e.g. Working in unshored excavation "
                "deeper than 1.2m"
            )
        )
        activity_at_time = st.text_input(
            "Activity Being Performed at Time of Incident",
            placeholder="e.g. Excavation works"
        )
    with col8:
        unsafe_condition = st.text_area(
            "Unsafe Condition "
            "(What was wrong with the environment?)",
            height=80,
            placeholder=(
                "e.g. No shoring installed, "
                "no edge protection"
            )
        )
        equipment_involved = st.text_input(
            "Equipment / Material Involved",
            placeholder="e.g. Excavator, hand tools"
        )

    st.divider()

    # SECTION 6: Root Causes
    st.subheader("🔍 Section 6 — Root Causes")
    col9, col10 = st.columns(2)
    with col9:
        root_personal = st.text_area(
            "Personal Factors",
            height=80,
            placeholder=(
                "e.g. Lack of training, "
                "risk taking behaviour"
            )
        )
        root_job = st.text_area(
            "Job / Task Factors",
            height=80,
            placeholder=(
                "e.g. No method statement, "
                "inadequate supervision"
            )
        )
    with col10:
        root_management = st.text_area(
            "Management System Factors",
            height=80,
            placeholder=(
                "e.g. Permit to Work not issued, "
                "inadequate induction"
            )
        )

    st.divider()

    # Witnesses
    st.subheader("👁️ Witnesses")
    witnesses = st.text_area(
        "Names and designations of witnesses",
        height=60,
        placeholder="e.g. Ahmed Ali — Foreman"
    )

    submitted = st.form_submit_button(
        "🚨 Generate Incident Report (.docx)",
        type="primary",
        use_container_width=True
    )

# ── GENERATE REPORT ───────────────────────────────────────
if submitted:
    if not incident_description:
        st.error(
            "Please provide an incident description."
        )
    else:
        with st.spinner(
            "Generating incident report with "
            "AI analysis..."
        ):
            try:
                client = anthropic.Anthropic(
                    api_key=os.getenv("ANTHROPIC_API_KEY")
                )

                if emirate == "Abu Dhabi":
                    reg_ref = (
                        "ADOSH-SF Version 4.0, "
                        "relevant ADOSH CoPs"
                    )
                else:
                    reg_ref = (
                        "Dubai Municipality Code "
                        "Article 2.4.5 and relevant chapters"
                    )

                ai_prompt = f"""You are a senior HSE 
investigation expert in {emirate}, UAE.

Analyse this incident and provide a structured report:

INCIDENT TYPE: {incident_type}
CLASSIFICATION: {injury_class}
ACTIVITY: {activity_at_time}
DESCRIPTION: {incident_description}
UNSAFE ACT: {unsafe_act}
UNSAFE CONDITION: {unsafe_condition}
ROOT CAUSES: 
- Personal: {root_personal}
- Job: {root_job}
- Management: {root_management}
REGULATION: {reg_ref}

Provide a detailed analysis with these sections:

## INCIDENT INVESTIGATION FINDINGS

### 1. Sequence of Events
(Step by step what happened)

### 2. Immediate Causes Analysis
(Unsafe acts and conditions identified)

### 3. Root Cause Analysis (5-Why Method)
(Apply 5-Why technique to find true root cause)

### 4. Corrective Actions Required
(Immediate, short-term, and long-term actions)
For each action specify:
- Action required
- Responsible person/role
- Target completion date
- Reference: specific {emirate} regulation or CoP

### 5. Preventive Actions
(What to do to prevent recurrence)

### 6. Regulatory Compliance
(What {emirate} regulations apply and 
reporting obligations)

### 7. Lessons Learned
(Key lessons for all site workers)

Keep analysis practical and cite specific 
{emirate} regulations."""

                response = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=3000,
                    messages=[
                        {
                            "role": "user",
                            "content": ai_prompt
                        }
                    ]
                )

                ai_analysis = response.content[0].text

                # Build data dictionary
                data = {
                    "report_no": report_no,
                    "incident_type": incident_type,
                    "injury_class": injury_class,
                    "incident_date": incident_date,
                    "incident_time": incident_time,
                    "day_of_week": day_of_week,
                    "project_name": project_name,
                    "company": company,
                    "location": location,
                    "reported_by": reported_by,
                    "hse_officer": hse_officer,
                    "report_date": report_date,
                    "injured_name": injured_name,
                    "injured_age": injured_age,
                    "injured_nationality":
                        injured_nationality,
                    "injured_trade": injured_trade,
                    "injured_employer": injured_employer,
                    "injured_experience":
                        injured_experience,
                    "nature_of_injury": nature_of_injury,
                    "body_part": body_part,
                    "medical_treatment":
                        medical_treatment,
                    "hospital": hospital,
                    "incident_description":
                        incident_description,
                    "unsafe_act": unsafe_act,
                    "unsafe_condition": unsafe_condition,
                    "activity_at_time": activity_at_time,
                    "equipment_involved":
                        equipment_involved,
                    "root_personal": root_personal,
                    "root_job": root_job,
                    "root_management": root_management,
                    "witnesses": witnesses,
                }

                docx_buf = build_incident_docx(
                    data, ai_analysis, emirate
                )

                st.success(
                    "✅ Incident Report generated "
                    "successfully!"
                )

                # Show AI analysis preview
                st.divider()
                st.subheader(
                    "📊 AI Investigation Analysis Preview"
                )
                st.markdown(ai_analysis)

                # Download button
                fname = (
                    f"Incident_Report_"
                    f"{report_no}_{incident_date}.docx"
                )
                st.download_button(
                    label=(
                        "⬇️ Download Incident Report (.docx)"
                    ),
                    data=docx_buf,
                    file_name=fname,
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument."
                        "wordprocessingml.document"
                    )
                )

                # Dubai reminder
                if emirate == "Dubai":
                    st.error(
                        "⚠️ REMINDER: Submit written report "
                        "to Dubai Municipality within "
                        "**72 HOURS**. "
                        "DM Emergency: 800900"
                    )

            except Exception as e:
                st.error(f"Error: {str(e)}")
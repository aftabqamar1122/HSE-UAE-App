import streamlit as st
import anthropic
import os
from dotenv import load_dotenv
from datetime import date
from io import BytesIO
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import base64

load_dotenv()

st.set_page_config(
    page_title="Method Statements",
    page_icon="📋"
)

emirate = st.session_state.get("emirate", None)
if not emirate:
    st.warning(
        "Please go to Home and select "
        "your Emirate first."
    )
    st.stop()

color = "🔵" if emirate == "Abu Dhabi" else "🔴"
st.title(
    f"📋 Method Statement Generator "
    f"— {color} {emirate}"
)

if emirate == "Abu Dhabi":
    st.markdown(
        "**Reference: ADOSH-SF v4.0 | "
        "ADOSH CoP 53.1 | Activity-Specific CoPs**"
    )
else:
    st.markdown(
        "**Reference: Dubai Municipality "
        "Code of Construction Safety Practice**"
    )

st.divider()


# ── DOCX HELPERS ──────────────────────────────────────────

def set_cell_color(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def set_borders(cell, color="1F3864"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = OxmlElement('w:tcBorders')
    for side in ['top','left','bottom','right']:
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '6')
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), color)
        borders.append(b)
    tcPr.append(borders)


def section_header(doc, text,
                   fill="1F3864"):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after  = Pt(4)
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill)
    pPr.append(shd)
    run = p.add_run(f"  {text.upper()}")
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(255, 255, 255)


def add_info_row(tbl, label, value,
                 lbl_color="E8F0FE"):
    row = tbl.add_row()
    lc  = row.cells[0]
    vc  = row.cells[1]
    lc.width = Cm(5)
    vc.width = Cm(12)
    lc.text = label
    p = lc.paragraphs[0]
    p.runs[0].bold = True
    p.runs[0].font.size = Pt(10)
    set_cell_color(lc, lbl_color)
    set_borders(lc)
    vc.text = str(value) if value else "—"
    vc.paragraphs[0].runs[0].font.size = Pt(10)
    set_borders(vc)


def add_body(doc, text, size=10,
             bold=False, indent=True):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    if indent:
        p.paragraph_format.left_indent = Cm(0.5)
    run = p.add_run(str(text))
    run.font.size = Pt(size)
    run.bold = bold


def add_bullet(doc, text, size=10):
    p = doc.add_paragraph(
        style='List Bullet'
    )
    run = p.add_run(str(text))
    run.font.size = Pt(size)


def build_ms_docx(meta, content, emirate):
    doc = Document()

    # A4 Portrait
    sec = doc.sections[0]
    sec.page_width    = Cm(21.0)
    sec.page_height   = Cm(29.7)
    sec.left_margin   = Cm(2.0)
    sec.right_margin  = Cm(2.0)
    sec.top_margin    = Cm(2.0)
    sec.bottom_margin = Cm(2.0)

    # ── HEADER BAR ────────────────────────────
    hdr = doc.add_paragraph()
    hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pPr = hdr._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), '1F3864')
    pPr.append(shd)
    r = hdr.add_run(
        f"\n  {meta['company']}"
        f"  —  METHOD STATEMENT  \n"
        f"  {meta['activity']}  \n"
    )
    r.bold = True
    r.font.size = Pt(16)
    r.font.color.rgb = RGBColor(255, 255, 255)

    doc.add_paragraph()

    # ── DOCUMENT INFO TABLE ───────────────────
    section_header(
        doc,
        "1. DOCUMENT INFORMATION",
        fill="1F3864"
    )
    info_tbl = doc.add_table(rows=0, cols=2)
    info_tbl.style = 'Table Grid'
    info_rows = [
        ("Method Statement Title",
         meta['activity']),
        ("Project",          meta['project']),
        ("Location",         meta['location']),
        ("Client",           meta['client']),
        ("Consultant",       meta['consultant']),
        ("Main Contractor",  meta['company']),
        ("Document No.",     meta['doc_no']),
        ("Revision",         meta['revision']),
        ("Date",             str(meta['ms_date'])),
        ("Prepared By",      meta['prepared_by']),
        ("Reviewed By",      meta['reviewed_by']),
        ("Approved By",      meta['approved_by']),
        ("Regulatory Ref.",
         "ADOSH-SF v4.0 | "
         + meta.get('cop_ref','')
         if emirate=="Abu Dhabi"
         else "Dubai Municipality Code | "
         + meta.get('cop_ref','')),
    ]
    for lbl, val in info_rows:
        add_info_row(info_tbl, lbl, val)

    doc.add_paragraph()

    # ── SCOPE ─────────────────────────────────
    section_header(
        doc,
        "2. SCOPE OF WORK",
        fill="1F3864"
    )
    add_body(doc, content.get(
        'scope',
        'Refer to project specifications.'
    ))
    doc.add_paragraph()

    # ── PERSONNEL & RESPONSIBILITIES ──────────
    section_header(
        doc,
        "3. PERSONNEL & RESPONSIBILITIES",
        fill="1F3864"
    )
    resp_tbl = doc.add_table(rows=1, cols=3)
    resp_tbl.style = 'Table Grid'
    for cell, hdr_text in zip(
        resp_tbl.rows[0].cells,
        ["Role", "Name / Title", "Responsibility"]
    ):
        cell.text = hdr_text
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(10)
        cell.paragraphs[0].runs[0].\
            font.color.rgb = RGBColor(255,255,255)
        set_cell_color(cell, "1F3864")
        set_borders(cell)

    personnel = content.get('personnel', [])
    for person in personnel:
        row = resp_tbl.add_row()
        vals = [
            person.get('role',''),
            person.get('name',''),
            person.get('responsibility','')
        ]
        for i, (c, v) in enumerate(
            zip(row.cells, vals)
        ):
            c.text = v
            c.paragraphs[0].runs[0].font.size = Pt(9)
            if i % 2 == 0:
                set_cell_color(c, "F2F2F2")
            set_borders(c)

    doc.add_paragraph()

    # ── PLANT & EQUIPMENT ─────────────────────
    section_header(
        doc,
        "4. PLANT, EQUIPMENT & MATERIALS",
        fill="1F5864"
    )
    equip_tbl = doc.add_table(rows=1, cols=3)
    equip_tbl.style = 'Table Grid'
    for cell, h in zip(
        equip_tbl.rows[0].cells,
        ["Item", "Specification / Type", "Quantity"]
    ):
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(10)
        cell.paragraphs[0].runs[0].\
            font.color.rgb = RGBColor(255,255,255)
        set_cell_color(cell, "1F5864")
        set_borders(cell)

    equipment = content.get('equipment', [])
    for i, item in enumerate(equipment):
        row = equip_tbl.add_row()
        vals = [
            item.get('item',''),
            item.get('spec',''),
            item.get('qty','')
        ]
        for j, (c, v) in enumerate(
            zip(row.cells, vals)
        ):
            c.text = v
            c.paragraphs[0].runs[0].font.size = Pt(9)
            if i % 2 == 0:
                set_cell_color(c, "EEF6FF")
            set_borders(c)

    doc.add_paragraph()

    # ── PPE ───────────────────────────────────
    section_header(
        doc,
        "5. PERSONAL PROTECTIVE EQUIPMENT (PPE)",
        fill="1B5E20"
    )
    ppe_tbl = doc.add_table(rows=1, cols=3)
    ppe_tbl.style = 'Table Grid'
    for cell, h in zip(
        ppe_tbl.rows[0].cells,
        ["PPE Item", "Standard / Reference",
         "Mandatory For"]
    ):
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(10)
        cell.paragraphs[0].runs[0].\
            font.color.rgb = RGBColor(255,255,255)
        set_cell_color(cell, "1B5E20")
        set_borders(cell)

    ppe_items = content.get('ppe', [])
    for i, item in enumerate(ppe_items):
        row = ppe_tbl.add_row()
        vals = [
            item.get('item',''),
            item.get('standard',''),
            item.get('mandatory_for','All workers')
        ]
        for j, (c, v) in enumerate(
            zip(row.cells, vals)
        ):
            c.text = v
            c.paragraphs[0].runs[0].font.size = Pt(9)
            if i % 2 == 0:
                set_cell_color(c, "E8F5E9")
            set_borders(c)

    doc.add_paragraph()

    # ── SEQUENCE OF WORK ──────────────────────
    section_header(
        doc,
        "6. SEQUENCE OF WORK / METHOD",
        fill="BF8F00"
    )
    steps = content.get('sequence', [])
    step_tbl = doc.add_table(rows=1, cols=3)
    step_tbl.style = 'Table Grid'
    for cell, h in zip(
        step_tbl.rows[0].cells,
        ["Step No.", "Activity / Task",
         "Key Safety Requirements"]
    ):
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(10)
        cell.paragraphs[0].runs[0].\
            font.color.rgb = RGBColor(255,255,255)
        set_cell_color(cell, "BF8F00")
        set_borders(cell)

    widths = [Cm(1.5), Cm(7), Cm(8.5)]
    for i, step in enumerate(steps):
        row = step_tbl.add_row()
        vals = [
            str(i+1),
            step.get('task',''),
            step.get('safety_req','')
        ]
        for j, (c, v, w) in enumerate(
            zip(row.cells, vals, widths)
        ):
            c.width = w
            c.text = v
            c.paragraphs[0].runs[0].font.size = Pt(9)
            if i % 2 == 0:
                set_cell_color(c, "FFF9E6")
            set_borders(c)

    doc.add_paragraph()

    # ── RISK SUMMARY ──────────────────────────
    section_header(
        doc,
        "7. HAZARD IDENTIFICATION & RISK SUMMARY",
        fill="C00000"
    )
    hazards = content.get('hazards', [])
    haz_tbl = doc.add_table(rows=1, cols=4)
    haz_tbl.style = 'Table Grid'
    for cell, h in zip(
        haz_tbl.rows[0].cells,
        ["Hazard", "Risk",
         "Controls", "Residual Risk"]
    ):
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(10)
        cell.paragraphs[0].runs[0].\
            font.color.rgb = RGBColor(255,255,255)
        set_cell_color(cell, "C00000")
        set_borders(cell)

    risk_colors = {
        "LOW":      "00AF50",
        "MODERATE": "FFFF00",
        "HIGH":     "FFC000",
        "EXTREME":  "FF0000",
    }
    for i, haz in enumerate(hazards):
        row = haz_tbl.add_row()
        ir = haz.get('initial_risk','HIGH')
        rr = haz.get('residual_risk','LOW')
        vals = [
            haz.get('hazard',''),
            haz.get('risk',''),
            haz.get('controls',''),
            rr
        ]
        for j, (c, v) in enumerate(
            zip(row.cells, vals)
        ):
            c.text = v
            c.paragraphs[0].runs[0].font.size = Pt(9)
            if j == 3:
                set_cell_color(
                    c,
                    risk_colors.get(
                        rr.upper(), "00AF50"
                    )
                )
                c.paragraphs[0].runs[0].bold = True
            elif i % 2 == 0:
                set_cell_color(c, "FFF2F2")
            set_borders(c)

    doc.add_paragraph()

    # ── EMERGENCY ─────────────────────────────
    section_header(
        doc,
        "8. EMERGENCY PROCEDURES",
        fill="C00000"
    )
    if emirate == "Abu Dhabi":
        emergency_text = (
            "In the event of any emergency:\n"
            "1. STOP WORK immediately and raise "
            "the alarm\n"
            "2. Call Emergency Services: "
            "Ambulance 998 | Police 999 | "
            "Civil Defence 997\n"
            "3. Notify Site HSE Officer: "
            f"{meta.get('prepared_by','HSE Officer')}\n"
            "4. Evacuate to designated Muster Point\n"
            "5. Do not re-enter until area declared safe\n"
            "6. Report to ADOSH as required by "
            "ADOSH-SF v4.0"
        )
    else:
        emergency_text = (
            "In the event of any emergency:\n"
            "1. STOP WORK immediately and raise "
            "the alarm\n"
            "2. Call Emergency Services: "
            "Police 999 | Civil Defence 997 | "
            "Ambulance 998 | DM Emergency 800900\n"
            "3. Notify Site HSE Officer: "
            f"{meta.get('prepared_by','HSE Officer')}\n"
            "4. Evacuate to designated Muster Point\n"
            "5. Report to Dubai Municipality "
            "within 72 hours (DM Code Art. 2.4.5)"
        )
    for line in emergency_text.split('\n'):
        if line.strip():
            add_bullet(doc, line.strip())

    doc.add_paragraph()

    # ── PERMITS REQUIRED ──────────────────────
    section_header(
        doc,
        "9. PERMITS REQUIRED",
        fill="4A148C"
    )
    permits = content.get('permits', [])
    if permits:
        ptw_tbl = doc.add_table(rows=1, cols=3)
        ptw_tbl.style = 'Table Grid'
        for cell, h in zip(
            ptw_tbl.rows[0].cells,
            ["Permit Type", "Required When",
             "Reference"]
        ):
            cell.text = h
            cell.paragraphs[0].runs[0].bold = True
            cell.paragraphs[0].runs[0].font.size = Pt(10)
            cell.paragraphs[0].runs[0].\
                font.color.rgb = RGBColor(255,255,255)
            set_cell_color(cell, "4A148C")
            set_borders(cell)

        for i, p in enumerate(permits):
            row = ptw_tbl.add_row()
            vals = [
                p.get('type',''),
                p.get('when',''),
                p.get('ref','')
            ]
            for j, (c, v) in enumerate(
                zip(row.cells, vals)
            ):
                c.text = v
                c.paragraphs[0].runs[0].font.size = Pt(9)
                if i % 2 == 0:
                    set_cell_color(c, "F3E5F5")
                set_borders(c)

    doc.add_paragraph()

    # ── INSPECTION CHECKLIST ──────────────────
    section_header(
        doc,
        "10. PRE-WORK INSPECTION CHECKLIST",
        fill="0D47A1"
    )
    checks = content.get('checklist', [])
    chk_tbl = doc.add_table(rows=1, cols=3)
    chk_tbl.style = 'Table Grid'
    for cell, h in zip(
        chk_tbl.rows[0].cells,
        ["Inspection Item", "Checked By",
         "Status"]
    ):
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(10)
        cell.paragraphs[0].runs[0].\
            font.color.rgb = RGBColor(255,255,255)
        set_cell_color(cell, "0D47A1")
        set_borders(cell)

    for i, item in enumerate(checks):
        row = chk_tbl.add_row()
        vals = [item, "HSE Officer", "☐ Yes  ☐ No"]
        for j, (c, v) in enumerate(
            zip(row.cells, vals)
        ):
            c.text = v
            c.paragraphs[0].runs[0].font.size = Pt(9)
            if i % 2 == 0:
                set_cell_color(c, "E8F0FE")
            set_borders(c)

    doc.add_paragraph()

    # ── REFERENCES ────────────────────────────
    section_header(
        doc,
        "11. REFERENCES & STANDARDS",
        fill="1F3864"
    )
    refs = content.get('references', [])
    for ref in refs:
        add_bullet(doc, ref)

    doc.add_paragraph()

    # ── SIGNATURES ────────────────────────────
    section_header(
        doc,
        "12. SIGNATURES & APPROVAL",
        fill="1F3864"
    )
    sig_tbl = doc.add_table(rows=3, cols=3)
    sig_tbl.style = 'Table Grid'
    for i, role in enumerate([
        "Prepared By\n(HSE Officer)",
        "Reviewed By\n(Project Manager)",
        "Approved By\n(Projects Director)"
    ]):
        c = sig_tbl.rows[0].cells[i]
        c.text = role
        c.paragraphs[0].runs[0].bold = True
        c.paragraphs[0].runs[0].font.size = Pt(10)
        c.paragraphs[0].runs[0].\
            font.color.rgb = RGBColor(255,255,255)
        set_cell_color(c, "1F3864")
        set_borders(c)
    for i in range(3):
        c = sig_tbl.rows[1].cells[i]
        c.text = "\n\nName: _______________\n"
        c.paragraphs[0].runs[0].font.size = Pt(10)
        set_borders(c)
        c = sig_tbl.rows[2].cells[i]
        c.text = (
            "Date: _______________\n\n"
            "Signature: _______________"
        )
        c.paragraphs[0].runs[0].font.size = Pt(10)
        set_borders(c)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ════════════════════════════════════════════════════════
# ACTIVITIES LIST
# ════════════════════════════════════════════════════════

activities = {
    "🏗️ Civil & Structural": [
        "Excavation & Earthworks",
        "Dewatering Operations",
        "Foundation Works & Piling",
        "Concrete Pouring & Pumping",
        "Formwork & Falsework (Shuttering)",
        "Reinforcement Steel Fixing",
        "Precast Concrete Installation",
        "Masonry & Blockwork",
        "Steel Erection & Structural Steel",
        "Demolition Works",
        "Underground Construction",
        "Waterproofing Works",
    ],
    "🏗️ Finishing Works": [
        "Plastering & Rendering",
        "Tiling & Stone Works",
        "Painting & Decorating",
        "False Ceiling Installation",
        "Flooring Installation",
        "Glazing & Curtain Wall",
        "Carpentry & Joinery",
    ],
    "⚡ MEP Works": [
        "Electrical Installation Works",
        "Plumbing & Drainage Works",
        "HVAC Installation",
        "Fire Fighting System Installation",
        "Cable Pulling & Termination",
        "Generator Installation",
        "Solar Panel Installation",
    ],
    "🏗️ Temporary Works": [
        "Scaffolding Erection & Dismantling",
        "Hoarding & Fencing Installation",
        "Site Accommodation Setup",
        "Crane Erection & Dismantling",
        "Traffic Diversion & Management",
        "Temporary Electrical Installation",
    ],
    "⚠️ High Risk Activities": [
        "Working at Heights (>1.8m)",
        "Confined Space Entry",
        "Hot Work (Welding & Cutting)",
        "Abrasive Blasting",
        "Chemical Handling & Storage",
        "Asbestos Removal",
        "Radiography / NDT Works",
        "LOTO (Lockout-Tagout)",
        "Lifting Operations (Crane & Rigging)",
        "Overhead Line Work",
        "Underground Services Works",
        "Spray Painting & Finishing",
    ],
    "🚛 Logistics & Transport": [
        "Material Delivery & Unloading",
        "Heavy Equipment Mobilization",
        "Material Handling & Storage",
        "Waste Management & Disposal",
    ],
}


# ════════════════════════════════════════════════════════
# MAIN UI — TWO TABS
# ════════════════════════════════════════════════════════

tab_gen, tab_upload = st.tabs([
    "⚡ Generate Method Statement (AI)",
    "📤 Upload Your Own Format"
])


# ════════════════════════════════════════════════════════
# TAB 1 — AI GENERATOR
# ════════════════════════════════════════════════════════

with tab_gen:
    st.markdown(
        "### 🤖 AI-Powered Method Statement"
    )
    st.info(
        "Select any activity below. The AI will "
        "generate a complete, professional Method "
        f"Statement referencing {emirate} "
        "regulations — ready to submit."
    )

    # Activity selector
    cat = st.selectbox(
        "Select Category",
        list(activities.keys()),
        key="ms_category"
    )
    activity = st.selectbox(
        "Select Activity",
        activities[cat],
        key="ms_activity"
    )

    # Or type custom
    custom = st.text_input(
        "Or type a custom activity "
        "(leave blank to use selection above)",
        placeholder="e.g. Pile Cap Construction",
        key="ms_custom"
    )
    final_activity = custom if custom else activity

    st.divider()

    # Project info
    st.subheader("📋 Project Information")
    c1, c2 = st.columns(2)
    with c1:
        ms_project = st.text_input(
            "Project Name",
            placeholder="e.g. Bloom Living Almeria",
            key="ms_proj"
        )
        ms_location = st.text_input(
            "Location",
            placeholder="e.g. Zayed City, Abu Dhabi",
            key="ms_loc"
        )
        ms_client = st.text_input(
            "Client",
            placeholder="e.g. Bloom District",
            key="ms_client"
        )
        ms_consultant = st.text_input(
            "Consultant",
            placeholder="e.g. Dar Al-Handasah",
            key="ms_cons"
        )
    with c2:
        ms_company = st.text_input(
            "Main Contractor",
            value="ENGC",
            key="ms_co"
        )
        ms_doc_no = st.text_input(
            "Document No.",
            placeholder="e.g. ENGC-MS-EXC-001",
            key="ms_docno"
        )
        ms_rev = st.selectbox(
            "Revision",
            ["Rev 00","Rev 01","Rev 02","Rev 03"],
            key="ms_rev"
        )
        ms_date = st.date_input(
            "Date",
            value=date.today(),
            key="ms_date"
        )

    c3, c4 = st.columns(2)
    with c3:
        ms_prep = st.text_input(
            "Prepared By",
            placeholder="e.g. Aftab Qamar — "
                        "Senior HSSE Engineer",
            key="ms_prep"
        )
        ms_review = st.text_input(
            "Reviewed By",
            placeholder="e.g. Ahmad Abdelrahman "
                        "— Project Manager",
            key="ms_review"
        )
    with c4:
        ms_approve = st.text_input(
            "Approved By",
            placeholder="e.g. Rami Kamal — "
                        "Projects Director",
            key="ms_approve"
        )
        ms_workers = st.number_input(
            "No. of Workers for this Activity",
            min_value=1,
            max_value=500,
            value=20,
            key="ms_workers"
        )

    ms_special = st.text_area(
        "Special Conditions / Notes "
        "(optional)",
        placeholder="e.g. Adjacent to live road, "
                    "night works required, "
                    "restricted access...",
        height=70,
        key="ms_special"
    )

    if st.button(
        f"⚡ Generate Method Statement for "
        f"'{final_activity}' (.docx)",
        type="primary",
        use_container_width=True,
        key="ms_gen_btn"
    ):
        if not ms_project:
            st.error(
                "Please enter a project name."
            )
        else:
            prog = st.progress(
                0, text="Starting..."
            )

            try:
                client_api = anthropic.Anthropic(
                    api_key=os.getenv(
                        "ANTHROPIC_API_KEY"
                    )
                )

                if emirate == "Abu Dhabi":
                    reg = (
                        "ADOSH-SF v4.0, relevant "
                        "ADOSH CoPs (cite specific "
                        "CoP numbers)"
                    )
                else:
                    reg = (
                        "Dubai Municipality Code "
                        "of Construction Safety "
                        "Practice (cite chapters)"
                    )

                prog.progress(
                    10, text="Generating scope..."
                )

                import json

                prompt = f"""You are a senior HSE 
engineer in {emirate}, UAE.

Generate a complete Method Statement for:
ACTIVITY: {final_activity}
PROJECT: {ms_project}
LOCATION: {ms_location}
CLIENT: {ms_client}
CONSULTANT: {ms_consultant}
CONTRACTOR: {ms_company}
WORKERS: {ms_workers}
SPECIAL CONDITIONS: {ms_special}
REGULATION: {reg}

Return ONLY valid JSON (no markdown, no text).
Start with {{ end with }}

{{
  "cop_ref": "specific CoP or Chapter ref",
  "scope": "2-3 sentence scope paragraph",
  "personnel": [
    {{"role": "Site Engineer",
      "name": "TBD",
      "responsibility": "Supervise works"}},
    {{"role": "HSE Officer",
      "name": "{ms_prep or 'Aftab Qamar'}",
      "responsibility": "HSE oversight"}},
    {{"role": "Foreman",
      "name": "TBD",
      "responsibility": "Direct workers"}},
    {{"role": "Competent Operator",
      "name": "TBD",
      "responsibility": "Operate equipment"}}
  ],
  "equipment": [
    {{"item": "Equipment name",
      "spec": "Type/capacity",
      "qty": "1"}},
    ... 5-8 items total
  ],
  "ppe": [
    {{"item": "Safety Helmet",
      "standard": "EN 397 / ANSI Z89.1",
      "mandatory_for": "All personnel"}},
    {{"item": "Safety Footwear",
      "standard": "EN ISO 20345",
      "mandatory_for": "All personnel"}},
    {{"item": "High-Vis Vest",
      "standard": "EN ISO 20471",
      "mandatory_for": "All personnel"}},
    ... 6-10 PPE items total specific to activity
  ],
  "sequence": [
    {{"task": "Pre-work inspection and briefing",
      "safety_req": "Toolbox talk conducted. "
        "PTW issued if required."}},
    ... 8-12 steps total
  ],
  "hazards": [
    {{"hazard": "Specific hazard",
      "risk": "Injury type",
      "controls": "Control measures with CoP ref",
      "initial_risk": "HIGH",
      "residual_risk": "LOW"}},
    ... 6-10 hazards
  ],
  "permits": [
    {{"type": "Permit name",
      "when": "When required",
      "ref": "CoP or Chapter ref"}},
    ... list all required permits
  ],
  "checklist": [
    "All personnel briefed on Method Statement",
    "PTW obtained if required",
    "All equipment inspected and certified",
    ... 10-15 checklist items
  ],
  "references": [
    "ADOSH-SF Version 4.0 (July 2024)",
    "ADOSH CoP XX.X — relevant CoP title",
    ... 5-8 references
  ]
}}

Make all content SPECIFIC to {final_activity}
in {emirate}. Cite exact regulation references.
Return ONLY the JSON object."""

                prog.progress(
                    20, text="AI generating..."
                )

                response = client_api.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=4000,
                    messages=[{
                        "role": "user",
                        "content": prompt
                    }]
                )

                prog.progress(
                    70,
                    text="Parsing response..."
                )

                raw = (
                    response.content[0].text.strip()
                )
                # Clean JSON
                import re
                raw = re.sub(
                    r'```json|```', '', raw
                ).strip()
                start = raw.find('{')
                end   = raw.rfind('}') + 1
                raw   = raw[start:end]

                content_data = json.loads(raw)

                prog.progress(
                    85,
                    text="Building Word document..."
                )

                meta = {
                    "activity":    final_activity,
                    "project":     ms_project,
                    "location":    ms_location,
                    "client":      ms_client,
                    "consultant":  ms_consultant,
                    "company":     ms_company,
                    "doc_no":      ms_doc_no,
                    "revision":    ms_rev,
                    "ms_date":     ms_date,
                    "prepared_by": ms_prep,
                    "reviewed_by": ms_review,
                    "approved_by": ms_approve,
                    "cop_ref": content_data.get(
                        'cop_ref', ''
                    ),
                }

                docx_buf = build_ms_docx(
                    meta,
                    content_data,
                    emirate
                )

                prog.progress(
                    100,
                    text="✅ Method Statement ready!"
                )

                st.success(
                    f"✅ Method Statement for "
                    f"'{final_activity}' generated!"
                )

                fname = (
                    f"MS_{final_activity"
                    f".replace(' ','_')}"
                    f"_{ms_project.replace(' ','_')}"
                    f"_{ms_date}.docx"
                )
                st.download_button(
                    label=(
                        f"⬇️ Download Method "
                        f"Statement — "
                        f"{final_activity} (.docx)"
                    ),
                    data=docx_buf,
                    file_name=fname,
                    mime=(
                        "application/vnd"
                        ".openxmlformats-officedocument"
                        ".wordprocessingml.document"
                    ),
                    key="ms_download"
                )

            except json.JSONDecodeError as e:
                st.error(
                    f"JSON error: {e}. "
                    "Please try again."
                )
            except Exception as e:
                st.error(f"Error: {str(e)}")


# ════════════════════════════════════════════════════════
# TAB 2 — UPLOAD YOUR OWN FORMAT
# ════════════════════════════════════════════════════════

with tab_upload:
    st.markdown(
        "### 📤 Upload Your Own Format"
    )
    st.info(
        "Upload your company's existing Method "
        "Statement format (Word or PDF). "
        "The AI will fill it with content for "
        "your selected activity and project — "
        "keeping your exact format, layout, "
        "and branding."
    )

    uploaded_file = st.file_uploader(
        "Upload Your Method Statement "
        "Template / Sample",
        type=["docx", "pdf"],
        help="Upload a blank template or a "
             "previously completed Method "
             "Statement to use as the format "
             "reference.",
        key="ms_upload"
    )

    if uploaded_file:
        st.success(
            f"✅ File uploaded: "
            f"**{uploaded_file.name}** "
            f"({uploaded_file.size:,} bytes)"
        )
        st.info(
            "📋 **What happens next:**\n\n"
            "1. The AI reads your uploaded format\n"
            "2. Extracts your layout, headings, "
            "table structure, and style\n"
            "3. Fills it with content for your "
            "selected activity\n"
            "4. Returns a completed document in "
            "YOUR format\n\n"
            "Select your activity and project "
            "details below, then click Generate."
        )

        st.divider()

        u_activity = st.text_input(
            "Activity for this Method Statement",
            placeholder="e.g. Excavation Works",
            key="upload_activity"
        )
        u_project = st.text_input(
            "Project Name",
            placeholder="e.g. Bloom Living Almeria",
            key="upload_project"
        )
        u_company = st.text_input(
            "Company Name",
            value="ENGC",
            key="upload_company"
        )

        if st.button(
            "⚡ Generate Using My Format",
            type="primary",
            use_container_width=True,
            key="upload_gen_btn"
        ):
            with st.spinner(
                "Reading your format and "
                "generating content..."
            ):
                try:
                    client_api = anthropic.Anthropic(
                        api_key=os.getenv(
                            "ANTHROPIC_API_KEY"
                        )
                    )

                    file_content = (
                        uploaded_file.read()
                    )
                    b64 = base64.standard_b64encode(
                        file_content
                    ).decode()

                    if uploaded_file.name\
                            .endswith('.pdf'):
                        media_type = (
                            "application/pdf"
                        )
                        doc_type = "document"
                    else:
                        # For DOCX, extract text
                        from io import BytesIO
                        doc_temp = Document(
                            BytesIO(file_content)
                        )
                        extracted = "\n".join([
                            p.text
                            for p in doc_temp.paragraphs
                            if p.text.strip()
                        ])

                        response = (
                            client_api.messages.create(
                            model="claude-sonnet-4-6",
                            max_tokens=3000,
                            messages=[{
                                "role": "user",
                                "content": (
                                    f"Here is a Method "
                                    f"Statement format:\n\n"
                                    f"{extracted}\n\n"
                                    f"Now fill this format "
                                    f"with content for:\n"
                                    f"Activity: {u_activity}\n"
                                    f"Project: {u_project}\n"
                                    f"Company: {u_company}\n"
                                    f"Emirate: {emirate}\n\n"
                                    f"Keep the EXACT same "
                                    f"structure, headings, and "
                                    f"sections. Just fill in "
                                    f"the content with specific "
                                    f"{emirate} regulatory "
                                    f"references."
                                )
                            }]
                        ))

                        st.success(
                            "✅ Format analyzed and "
                            "content generated!"
                        )
                        st.markdown(
                            "### Generated Content "
                            "(in your format):"
                        )
                        st.markdown(
                            response.content[0].text
                        )
                        st.download_button(
                            "⬇️ Download as Text",
                            data=(
                                response.content[0].text
                            ),
                            file_name=(
                                f"MS_{u_activity"
                                f".replace(' ','_')}"
                                "_filled.txt"
                            ),
                            mime="text/plain",
                            key="upload_dl"
                        )

                except Exception as e:
                    st.error(f"Error: {str(e)}")

    else:
        # Show tips
        st.markdown(
            "#### 💡 Tips for Best Results:"
        )
        st.markdown("""
- Upload a **Word (.docx)** file for best
  format preservation
- Your template can be blank (just the
  structure) or a completed example
- The AI will match your headings, table
  layouts, and section order exactly
- All content will be Emirates-specific
  with correct regulatory references
""")
        st.markdown(
            "#### 📋 Don't Have a Template?"
        )
        st.info(
            "Use the **'Generate Method "
            "Statement (AI)'** tab on the left "
            "to create a professional Method "
            "Statement from scratch — ready "
            "for immediate use."
        )
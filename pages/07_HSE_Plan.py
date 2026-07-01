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

load_dotenv()

st.set_page_config(
    page_title="HSE Plan Generator",
    page_icon="📄"
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
    f"📄 HSE Plan Generator — {color} {emirate}"
)

if emirate == "Abu Dhabi":
    st.markdown(
        "**Reference: ADOSH CoP 53.1 — "
        "OSH Construction Management Plan | "
        "ADOSH-SF Version 4.0**"
    )
    st.info(
        "This generates a full OSH Construction "
        "Management Plan as required by "
        "ADOSH CoP 53.1 before work commences."
    )
else:
    st.markdown(
        "**Reference: Dubai Municipality "
        "Code Art. 2.3.2 — Safety Plan**"
    )
    st.info(
        "This generates a Safety Plan as "
        "required by Dubai Municipality before "
        "construction commences."
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


def set_borders(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = OxmlElement('w:tcBorders')
    for side in ['top','left','bottom','right']:
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '6')
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), '1F3864')
        borders.append(b)
    tcPr.append(borders)


def add_heading(doc, text, level=1,
                color_hex="1F3864"):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after  = Pt(4)
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    pPr.append(shd)
    run = p.add_run(f"  {text}")
    run.bold = True
    run.font.size = Pt(13 if level==1 else 11)
    run.font.color.rgb = RGBColor(255,255,255)
    return p


def add_info_row(tbl, label, value,
                 label_color="E8F0FE"):
    row = tbl.add_row()
    lc  = row.cells[0]
    vc  = row.cells[1]
    lc.width = Cm(5)
    vc.width = Cm(12)
    lc.text  = label
    lc.paragraphs[0].runs[0].bold = True
    lc.paragraphs[0].runs[0].font.size = Pt(10)
    set_cell_color(lc, label_color)
    set_borders(lc)
    vc.text  = str(value) if value else "—"
    vc.paragraphs[0].runs[0].font.size = Pt(10)
    set_borders(vc)


def add_body_text(doc, text, size=10,
                  bold=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.left_indent = Cm(0.5)
    run = p.add_run(str(text))
    run.font.size = Pt(size)
    run.bold = bold
    return p


def add_bullet(doc, text, size=10):
    p = doc.add_paragraph(
        style='List Bullet'
    )
    run = p.add_run(text)
    run.font.size = Pt(size)
    return p


def build_hse_plan_docx(meta, sections,
                         emirate):
    doc = Document()

    # A4 Portrait
    sec = doc.sections[0]
    sec.page_width    = Cm(21.0)
    sec.page_height   = Cm(29.7)
    sec.left_margin   = Cm(2.5)
    sec.right_margin  = Cm(2.0)
    sec.top_margin    = Cm(2.0)
    sec.bottom_margin = Cm(2.0)

    # ── COVER PAGE ────────────────────────────
    cover = doc.add_paragraph()
    cover.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pPr = cover._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), '1F3864')
    pPr.append(shd)
    r = cover.add_run(
        "\n\n\n"
        f"{'🔵' if emirate=='Abu Dhabi' else '🔴'}"
        f"  {meta['company']}\n\n"
    )
    r.bold = True
    r.font.size = Pt(22)
    r.font.color.rgb = RGBColor(255, 255, 255)

    doc_title = (
        "OSH CONSTRUCTION MANAGEMENT PLAN"
        if emirate == "Abu Dhabi"
        else "CONSTRUCTION SAFETY PLAN"
    )

    cover2 = doc.add_paragraph()
    cover2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pPr2 = cover2._p.get_or_add_pPr()
    shd2 = OxmlElement('w:shd')
    shd2.set(qn('w:val'), 'clear')
    shd2.set(qn('w:color'), 'auto')
    shd2.set(qn('w:fill'), 'C00000')
    pPr2.append(shd2)
    r2 = cover2.add_run(
        f"\n  {doc_title}  \n\n"
        f"  {meta['project']}  \n\n"
    )
    r2.bold = True
    r2.font.size = Pt(18)
    r2.font.color.rgb = RGBColor(255, 255, 255)

    doc.add_page_break()

    # ── DOCUMENT CONTROL TABLE ────────────────
    add_heading(doc,
                "DOCUMENT CONTROL INFORMATION",
                color_hex="1F3864")
    ctrl_tbl = doc.add_table(rows=0, cols=2)
    ctrl_tbl.style = 'Table Grid'

    ctrl_rows = [
        ("Document Title",       doc_title),
        ("Project Name",         meta['project']),
        ("Project Location",     meta['location']),
        ("Client",               meta['client']),
        ("Consultant / Engineer",meta['consultant']),
        ("Main Contractor",      meta['company']),
        ("Contract No.",         meta['contract_no']),
        ("Document No.",         meta['doc_no']),
        ("Revision",             meta['revision']),
        ("Prepared By",          meta['prepared_by']),
        ("Approved By",          meta['approved_by']),
        ("Date",                 str(meta['doc_date'])),
        ("Regulatory Reference",
         "ADOSH CoP 53.1 | ADOSH-SF v4.0"
         if emirate == "Abu Dhabi"
         else "Dubai Municipality Code Art.2.3.2"),
    ]
    for lbl, val in ctrl_rows:
        add_info_row(ctrl_tbl, lbl, val)

    doc.add_paragraph()
    doc.add_page_break()

    # ── TABLE OF CONTENTS ─────────────────────
    add_heading(doc, "TABLE OF CONTENTS",
                color_hex="1F3864")
    toc_items = [
        "1.0  Introduction & Project Overview",
        "2.0  Scope of Work",
        "3.0  HSE Policy Statement",
        "4.0  Regulatory & Legal Requirements",
        "5.0  HSE Organisation & Responsibilities",
        "6.0  Hazard Identification & Risk Assessment",
        "7.0  Emergency Response Plan",
        "8.0  Permit to Work System",
        "9.0  HSE Training & Competency",
        "10.0 Personal Protective Equipment",
        "11.0 Occupational Health & Worker Welfare",
        "12.0 Environmental Management",
        "13.0 Incident Reporting & Investigation",
        "14.0 HSE Inspections & Audits",
        "15.0 Sub-contractor Management",
        "16.0 Traffic & Logistics Management",
        "17.0 Communication & Consultation",
        "18.0 Document Control",
        "19.0 Performance Monitoring & KPIs",
        "20.0 Signatures & Approvals",
    ]
    for item in toc_items:
        add_bullet(doc, item)
    doc.add_page_break()

    # ── MAIN SECTIONS FROM AI ─────────────────
    section_colors = {
        "1":  "1F3864",
        "2":  "1F3864",
        "3":  "C00000",
        "4":  "1F3864",
        "5":  "1F5864",
        "6":  "BF8F00",
        "7":  "C00000",
        "8":  "C00000",
        "9":  "1F3864",
        "10": "1F3864",
        "11": "1B5E20",
        "12": "1B5E20",
        "13": "C00000",
        "14": "1F3864",
        "15": "1F3864",
        "16": "1F3864",
        "17": "1F3864",
        "18": "1F3864",
        "19": "1F3864",
        "20": "1F3864",
    }

    for section in sections:
        num   = section.get("number", "")
        title = section.get("title", "")
        body  = section.get("body", "")
        color = section_colors.get(
            str(num).split(".")[0], "1F3864"
        )
        add_heading(
            doc,
            f"{num}  {title.upper()}",
            color_hex=color
        )
        if isinstance(body, list):
            for item in body:
                if item.startswith("•") or \
                   item.startswith("-"):
                    add_bullet(
                        doc,
                        item.lstrip("•- ").strip()
                    )
                else:
                    add_body_text(doc, item)
        else:
            paragraphs = body.split("\n")
            for para in paragraphs:
                para = para.strip()
                if not para:
                    continue
                if para.startswith("•") or \
                   para.startswith("-"):
                    add_bullet(
                        doc,
                        para.lstrip("•- ").strip()
                    )
                elif para.startswith("**"):
                    add_body_text(
                        doc,
                        para.strip("*"),
                        bold=True
                    )
                else:
                    add_body_text(doc, para)

        doc.add_paragraph()

    # ── SIGNATURE BLOCK ───────────────────────
    add_heading(doc, "20.0  SIGNATURES & APPROVAL",
                color_hex="1F3864")
    sig_tbl = doc.add_table(rows=3, cols=3)
    sig_tbl.style = 'Table Grid'
    roles = [
        "Prepared By\n(HSE Officer)",
        "Reviewed By\n(Project Manager)",
        "Approved By\n(Projects Director)"
    ]
    for i, role in enumerate(roles):
        cell = sig_tbl.rows[0].cells[i]
        cell.text = role
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(10)
        set_cell_color(cell, "1F3864")
        cell.paragraphs[0].runs[0].\
            font.color.rgb = RGBColor(255,255,255)
        set_borders(cell)
    for i in range(3):
        cell = sig_tbl.rows[1].cells[i]
        cell.text = "\n\nName: _______________\n"
        cell.paragraphs[0].runs[0].font.size = Pt(10)
        set_borders(cell)
    for i in range(3):
        cell = sig_tbl.rows[2].cells[i]
        cell.text = "Date: _______________\n\nSignature: _______________"
        cell.paragraphs[0].runs[0].font.size = Pt(10)
        set_borders(cell)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ── INPUT FORM ────────────────────────────────────────────
st.subheader("📋 Project Information")

c1, c2 = st.columns(2)
with c1:
    project   = st.text_input(
        "Project Name",
        placeholder="e.g. Bloom Living Almeria"
    )
    location  = st.text_input(
        "Project Location",
        placeholder="e.g. Zayed City, Abu Dhabi"
    )
    client    = st.text_input(
        "Client",
        placeholder="e.g. Bloom District Properties"
    )
    consultant = st.text_input(
        "Consultant / Engineer",
        placeholder="e.g. Dar Al-Handasah"
    )

with c2:
    company   = st.text_input(
        "Main Contractor",
        value="ENGC"
    )
    contract_no = st.text_input(
        "Contract Number",
        placeholder="e.g. BL-2024-001"
    )
    doc_no    = st.text_input(
        "Document Number",
        placeholder="e.g. ENGC-HSE-PLAN-001"
    )
    revision  = st.selectbox(
        "Revision",
        ["Rev 00","Rev 01","Rev 02","Rev 03"]
    )

st.divider()
c3, c4 = st.columns(2)
with c3:
    prepared_by = st.text_input(
        "Prepared By",
        placeholder="e.g. Aftab Qamar — "
                    "Senior HSSE Engineer"
    )
    approved_by = st.text_input(
        "Approved By",
        placeholder="e.g. Rami Kamal — "
                    "Projects Director"
    )
with c4:
    doc_date = st.date_input(
        "Document Date",
        value=date.today()
    )
    project_value = st.text_input(
        "Approximate Project Value",
        placeholder="e.g. AED 50 Million"
    )

st.divider()
st.subheader("🏗️ Project Scope Details")

c5, c6 = st.columns(2)
with c5:
    project_type = st.multiselect(
        "Project Type",
        [
            "Residential Villas",
            "Community Center",
            "Service Station",
            "Commercial Building",
            "Infrastructure Works",
            "Fit-out Works",
            "Demolition",
            "MEP Works",
        ],
        default=["Residential Villas"]
    )
    peak_workers = st.number_input(
        "Peak Workforce (number of workers)",
        min_value=10,
        max_value=5000,
        value=150
    )

with c6:
    project_duration = st.text_input(
        "Project Duration",
        placeholder="e.g. 24 months "
                    "(Jan 2025 – Dec 2026)"
    )
    high_risk_activities = st.multiselect(
        "High-Risk Activities on Site",
        [
            "Excavation Works",
            "Working at Heights",
            "Scaffolding Erection",
            "Lifting Operations (Crane)",
            "Precast Installation",
            "Formwork & Falsework",
            "Hot Work (Welding/Cutting)",
            "Electrical Works",
            "Confined Space Entry",
            "Demolition Works",
            "Underground Services",
            "Traffic Management",
            "Chemical Handling",
            "Concrete Pumping",
        ],
        default=[
            "Excavation Works",
            "Working at Heights",
            "Lifting Operations (Crane)",
            "Scaffolding Erection",
        ]
    )

st.divider()
st.subheader("📌 Additional Information")
c7, c8 = st.columns(2)
with c7:
    hse_manager = st.text_input(
        "Project HSE Manager / Officer",
        placeholder="e.g. Aftab Qamar"
    )
    pm_name = st.text_input(
        "Project Manager",
        placeholder="e.g. Ahmad Abdelrahman"
    )
with c8:
    pd_name = st.text_input(
        "Projects Director",
        placeholder="e.g. Rami Kamal Yassin"
    )
    site_address = st.text_input(
        "Site Address / Plot Number",
        placeholder="e.g. Plot C13, Zayed City"
    )

special_notes = st.text_area(
    "Special Site Conditions / Notes "
    "(optional)",
    placeholder="e.g. Site adjacent to live road, "
                "groundwater expected at 2m depth...",
    height=80
)

# ── GENERATE ──────────────────────────────────────────────
st.divider()
if st.button(
    "⚡ Generate Complete HSE Plan (.docx)",
    type="primary",
    use_container_width=True
):
    if not project:
        st.error("Please enter a project name.")
    else:
        prog = st.progress(0,
            text="Starting HSE Plan generation...")

        try:
            client_api = anthropic.Anthropic(
                api_key=os.getenv("ANTHROPIC_API_KEY")
            )

            if emirate == "Abu Dhabi":
                reg_ref = (
                    "ADOSH-SF Version 4.0 (July 2024), "
                    "ADOSH CoP 53.1 (OSH Construction "
                    "Management Plan), ADOSH CoP 53.0, "
                    "MOHRE Resolution 44/2022"
                )
                auth = "ADOSH / ADPHC"
            else:
                reg_ref = (
                    "Dubai Municipality Code of "
                    "Construction Safety Practice, "
                    "Articles 2.3.2, 2.4.5, "
                    "Local Order 61/1991"
                )
                auth = "Dubai Municipality"

            activities_str = ", ".join(
                high_risk_activities
            )
            project_type_str = ", ".join(
                project_type
            )

            sections_to_generate = [
                ("1.0", "Introduction & Project Overview"),
                ("2.0", "Scope of Work"),
                ("3.0", "HSE Policy Statement"),
                ("4.0", "Regulatory & Legal Requirements"),
                ("5.0", "HSE Organisation & Responsibilities"),
                ("6.0", "Hazard Identification & Risk Assessment"),
                ("7.0", "Emergency Response Plan"),
                ("8.0", "Permit to Work System"),
                ("9.0", "HSE Training & Competency"),
                ("10.0", "Personal Protective Equipment"),
                ("11.0", "Occupational Health & Worker Welfare"),
                ("12.0", "Environmental Management"),
                ("13.0", "Incident Reporting & Investigation"),
                ("14.0", "HSE Inspections & Audits"),
                ("15.0", "Sub-contractor Management"),
                ("16.0", "Traffic & Logistics Management"),
                ("17.0", "Communication & Consultation"),
                ("18.0", "Document Control"),
                ("19.0", "Performance Monitoring & KPIs"),
            ]

            all_sections = []
            total = len(sections_to_generate)

            for i, (num, title) in enumerate(
                sections_to_generate
            ):
                prog.progress(
                    int((i / total) * 90),
                    text=f"Generating Section {num} "
                         f"— {title}..."
                )

                prompt = f"""You are a senior HSE 
professional writing a formal
{('OSH Construction Management Plan' if emirate=='Abu Dhabi' else 'Construction Safety Plan')}
for {emirate}, UAE.

PROJECT DETAILS:
- Project: {project}
- Location: {location}
- Client: {client}
- Consultant: {consultant}
- Main Contractor: {company}
- Project Type: {project_type_str}
- Duration: {project_duration}
- Peak Workers: {peak_workers}
- High-Risk Activities: {activities_str}
- HSE Officer: {hse_manager}
- Project Manager: {pm_name}
- Special Notes: {special_notes}

REGULATORY REFERENCE: {reg_ref}
AUTHORITY: {auth}

Write SECTION {num} — {title.upper()}
for this specific project.

Requirements:
1. Write professionally as a REAL HSE Plan
   (not a template)
2. Reference specific {emirate} regulations
   throughout
3. Use actual project details provided
4. Include specific procedures, responsibilities,
   and requirements
5. For Abu Dhabi: cite specific ADOSH CoP numbers
   For Dubai: cite specific DM Code chapters
6. Make it comprehensive — this will be submitted
   to the consultant/authority
7. Format with clear paragraphs and bullet points
8. Length: 250-400 words for this section

Write ONLY the section content.
No section number or title in your response
(those are added separately).
Start directly with the content."""

                response = client_api.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=1000,
                    messages=[{
                        "role": "user",
                        "content": prompt
                    }]
                )

                all_sections.append({
                    "number": num,
                    "title":  title,
                    "body":   response.content[0].text
                })

            prog.progress(92,
                text="Building Word document...")

            meta = {
                "project":     project,
                "location":    location,
                "client":      client,
                "consultant":  consultant,
                "company":     company,
                "contract_no": contract_no,
                "doc_no":      doc_no,
                "revision":    revision,
                "prepared_by": prepared_by,
                "approved_by": approved_by,
                "doc_date":    doc_date,
                "project_value": project_value,
                "hse_manager": hse_manager,
                "pm_name":     pm_name,
                "pd_name":     pd_name,
            }

            docx_buf = build_hse_plan_docx(
                meta, all_sections, emirate
            )

            prog.progress(100,
                text="✅ HSE Plan complete!")

            st.success(
                f"✅ Complete HSE Plan generated! "
                f"{len(all_sections)} sections — "
                f"Ready for submission to "
                f"{'consultant' if emirate=='Abu Dhabi' else 'Dubai Municipality'}"
            )

            # Section preview
            with st.expander(
                "📋 Preview Generated Sections",
                expanded=False
            ):
                for s in all_sections:
                    st.markdown(
                        f"### {s['number']} "
                        f"{s['title']}"
                    )
                    st.markdown(s['body'])
                    st.divider()

            fname = (
                f"HSE_Plan_{project.replace(' ','_')}"
                f"_{emirate.replace(' ','_')}"
                f"_{doc_date}.docx"
            )

            st.download_button(
                label=(
                    "⬇️ Download Complete "
                    "HSE Plan (.docx)"
                ),
                data=docx_buf,
                file_name=fname,
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument."
                    "wordprocessingml.document"
                )
            )

            if emirate == "Abu Dhabi":
                st.info(
                    "📋 **Next Steps (Abu Dhabi):**\n\n"
                    "1. Review and customize this plan "
                    "for your specific site conditions\n\n"
                    "2. Submit to Consultant "
                    "(Dar Al-Handasah) for review\n\n"
                    "3. Consultant submits to ADOSH "
                    "if required\n\n"
                    "4. Keep approved copy on site "
                    "at all times\n\n"
                    "**Reference: ADOSH CoP 53.1**"
                )
            else:
                st.info(
                    "📋 **Next Steps (Dubai):**\n\n"
                    "1. Engineer reviews and approves "
                    "the Safety Plan\n\n"
                    "2. Submit to Dubai Municipality "
                    "if required\n\n"
                    "3. Safety meetings every 15 days "
                    "as per DM Code Art. 2.2.8\n\n"
                    "**Reference: DM Code Art. 2.3.2**"
                )

        except Exception as e:
            prog.progress(0)
            st.error(f"Error: {str(e)}")
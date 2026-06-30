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
import json

load_dotenv()

st.set_page_config(page_title="Risk Assessment", page_icon="📋")

emirate = st.session_state.get("emirate", None)
if not emirate:
    st.warning("Please go to Home and select your Emirate first.")
    st.stop()

color = "🔵" if emirate == "Abu Dhabi" else "🔴"
st.title(f"📋 Risk Assessment — {color} {emirate}")
st.markdown(
    "**Format: ENGC 14-Column | ADOSH-SF v4.0 | "
    "Hierarchy of Controls**"
)
st.divider()

# ── DOCX Helpers ─────────────────────────────────────────

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


def risk_color(score):
    s = int(score)
    if s >= 16: return "FF0000"    # Extreme
    if s >= 10: return "FFC000"    # High
    if s >= 5:  return "FFFF00"    # Moderate
    return "00AF50"                 # Low


def risk_label(score):
    s = int(score)
    if s >= 16: return "Extreme Risk"
    if s >= 10: return "High Risk"
    if s >= 5:  return "Moderate Risk"
    return "Low Risk"


def add_cell_text(cell, text, bold=False,
                  size=9, center=False, color_rgb=None):
    cell.text = ""
    para = cell.paragraphs[0]
    if center:
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.add_run(str(text))
    run.bold = bold
    run.font.size = Pt(size)
    if color_rgb:
        run.font.color.rgb = RGBColor(*color_rgb)


def build_docx(meta, rows):
    doc = Document()

    # A4 Landscape
    sec = doc.sections[0]
    sec.page_width    = Cm(29.7)
    sec.page_height   = Cm(21.0)
    sec.left_margin   = Cm(1.27)
    sec.right_margin  = Cm(1.27)
    sec.top_margin    = Cm(1.27)
    sec.bottom_margin = Cm(1.27)

    # ── Company Header ────────────────────────────────────
    hdr = doc.add_paragraph()
    hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = hdr.add_run(
        f"{meta['company']} — RISK ASSESSMENT"
    )
    r.bold = True
    r.font.size = Pt(14)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = sub.add_run(
        f"Project: {meta['project']}   |   "
        f"Activity: {meta['activity']}   |   "
        f"Date: {meta['date']}   |   "
        f"Prepared by: {meta['prepared_by']}   |   "
        f"Location: {meta['location']}   |   "
        f"Emirate: {meta['emirate']}"
    )
    r2.font.size = Pt(9)

    doc.add_paragraph()

    # ── 14-Column Table ───────────────────────────────────
    # Column widths in cm (converted from ENGC DXA values)
    # DXA values: [524,950,1582,1830,1206,
    #              552,634,633,966,3286,
    #              567,603,729,1336]
    col_w = [
        Cm(0.93), Cm(1.68), Cm(2.80), Cm(3.23), Cm(2.13),
        Cm(0.98), Cm(1.12), Cm(1.12), Cm(1.71), Cm(5.81),
        Cm(1.00), Cm(1.07), Cm(1.29), Cm(2.36)
    ]

    tbl = doc.add_table(rows=0, cols=14)
    tbl.style = 'Table Grid'

    # ── HEADER ROW 1 ──────────────────────────────────────
    r1 = tbl.add_row()
    # Merged header labels and their column spans
    h1_items = [
        ("S/N", 1, "C0C0C0"),
        ("Activity\nElement", 1, "C0C0C0"),
        ("Hazards /\nImpact", 1, "C0C0C0"),
        ("Who Might Be\nHarmed and How?", 1, "C0C0C0"),
        ("Risk & Potential\nConsequences", 1, "C0C0C0"),
        ("Risk\nClassification", 3, "C0C0C0"),
        ("Initial Risk Level\nLow/Moderate/High/Extreme",
         1, "C0C0C0"),
        (
            "Action to be Taken to Reduce the Risk\n"
            "(Abide by the '*Hierarchy of Safety Controls')\n"
            "(Include reference to corresponding legal "
            "and other requirements)",
            1, "C0C0C0"
        ),
        ("Revised Risk\nClassification", 3, "C0C0C0"),
        ("Residual Risk Level\nLow/Moderate/High/Extreme",
         1, "C0C0C0"),
    ]

    col_cursor = 0
    for text, span, fill in h1_items:
        cell = r1.cells[col_cursor]
        cell.text = ""
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(text)
        run.bold = True
        run.font.size = Pt(8)
        set_cell_color(cell, fill)
        set_borders(cell)
        if span > 1:
            cell.merge(r1.cells[col_cursor + span - 1])
        col_cursor += span

    # ── HEADER ROW 2 ──────────────────────────────────────
    r2 = tbl.add_row()
    h2_labels = [
        "S/N", "Activity\nElement",
        "Hazards /\nImpact",
        "Who Might Be\nHarmed and How?",
        "Risk & Potential\nConsequences",
        "Probability\n(P)", "Severity\n(S)",
        "Risk Rating\n(R) P X S",
        "Initial Risk Level\nLow/Moderate/\nHigh/Extreme",
        "Action to be Taken to Reduce the Risk",
        "Probability\n(P)", "Severity\n(S)",
        "Risk Rating\n(R) P X S",
        "Residual Risk Level\nLow/Moderate/\nHigh/Extreme",
    ]
    for idx, label in enumerate(h2_labels):
        cell = r2.cells[idx]
        cell.text = ""
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run(label)
        run.bold = True
        run.font.size = Pt(7)
        set_cell_color(cell, "BFBFBF")
        set_borders(cell)
        cell.width = col_w[idx]

    # ── DATA ROWS ─────────────────────────────────────────
    for i, row_data in enumerate(rows, start=1):
        ip = int(row_data.get("initial_p", 3))
        is_ = int(row_data.get("initial_s", 3))
        ir = ip * is_
        rp = int(row_data.get("residual_p", 1))
        rs = int(row_data.get("residual_s", 3))
        rr = rp * rs

        # Build controls text with hierarchy
        controls_text = row_data.get("controls", {})
        if isinstance(controls_text, dict):
            ctrl_lines = []
            order = [
                ("elimination", "ELIMINATION"),
                ("substitution", "SUBSTITUTION"),
                ("engineering", "ENGINEERING CONTROLS"),
                ("administrative", "ADMINISTRATIVE CONTROLS"),
                ("ppe", "PPE"),
            ]
            for key, label in order:
                items = controls_text.get(key, [])
                if items:
                    ctrl_lines.append(f"{label}:")
                    for item in items:
                        ctrl_lines.append(f"• {item}")
            controls_str = "\n".join(ctrl_lines)
        else:
            controls_str = str(controls_text)

        row = tbl.add_row()
        cells = row.cells

        # Set column widths
        for idx in range(14):
            cells[idx].width = col_w[idx]

        # Col 0: S/N
        add_cell_text(cells[0], str(i),
                      bold=True, size=9, center=True)
        set_borders(cells[0])

        # Col 1: Activity Element
        add_cell_text(cells[1],
                      row_data.get("activity", ""),
                      bold=True, size=9)
        set_borders(cells[1])

        # Col 2: Hazards
        add_cell_text(cells[2],
                      row_data.get("hazard", ""), size=9)
        set_borders(cells[2])

        # Col 3: Who harmed
        add_cell_text(cells[3],
                      row_data.get("who_harmed", ""),
                      size=9)
        set_borders(cells[3])

        # Col 4: Consequences
        add_cell_text(cells[4],
                      row_data.get("consequences", ""),
                      size=9)
        set_borders(cells[4])

        # Col 5: Initial P
        add_cell_text(cells[5], str(ip),
                      bold=True, size=10, center=True)
        set_borders(cells[5])

        # Col 6: Initial S
        add_cell_text(cells[6], str(is_),
                      bold=True, size=10, center=True)
        set_borders(cells[6])

        # Col 7: Initial R
        add_cell_text(cells[7], str(ir),
                      bold=True, size=10, center=True)
        set_cell_color(cells[7], risk_color(ir))
        set_borders(cells[7])

        # Col 8: Initial Risk Level
        cells[8].text = ""
        para8 = cells[8].paragraphs[0]
        para8.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run8 = para8.add_run(risk_label(ir))
        run8.bold = True
        run8.font.size = Pt(9)
        set_cell_color(cells[8], risk_color(ir))
        set_borders(cells[8])

        # Col 9: Controls (with hierarchy)
        cells[9].text = ""
        para9 = cells[9].paragraphs[0]
        for line in controls_str.split("\n"):
            if para9.runs:
                para9 = cells[9].add_paragraph()
            run9 = para9.add_run(line)
            run9.font.size = Pt(8)
            if line.endswith(":") and line.isupper():
                run9.bold = True
                run9.font.color.rgb = RGBColor(192, 0, 0)
        set_borders(cells[9])

        # Col 10: Residual P
        add_cell_text(cells[10], str(rp),
                      bold=True, size=10, center=True)
        set_borders(cells[10])

        # Col 11: Residual S
        add_cell_text(cells[11], str(rs),
                      bold=True, size=10, center=True)
        set_borders(cells[11])

        # Col 12: Residual R
        add_cell_text(cells[12], str(rr),
                      bold=True, size=10, center=True)
        set_cell_color(cells[12], risk_color(rr))
        set_borders(cells[12])

        # Col 13: Residual Risk Level
        cells[13].text = ""
        para13 = cells[13].paragraphs[0]
        para13.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run13 = para13.add_run(risk_label(rr))
        run13.bold = True
        run13.font.size = Pt(9)
        set_cell_color(cells[13], risk_color(rr))
        set_borders(cells[13])

    # Save
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ── INPUT FORM ────────────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    project_name = st.text_input(
        "Project Name",
        placeholder="e.g. Bloom Living Almeria"
    )
    company_name = st.text_input(
        "Company Name", value="ENGC"
    )
    prepared_by = st.text_input(
        "Prepared By",
        placeholder="e.g. Aftab Qamar — Senior HSSE Engineer"
    )

with col2:
    activity = st.text_input(
        "Work Activity",
        placeholder="e.g. Excavation Work"
    )
    ra_date = st.date_input("Date", value=date.today())
    location = st.text_input(
        "Location",
        placeholder="e.g. Plot C13, Zayed City"
    )

st.divider()
work_description = st.text_area(
    "Describe the work activity",
    placeholder="e.g. Excavation for foundation to 2.5m...",
    height=80
)
workers_involved = st.text_input(
    "Who might be harmed?",
    placeholder="e.g. Excavation crew, site workers, public"
)

# ── GENERATE ──────────────────────────────────────────────
if st.button(
    "⚡ Generate Risk Assessment (.docx)",
    type="primary",
    use_container_width=True
):
    if not activity:
        st.error("Please enter a work activity.")
    else:
        with st.spinner("Generating..."):
            try:
                client = anthropic.Anthropic(
                    api_key=os.getenv("ANTHROPIC_API_KEY")
                )

                if emirate == "Abu Dhabi":
                    reg = (
                        "ADOSH-SF v4.0 and ADOSH Codes of "
                        "Practice (cite specific CoP numbers)"
                    )
                else:
                    reg = (
                        "Dubai Municipality Code of "
                        "Construction Safety Practice "
                        "(cite specific chapters)"
                    )

                prompt = f"""You are a senior HSE professional 
in {emirate}, UAE.

Generate a Risk Assessment for:
Activity: {activity}
Project: {project_name}
Description: {work_description}
Who harmed: {workers_involved}
Regulation: {reg}

Return ONLY a valid JSON array with exactly 6 objects.
No markdown, no explanation, just raw JSON.

Each object must have these EXACT keys:
{{
  "activity": "specific task step",
  "hazard": "specific hazard description",
  "consequences": "what injury/damage could occur",
  "who_harmed": "who is at risk and how",
  "initial_p": 4,
  "initial_s": 4,
  "controls": {{
    "elimination": ["list of elimination measures"],
    "substitution": ["list of substitution measures"],
    "engineering": [
      "engineering control 1 (ref: ADOSH CoP XX.X)",
      "engineering control 2"
    ],
    "administrative": [
      "admin control 1 (ref: ADOSH CoP XX.X)",
      "Permit to Work required (CoP 21.0)"
    ],
    "ppe": [
      "PPE item 1 (ref: ADOSH CoP 2.0)",
      "PPE item 2"
    ]
  }},
  "residual_p": 1,
  "residual_s": 3
}}

IMPORTANT RULES:
1. Controls MUST follow hierarchy:
   Elimination → Substitution → Engineering → 
   Administrative → PPE
2. Every engineering/admin control MUST reference 
   the specific {emirate} regulation or ADOSH CoP number
3. initial_p x initial_s should give HIGH or EXTREME risk
4. residual_p x residual_s should give LOW or MODERATE risk
5. Return ONLY the JSON array starting with [ 
   and ending with ]
"""

                response = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=4000,
                    messages=[
                        {"role": "user", "content": prompt}
                    ]
                )

                raw = response.content[0].text.strip()

                # Clean if wrapped in markdown
                if "```" in raw:
                    raw = raw.split("```")[1]
                    if raw.startswith("json"):
                        raw = raw[4:]
                raw = raw.strip()

                # Find JSON array
                start = raw.find("[")
                end = raw.rfind("]") + 1
                raw = raw[start:end]

                ra_rows = json.loads(raw)

                meta = {
                    "project":     project_name,
                    "company":     company_name,
                    "prepared_by": prepared_by,
                    "date":        ra_date,
                    "activity":    activity,
                    "location":    location,
                    "emirate":     emirate,
                }

                docx_buf = build_docx(meta, ra_rows)

                st.success(
                    "✅ Risk Assessment generated in ENGC "
                    "14-Column Format with Hierarchy of Controls!"
                )

                # Preview
                import pandas as pd
                preview = []
                for i, r in enumerate(ra_rows, 1):
                    ip = int(r.get("initial_p", 3))
                    is_ = int(r.get("initial_s", 3))
                    rp = int(r.get("residual_p", 1))
                    rs_ = int(r.get("residual_s", 3))
                    preview.append({
                        "#": i,
                        "Activity": r.get("activity", ""),
                        "Hazard": r.get("hazard", ""),
                        "Who Harmed": r.get("who_harmed",""),
                        "P": ip,
                        "S": is_,
                        "R": ip * is_,
                        "Initial Level": risk_label(ip*is_),
                        "P2": rp,
                        "S2": rs_,
                        "R2": rp * rs_,
                        "Residual Level": risk_label(rp*rs_),
                    })
                st.dataframe(
                    pd.DataFrame(preview),
                    use_container_width=True
                )

                fname = (
                    f"RA_{activity.replace(' ','_')}"
                    f"_{ra_date}_{emirate}.docx"
                )
                st.download_button(
                    label="⬇️ Download Risk Assessment (.docx)",
                    data=docx_buf,
                    file_name=fname,
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument."
                        "wordprocessingml.document"
                    )
                )

            except json.JSONDecodeError as e:
                st.error(
                    f"JSON parse error: {e}\n"
                    "Please try again."
                )
            except Exception as e:
                st.error(f"Error: {str(e)}")
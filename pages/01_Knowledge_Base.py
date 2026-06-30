import streamlit as st
import anthropic
import os
from dotenv import load_dotenv

load_dotenv()

st.set_page_config(
    page_title="HSE Knowledge Base", page_icon="📚"
)

emirate = st.session_state.get("emirate", None)
if not emirate:
    st.warning(
        "Please go to Home and select your Emirate first."
    )
    st.stop()

color = "🔵" if emirate == "Abu Dhabi" else "🔴"
st.title(f"📚 HSE Knowledge Base — {color} {emirate}")

if emirate == "Abu Dhabi":
    st.markdown(
        "**Source: ADOSH-SF Version 4.0 | ADPHC**"
    )
else:
    st.markdown(
        "**Source: Dubai Municipality Code of "
        "Construction Safety Practice**"
    )

st.divider()

# ── TAB LAYOUT ────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs([
    "🔍 AI Search",
    "📋 Browse All Topics",
    "⚡ Quick Reference"
])

# ════════════════════════════════════════════════════════
# TAB 1 — AI SEARCH
# ════════════════════════════════════════════════════════
with tab1:
    st.subheader(f"🔍 Ask Anything — {color} {emirate}")
    st.markdown(
        "Type any HSE topic, activity, or regulation — "
        "get an instant Emirates-specific answer."
    )

    search = st.text_input(
        "Search",
        placeholder=(
            "e.g. NDT Test, excavation permit, "
            "scaffolding inspection, heat stress..."
        ),
        label_visibility="collapsed"
    )

    if search:
        with st.spinner("Searching knowledge base..."):
            try:
                client = anthropic.Anthropic(
                    api_key=os.getenv("ANTHROPIC_API_KEY")
                )

                if emirate == "Abu Dhabi":
                    reg_context = """
Abu Dhabi Regulatory Framework:
- Authority: ADOSH (Abu Dhabi OSH) / ADPHC
- Framework: ADOSH-SF Version 4.0 (July 2024)
- Key CoPs: 2.0 PPE, 11.0 Heat, 15.0 Electrical,
  20.0 Safety in Design, 21.0 Permit to Work,
  23.0 Working at Heights, 24.0 LOTO,
  26.0 Scaffolding, 27.0 Confined Spaces,
  28.0 Hot Work, 29.0 Excavation,
  34.0 Lifting Equipment, 35.0 Power Tools,
  36.0 Plant & Equipment, 39.0 Services,
  40.0 Formwork, 42.0 Precast, 44.0 Traffic,
  53.0 OSH Management, 53.1 Construction Plan,
  54.0 Waste Management
- Summer ban: 12:30-15:00, 15 Jun-15 Sep
- Emergency: Police 999, Ambulance 998,
  Civil Defence 997
"""
                else:
                    reg_context = """
Dubai Regulatory Framework:
- Authority: Dubai Municipality (DM)
- Code: Code of Construction Safety Practice
- Key Chapters: Ch.2 General Safety,
  Ch.3 Occupational Health, Ch.4 PPE,
  Ch.5 Fire, Ch.6 Signs & Barricades,
  Ch.7 Material Handling, Ch.8 Scaffolding,
  Ch.9 Excavation, Ch.10 Concrete,
  Ch.11 Steel Erection, Ch.13 Demolition,
  Ch.14 Hand Tools, Ch.15 Welding,
  Ch.16 Electrical, Ch.17 LOTO,
  Ch.18 Chemical Hazards, Ch.19 Confined Spaces,
  Ch.20 Road Works, Ch.21 Rigging,
  Ch.22 Cranes
- Emergency: Police 999, Civil Defence 997,
  DM Emergency 800900, Ambulance 998
- Accident reporting: within 72 hours to DM
- Records retention: 5 years
"""

                prompt = f"""You are an expert HSE knowledge 
base for {emirate}, UAE construction sites.

{reg_context}

User question: {search}

Provide a comprehensive, well-structured answer that includes:

1. **Direct Answer** — what this is / what it means
2. **{emirate} Regulatory Requirement** — specific CoP or 
   Chapter reference
3. **Key Requirements** — bullet points of main obligations
4. **Practical Steps** — what to do on site
5. **Responsible Persons** — who is responsible
6. **Common Mistakes** — what to avoid
7. **Related Topics** — other relevant CoPs/Chapters

Always label your answer: {color} {emirate} SPECIFIC
Always cite specific regulation references.
Be practical and useful for HSE engineers on site."""

                response = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=2000,
                    messages=[
                        {"role": "user", "content": prompt}
                    ]
                )

                answer = response.content[0].text
                st.markdown(answer)

                # Save to search history
                if "search_history" not in st.session_state:
                    st.session_state.search_history = []
                if search not in st.session_state.search_history:
                    st.session_state.search_history.insert(
                        0, search
                    )
                    # Keep only last 10
                    st.session_state.search_history = \
                        st.session_state.search_history[:10]

            except Exception as e:
                st.error(f"Error: {str(e)}")

    # Recent searches
    if st.session_state.get("search_history"):
        st.divider()
        st.markdown("**Recent Searches:**")
        cols = st.columns(5)
        for i, term in enumerate(
            st.session_state.search_history[:5]
        ):
            if cols[i % 5].button(
                term, key=f"hist_{i}",
                use_container_width=True
            ):
                st.session_state.search_term = term
                st.rerun()

# ════════════════════════════════════════════════════════
# TAB 2 — BROWSE ALL TOPICS
# ════════════════════════════════════════════════════════
with tab2:
    st.subheader(f"📋 All Topics — {color} {emirate}")

    if emirate == "Abu Dhabi":
        # Abu Dhabi CoP Categories
        categories = {
            "🏗️ Construction Critical": {
                "CoP 20.0": "Safety in Design (Construction)",
                "CoP 21.0": "Permit to Work Systems",
                "CoP 23.0": "Working at Heights",
                "CoP 26.0": "Scaffolding",
                "CoP 29.0": "Excavation Work",
                "CoP 34.0": "Lifting Equipment & Accessories",
                "CoP 40.0": "False Work (Formwork)",
                "CoP 42.0": "Pre-Cast Construction",
                "CoP 53.0": "OSH Management — Construction",
                "CoP 53.1": "OSH Construction Management Plan",
            },
            "⚡ Hazardous Activities": {
                "CoP 15.0": "Electrical Safety",
                "CoP 24.0": "Lock-out / Tag-out (LOTO)",
                "CoP 27.0": "Confined Spaces",
                "CoP 28.0": "Hot Work Operations",
                "CoP 39.0": "Overhead & Underground Services",
                "CoP 44.0": "Traffic Management & Logistics",
                "CoP 46.0": "Underground Construction",
            },
            "🌡️ Health & Welfare": {
                "CoP 2.0":  "Personal Protective Equipment",
                "CoP 4.0":  "First Aid & Medical Emergency",
                "CoP 5.0":  "Health Screening & Surveillance",
                "CoP 8.0":  "General Workplace Amenities",
                "CoP 11.0": "Safety in the Heat",
                "CoP 14.0": "Manual Handling & Ergonomics",
                "CoP 17.0": "Safety Signage & Signals",
                "CoP 18.0": "Employer Supplied Accommodation",
            },
            "🔧 Plant & Equipment": {
                "CoP 35.0": "Portable Power Tools",
                "CoP 36.0": "Plant and Equipment",
                "CoP 37.0": "Ladders",
                "CoP 38.0": "Concrete Placing Equipment",
                "CoP 41.0": "Steel Erection",
                "CoP 47.0": "Machine Guarding",
                "CoP 51.0": "Powered Lift Trucks",
            },
            "🌿 Environment & Waste": {
                "CoP 1.0":  "Hazardous Materials",
                "CoP 1.1":  "Asbestos Management",
                "CoP 3.0":  "Occupational Noise",
                "CoP 52.0": "Local Exhaust Ventilation",
                "CoP 54.0": "Waste Management",
            },
        }
    else:
        # Dubai Chapters
        categories = {
            "🏗️ Construction Critical": {
                "Chapter 2":  "General Health & Safety",
                "Chapter 8":  "Scaffolding Safety",
                "Chapter 9":  "Excavation",
                "Chapter 10": "Concrete and Masonry",
                "Chapter 11": "Steel Erection",
                "Chapter 12": "Underground Construction",
                "Chapter 21": "Sling & Rigging Equipment",
                "Chapter 22": "Cranes",
            },
            "⚡ Hazardous Activities": {
                "Chapter 15": "Welding and Cutting",
                "Chapter 16": "Electrical Hazards",
                "Chapter 17": "Control of Hazardous Energy",
                "Chapter 18": "Chemical Hazards & Labeling",
                "Chapter 19": "Confined Spaces",
                "Chapter 20": "Road Works & Site Transport",
            },
            "🌡️ Health & Welfare": {
                "Chapter 3": "Occupational Health & Environment",
                "Chapter 4": "Personal Protective Equipment",
                "Chapter 5": "Fire Protection & Prevention",
                "Chapter 6": "Signs, Signals & Barricades",
                "Chapter 13": "Demolition",
                "Chapter 14": "Hand & Power-Operated Tools",
            },
            "📋 Management": {
                "Chapter 1": "Preliminary — Definitions & Scope",
                "Chapter 7": "Material Handling & Storage",
                "Chapter 23": "References",
            },
        }

    # Display categories
    for cat_name, topics in categories.items():
        with st.expander(cat_name, expanded=False):
            for ref, title in topics.items():
                col_a, col_b = st.columns([1, 3])
                col_a.markdown(f"**{ref}**")
                col_b.markdown(title)
                # Quick look button
                if st.button(
                    f"📖 Ask AI about {ref}",
                    key=f"btn_{ref}",
                    use_container_width=False
                ):
                    with st.spinner(f"Loading {ref}..."):
                        try:
                            client = anthropic.Anthropic(
                                api_key=os.getenv(
                                    "ANTHROPIC_API_KEY"
                                )
                            )
                            q = (
                                f"Explain {ref} — {title} "
                                f"requirements in {emirate}"
                            )
                            resp = client.messages.create(
                                model="claude-sonnet-4-6",
                                max_tokens=1000,
                                messages=[
                                    {
                                        "role": "user",
                                        "content": q
                                    }
                                ]
                            )
                            st.info(resp.content[0].text)
                        except Exception as e:
                            st.error(str(e))
                st.divider()

# ════════════════════════════════════════════════════════
# TAB 3 — QUICK REFERENCE
# ════════════════════════════════════════════════════════
with tab3:
    st.subheader(f"⚡ Quick Reference — {color} {emirate}")

    if emirate == "Abu Dhabi":
        st.markdown("### 🚨 Emergency Numbers")
        c1, c2, c3 = st.columns(3)
        c1.metric("Police", "999")
        c2.metric("Ambulance", "998")
        c3.metric("Civil Defence", "997")

        st.divider()
        st.markdown("### ☀️ Summer Midday Work Ban")
        st.error(
            "**NO OUTDOOR WORK: 12:30 PM — 3:00 PM**\n\n"
            "Period: **15 June to 15 September** (Annual)\n\n"
            "Reference: MOHRE Resolution No. 44/2022 | "
            "ADOSH CoP 11.0"
        )

        st.divider()
        st.markdown("### 📊 Risk Rating Matrix")
        st.markdown("""
| Rating | Score | Color | Action |
|---|---|---|---|
| **EXTREME** | 16–25 | 🔴 Red | Stop work immediately |
| **HIGH** | 10–15 | 🟠 Orange | Senior management approval |
| **MODERATE** | 5–9 | 🟡 Yellow | Management attention |
| **LOW** | 1–4 | 🟢 Green | Routine controls |
        """)

        st.divider()
        st.markdown("### 🪜 Hierarchy of Controls")
        st.markdown("""
1. 🔴 **ELIMINATION** — Remove the hazard completely
2. 🟠 **SUBSTITUTION** — Replace with something safer
3. 🟡 **ENGINEERING CONTROLS** — Isolate people from hazard
4. 🔵 **ADMINISTRATIVE CONTROLS** — Change the way people work
5. 🟢 **PPE** — Protect the individual (last resort)
        """)

        st.divider()
        st.markdown("### 📋 Key Permit to Work Types")
        st.markdown("""
| Permit Type | CoP Reference |
|---|---|
| Hot Work | CoP 21.0 + CoP 28.0 |
| Confined Space Entry | CoP 21.0 + CoP 27.0 |
| Excavation | CoP 21.0 + CoP 29.0 |
| Working at Heights | CoP 21.0 + CoP 23.0 |
| Electrical Isolation (LOTO) | CoP 21.0 + CoP 15.0 |
| Lifting Operations | CoP 21.0 + CoP 34.0 |
| Radiography / NDT | CoP 21.0 |
        """)

        st.divider()
        st.markdown("### 📏 Key Dimensions to Remember")
        st.markdown("""
| Item | Measurement |
|---|---|
| Excavation permit trigger | > 0.5m depth |
| Ladder access in excavation | > 1.2m depth |
| Spoil distance from edge | Minimum 0.6m |
| Guardrail height | Minimum 950mm |
| Toe board height | Minimum 150mm |
| Fall arrest anchor strength | Minimum 10kN |
| Critical lift threshold | > 75% SWL or > 10T |
| Scaffold inspection interval | Every 7 days |
        """)

    else:
        st.markdown("### 🚨 Emergency Numbers (Dubai)")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Police", "999")
        c2.metric("Civil Defence", "997")
        c3.metric("Ambulance", "998")
        c4.metric("DM Emergency", "800900")

        st.divider()
        st.markdown("### ⏱️ Key Deadlines")
        st.error(
            "**Accident Report to Dubai Municipality: "
            "Within 72 HOURS**\n\n"
            "Records to be kept: **5 YEARS**\n\n"
            "Reference: DM Code Article 2.4.5"
        )

        st.divider()
        st.markdown("### 👷 Safety Staff Requirements")
        st.markdown("""
**Contractor On-Site (Table 4):**

| Workers | Safety Inspector | Safety Officer | Senior SO |
|---|---|---|---|
| < 50 | 1 (part-time) | — | — |
| 50–150 | 1 | — | — |
| 151–500 | 1 | 1 | — |
| 501–1000 | 1 | 1 | 1 |
| 1001–1250 | 2 | 1 | 1 |
| 1251–1500 | 2 | 2 | 1 |
| 1501–2000 | 2 | 2 | 2 |
        """)

        st.divider()
        st.markdown("### 📏 Key Dubai Dimensions")
        st.markdown("""
| Item | Measurement |
|---|---|
| Excavation shoring trigger | > 1.25m depth |
| Ladder access in excavation | > 1.2m depth |
| Ladder extends above surface | ≥ 90cm |
| Spoil distance from edge | Minimum 60cm |
| Max travel distance to ladder | 15 metres |
| O₂ range for confined space | 19.5% — 23.5% |
| Combustibles limit (confined space) | < 10% LEL |
| Ventilation rate (confined space) | 20 air changes/hr |
| Eye wash station max distance | 30 metres |
| Crane inspection interval | Every 12 months |
        """)

        st.divider()
        st.markdown("### 🏆 Safety Staff Qualifications")
        st.markdown("""
| Role | Education | Experience |
|---|---|---|
| Senior Safety Officer | University degree (Engineering) | 5+ years |
| Safety Officer | Diploma (Engineering) | 3+ years |
| Safety Inspector | Secondary / Technical | 2+ years |

*All must have DM-approved qualification certificate*
        """)
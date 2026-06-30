import streamlit as st
import anthropic
import os
from dotenv import load_dotenv

load_dotenv()

st.set_page_config(page_title="HSE AI Chatbot", page_icon="🤖")

# Check emirate selected
emirate = st.session_state.get("emirate", None)
if not emirate:
    st.warning("Please go to the Home page and select your Emirate first.")
    st.stop()

# Colour coding
color = "🔵" if emirate == "Abu Dhabi" else "🔴"
st.title(f"🤖 HSE AI Assistant — {color} {emirate}")
st.caption(f"All answers are specific to {emirate} regulations")

# System prompt — the brain of your chatbot
SYSTEM_PROMPT = f"""You are an expert HSE (Health, Safety and Environment) advisor 
specialising in UAE construction compliance, specifically for {emirate}.

CRITICAL RULES:
1. ALL answers must be specific to {emirate} ONLY. Never mix Abu Dhabi and Dubai rules.
2. Always cite the specific regulation (Abu Dhabi: ADOSH CoP number; Dubai: DM Code Chapter/Article number).
3. Always label your answer with {color} {emirate} SPECIFIC at the top.
4. Structure answers: Summary → Key Requirements → Legal Reference → Practical Steps.
5. If asked about the other emirate, say "Please switch to [other emirate] mode on the home page."
6. For Abu Dhabi: Reference ADOSH-SF Version 4.0 (July 2024) and ADOSH CoPs.
   Key CoPs: PPE=2.0, Heat=11.0, PTW=21.0, Heights=23.0, Scaffold=26.0, 
   Confined Space=27.0, Hot Work=28.0, Excavation=29.0, Lifting=34.0, 
   Traffic=44.0, Construction Management Plan=53.1, Waste=54.0.
   IMPORTANT: Authority is ADOSH (NOT OSHAD). Forms are ADOSH-SF (NOT OSHAD-SF).
   Summer midday ban: 12:30-15:00, 15 June to 15 September (MOHRE Resolution 44/2022).
7. For Dubai: Reference Dubai Municipality Code of Construction Safety Practice.
   Key chapters: Ch.2=General Safety, Ch.3=Occupational Health, Ch.4=PPE, 
   Ch.5=Fire, Ch.8=Scaffolding, Ch.9=Excavation, Ch.16=Electrical, 
   Ch.17=LOTO, Ch.19=Confined Spaces, Ch.22=Cranes.
   Emergency numbers: Police 999, Civil Defence 997, DM Emergency 800900.
   Accident reports: within 72 hours to Dubai Municipality. Records: 5 years.
8. Always end with: ⚠️ Verify this information against current official publications 
   before operational use.
"""

# Initialize chat history
if "messages" not in st.session_state:
    st.session_state.messages = []

# Display chat history
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

# Quick question buttons
st.markdown("**Quick Questions:**")
q_col1, q_col2, q_col3 = st.columns(3)
quick_questions = {
    "abu_dhabi": [
        "What is the midday work ban in Abu Dhabi?",
        "What permits do I need for excavation?",
        "How many safety officers do I need on site?"
    ],
    "dubai": [
        "What are the excavation requirements in Dubai?",
        "How many safety staff do I need for 300 workers?",
        "What is the accident reporting deadline in Dubai?"
    ]
}

emirate_key = "abu_dhabi" if emirate == "Abu Dhabi" else "dubai"
questions = quick_questions[emirate_key]

if q_col1.button(questions[0], use_container_width=True):
    st.session_state.pending_question = questions[0]
if q_col2.button(questions[1], use_container_width=True):
    st.session_state.pending_question = questions[1]
if q_col3.button(questions[2], use_container_width=True):
    st.session_state.pending_question = questions[2]

# Get user input
user_input = st.chat_input(f"Ask any {emirate} HSE question...")

# Handle quick questions
if "pending_question" in st.session_state:
    user_input = st.session_state.pending_question
    del st.session_state.pending_question

# Process input
if user_input:
    # Add user message
    st.session_state.messages.append({"role": "user", "content": user_input})
    with st.chat_message("user"):
        st.markdown(user_input)

    # Call Anthropic API
    with st.chat_message("assistant"):
        with st.spinner("Checking regulations..."):
            try:
                client = anthropic.Anthropic(
                    api_key=os.getenv("ANTHROPIC_API_KEY")
                )
                
                response = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=1500,
                    system=SYSTEM_PROMPT,
                    messages=[
                        {"role": m["role"], "content": m["content"]}
                        for m in st.session_state.messages
                    ]
                )
                
                assistant_reply = response.content[0].text
                st.markdown(assistant_reply)
                st.session_state.messages.append({
                    "role": "assistant",
                    "content": assistant_reply
                })
                
            except Exception as e:
                st.error(f"API Error: {str(e)}")

# Clear chat button
if st.button("🗑️ Clear Chat"):
    st.session_state.messages = []
    st.rerun()
import streamlit as st
from groq import Groq
from fpdf import FPDF

# Page Configuration
st.set_page_config(
    page_title="AI Multi-Tool Studio", 
    page_icon="⚡", 
    layout="wide"
)

# Custom Styling
st.markdown("""
    <style>
    .stApp { background-color: #f8fafc; color: #1e293b; }
    .header-box {
        background: linear-gradient(135deg, #e0f2fe 0%, #bae6fd 100%);
        padding: 20px; border-radius: 12px; border: 1px solid #7dd3fc;
        margin-bottom: 20px;
    }
    .main-title { color: #0369a1; font-size: 26px; font-weight: 800; margin: 0; }
    .stButton>button {
        background: #0284c7 !important; color: #ffffff !important;
        border: none !important; border-radius: 8px !important;
        font-weight: 600 !important; width: 100%;
    }
    </style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
    <div class="header-box">
        <div class="main-title">⚡ AI Multi-Tool Studio</div>
    </div>
""", unsafe_allow_html=True)

# PDF Generator Function
def create_pdf(text_content):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=11)
    clean_text = text_content.encode('latin-1', 'replace').decode('latin-1')
    pdf.multi_cell(0, 7, clean_text)
    return bytes(pdf.output())

# AI Engine Function
def call_groq_ai(prompt_text):
    api_key = st.secrets.get("GROQ_API_KEY", "")
    if not api_key:
        raise Exception("Streamlit Secrets mein GROQ_API_KEY nahi mili!")
    
    clean_key = str(api_key).strip().strip('"').strip("'")
    client = Groq(api_key=clean_key)
    
    # Active Models Priority
    models = ["llama-3.1-8b-instant", "llama-3.3-70b-versatile"]
    
    for model in models:
        try:
            res = client.chat.completions.create(
                messages=[{"role": "user", "content": prompt_text}],
                model=model
            )
            return res.choices[0].message.content
        except Exception:
            continue
            
    raise Exception("Groq API Call Failed. API Key check karein.")

# Sidebar
st.sidebar.success("⚡ Engine Active")

# Main Interface Tabs
tab1, tab2, tab3 = st.tabs(["✍️ Content Creator", "🌐 Translator", "📚 Academic Writer"])

# TAB 1
with tab1:
    st.subheader("✍️ Content Generator")
    topic = st.text_area("Enter Topic / Instructions", placeholder="Write about emission control...")
    btn1 = st.button("🚀 Generate Output", key="b1")
    
    if btn1:
        if not topic:
            st.warning("Topic likhein pehle!")
        else:
            try:
                with st.spinner("Generating..."):
                    out = call_groq_ai(topic)
                st.success("Done!")
                st.markdown(out)
                st.download_button("📥 Download PDF", data=create_pdf(out), file_name="output.pdf", mime="application/pdf")
            except Exception as e:
                st.error(f"Error: {e}")

# TAB 2
with tab2:
    st.subheader("🌐 Translator")
    lang = st.selectbox("Target Language", ["Urdu", "English", "Arabic", "Spanish", "French"])
    trans_text = st.text_area("Source Text", placeholder="Paste text here...")
    btn2 = st.button("🌐 Translate", key="b2")
    
    if btn2:
        if not trans_text:
            st.warning("Text provide karein!")
        else:
            try:
                with st.spinner("Translating..."):
                    out = call_groq_ai(f"Translate to {lang}:\n\n{trans_text}")
                st.success("Done!")
                st.markdown(out)
                st.download_button("📥 Download PDF", data=create_pdf(out), file_name="translation.pdf", mime="application/pdf")
            except Exception as e:
                st.error(f"Error: {e}")

# TAB 3
with tab4 if 'tab4' in locals() else tab3:
    st.subheader("📚 Academic Writer")
    academic_topic = st.text_input("Academic Subject/Topic")
    btn3 = st.button("✨ Create Paper", key="b3")
    
    if btn3:
        if not academic_topic:
            st.warning("Topic likhein!")
        else:
            try:
                with st.spinner("Writing..."):
                    out = call_groq_ai(f"Write a detailed academic paper on: {academic_topic}")
                st.success("Done!")
                st.markdown(out)
                st.download_button("📥 Download PDF", data=create_pdf(out), file_name="academic.pdf", mime="application/pdf")
            except Exception as e:
                st.error(f"Error: {e}")

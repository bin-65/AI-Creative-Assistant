import streamlit as st
from google import genai
import pypdf
import docx
from docx import Document
from fpdf import FPDF
import io
import time

# Page configuration
st.set_page_config(
    page_title="Usman AI Studio", 
    page_icon="⚡", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
    <style>
    .stApp { background-color: #f8fafc; color: #1e293b; }
    .header-box {
        background: linear-gradient(135deg, #e0f2fe 0%, #bae6fd 100%);
        padding: 24px; border-radius: 12px; border: 1px solid #7dd3fc;
        margin-bottom: 20px; box-shadow: 0 4px 12px rgba(14, 165, 233, 0.12);
    }
    .main-title { color: #0369a1; font-size: 28px; font-weight: 800; margin: 0; }
    .sub-title { color: #0284c7; font-size: 14px; margin-top: 4px; font-weight: 500; }
    
    .stButton>button {
        background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%) !important;
        color: #ffffff !important; border: none !important;
        border-radius: 8px !important; font-weight: 600 !important;
        padding: 8px 18px !important; transition: all 0.3s ease !important;
    }
    .stDownloadButton>button {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
        color: #ffffff !important; border: none !important;
        border-radius: 8px !important; font-weight: 600 !important;
        padding: 8px 18px !important; width: 100% !important;
    }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px; padding: 8px 18px; background-color: #ffffff;
        color: #475569; border: 1px solid #e2e8f0;
    }
    .stTabs [aria-selected="true"] { background-color: #0284c7 !important; color: #ffffff !important; border: none !important; }
    </style>
""", unsafe_allow_html=True)

# Main Top Header Bar
st.markdown("""
    <div class="header-box">
        <div class="main-title">⚡ Usman AI Studio</div>
        <div class="sub-title">Powered by Google Gemini AI</div>
    </div>
""", unsafe_allow_html=True)

# Helper Export Functions (PDF & Word Generation)
def generate_pdf_bytes(title, text_content):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, title.encode('latin-1', 'replace').decode('latin-1'), ln=True)
    pdf.set_font("Arial", size=10)
    pdf.ln(5)
    clean_text = text_content.encode('latin-1', 'replace').decode('latin-1')
    pdf.multi_cell(0, 7, clean_text)
    return bytes(pdf.output())

def generate_docx_bytes(title, text_content):
    doc = Document()
    doc.add_heading(title, 0)
    for paragraph in text_content.split("\n\n"):
        if paragraph.strip():
            doc.add_paragraph(paragraph.strip())
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# Gemini AI Engine
def call_gemini_ai(prompt_text):
    api_key = st.secrets.get("GEMINI_API_KEY", "")
    if not api_key:
        raise Exception("GEMINI_API_KEY Missing in Streamlit Secrets!")
    
    clean_key = str(api_key).strip().strip('"').strip("'")
    client = genai.Client(api_key=clean_key)
    models_to_try = ['gemini-3.6-flash', 'gemini-2.5-flash']
    
    for model_name in models_to_try:
        for attempt in range(3):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=prompt_text,
                )
                return response.text
            except Exception as e:
                error_msg = str(e)
                if "503" in error_msg or "UNAVAILABLE" in error_msg:
                    time.sleep(1.5)
                    continue
                else:
                    break
                    
    raise Exception("Server is busy. Please try clicking generate again in a few seconds.")

# Universal Output Component for ALL features
def render_output_section(result_text, doc_title, key_prefix):
    st.markdown("---")
    st.success("🎉 Generated Successfully!")
    
    st.markdown("### 📥 Download Options")
    try:
        pdf_bytes = generate_pdf_bytes(doc_title, result_text)
        docx_bytes = generate_docx_bytes(doc_title, result_text)
        
        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                label="📥 Download as PDF (.pdf)",
                data=pdf_bytes,
                file_name=f"{doc_title}.pdf",
                mime="application/pdf",
                key=f"{key_prefix}_pdf_btn"
            )
        with c2:
            st.download_button(
                label="📄 Download as Word (.docx)",
                data=docx_bytes,
                file_name=f"{doc_title}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"{key_prefix}_docx_btn"
            )
    except Exception as e:
        st.error(f"Error preparing download files: {e}")

    st.markdown("---")
    st.markdown("### 📄 Formatted Text Output")
    st.markdown(result_text)

    st.markdown("### 📋 Copy Raw Text from Screen")
    st.text_area(
        label="Copyable Text Area",
        value=result_text,
        height=200,
        key=f"{key_prefix}_text_area"
    )

# Sidebar
st.sidebar.markdown("### ⚙️ System Status")
st.sidebar.success("✨ **Usman AI Studio Active**")
st.sidebar.markdown("---")

# Main Tabs Navigation
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "✍️ Content Creator", 
    "🌐 Translator", 
    "⚡ Smart AI Workspace",
    "📚 Academic Writer",
    "📑 Advanced Doc Hub",
    "❓ MCQs & Quiz Generator"
])

# TAB 1: CONTENT CREATOR
with tab1:
    st.subheader("✍️ Social Media Content Generator")
    c1, c2 = st.columns(2)
    with c1:
        platform = st.selectbox("Target Platform", ["LinkedIn", "Instagram", "Twitter / X", "Facebook"], key="t1_plat")
        content_type = st.selectbox("Content Style", ["Informational Post", "Promotional / Ad", "Storytelling"], key="t1_style")
    with c2:
        tone = st.selectbox("Tone & Persona", ["Professional", "Casual & Friendly", "Persuasive"], key="t1_tone")
        target_audience = st.text_input("Target Audience", placeholder="e.g., Engineers, Students", key="t1_aud")
    
    topic = st.text_area("Core Brief / Topic", key="t1_topic")
    if st.button("🚀 Generate Content", key="t1_gen_btn"):
        if not topic or not target_audience:
            st.warning("⚠️ Please enter Topic and Target Audience.")
        else:
            try:
                with st.spinner("Generating Content..."):
                    res_text = call_gemini_ai(f"Platform: {platform}\nType: {content_type}\nTone: {tone}\nAudience: {target_audience}\nTopic: {topic}")
                st.session_state["res_tab1"] = res_text
            except Exception as e:
                st.error(f"Execution Error: {e}")

    if "res_tab1" in st.session_state:
        render_output_section(st.session_state["res_tab1"], f"{platform}_Content", "tab1")

# TAB 2: TRANSLATOR
with tab2:
    st.subheader("🌐 Global Multi-Language Translator")
    target_language = st.selectbox("Target Language", ["English", "Urdu", "Arabic", "Hindi", "Spanish", "French", "German"], key="t2_lang")
    input_text = st.text_area("Source Text", height=150, key="t2_text")
    if st.button("🌐 Translate Content", key="t2_gen_btn"):
        if not input_text.strip():
            st.warning("⚠️ Please enter text to translate.")
        else:
            try:
                with st.spinner("Translating..."):
                    res_text = call_gemini_ai(f"Translate to {target_language}:\n\n{input_text}")
                st.session_state["res_tab2"] = res_text
            except Exception as e:
                st.error(f"Execution Error: {e}")

    if "res_tab2" in st.session_state:
        render_output_section(st.session_state["res_tab2"], f"Translation_{target_language}", "tab2")

# TAB 3: SMART WORKSPACE
with tab3:
    st.subheader("⚡ Smart File Workspace")
    user_prompt = st.text_area("Analysis Prompt", height=100, key="t3_prompt")
    uploaded_files = st.file_uploader("Upload Documents", type=["pdf", "docx", "txt"], accept_multiple_files=True, key="t3_files")
    if st.button("⚡ Process Query", key="t3_gen_btn"):
        try:
            with st.spinner("Processing Documents..."):
                extracted_text = ""
                if uploaded_files:
                    for file in uploaded_files:
                        if file.name.endswith(".pdf"):
                            reader = pypdf.PdfReader(file)
                            extracted_text += "\n".join([p.extract_text() or "" for p in reader.pages[:10]])
                        elif file.name.endswith(".docx"):
                            doc = docx.Document(file)
                            extracted_text += "\n".join([p.text for p in doc.paragraphs])
                res_text = call_gemini_ai(f"Extracted Text:\n{extracted_text[:8000]}\nUser Prompt: {user_prompt}")
            st.session_state["res_tab3"] = res_text
        except Exception as e:
            st.error(f"Execution Error: {e}")

    if "res_tab3" in st.session_state:
        render_output_section(st.session_state["res_tab3"], "Workspace_Output", "tab3")

# TAB 4: ACADEMIC WRITER
with tab4:
    st.subheader("📚 Academic & Assignment Writer")
    subject_topic = st.text_input("Topic Header", key="t4_topic")
    academic_level = st.selectbox("Academic Level", ["Undergraduate", "Postgraduate / PhD", "College"], key="t4_level")
    if st.button("✨ Generate Assignment", key="t4_gen_btn"):
        if not subject_topic.strip():
            st.warning("⚠️ Please enter a topic header.")
        else:
            try:
                with st.spinner("Writing Academic Content..."):
                    res_text = call_gemini_ai(f"Write paper on '{subject_topic}' for {academic_level} level.")
                st.session_state["res_tab4"] = res_text
            except Exception as e:
                st.error(f"Execution Error: {e}")

    if "res_tab4" in st.session_state:
        render_output_section(st.session_state["res_tab4"], f"{subject_topic}_Assignment", "tab4")

# TAB 5: DOCUMENT HUB
with tab5:
    st.subheader("📑 Advanced Document Hub")
    doc_topic = st.text_input("Document Topic", key="t5_topic")
    if st.button("✨ Build Document", key="t5_gen_btn"):
        if not doc_topic.strip():
            st.warning("⚠️ Please enter a document topic.")
        else:
            try:
                with st.spinner("Building Document..."):
                    res_text = call_gemini_ai(f"Create a detailed document on '{doc_topic}'.")
                st.session_state["res_tab5"] = res_text
            except Exception as e:
                st.error(f"Execution Error: {e}")

    if "res_tab5" in st.session_state:
        render_output_section(st.session_state["res_tab5"], f"{doc_topic}_Document", "tab5")

# TAB 6: MCQS & QUIZ GENERATOR
with tab6:
    st.subheader("❓ MCQs & Quiz Generator")
    quiz_topic = st.text_input("Quiz Topic / Subject Header", placeholder="e.g. Thermodynamics, Machine Learning", key="t6_topic")
    c1, c2 = st.columns(2)
    with c1:
        num_mcqs = st.slider("Number of Questions", min_value=5, max_value=30, value=10, key="t6_num")
    with c2:
        difficulty = st.selectbox("Difficulty Level", ["Easy", "Medium", "Hard", "Mixed"], key="t6_diff")
    
    quiz_file = st.file_uploader("Optional: Upload Source Document (PDF/DOCX)", type=["pdf", "docx", "txt"], key="t6_file")

    if st.button("🎯 Generate MCQs Test", key="t6_gen_btn"):
        if not quiz_topic.strip() and not quiz_file:
            st.warning("⚠️ Please provide either a topic name or upload a document.")
        else:
            try:
                with st.spinner("Generating MCQs Quiz..."):
                    extracted_text = ""
                    if quiz_file:
                        if quiz_file.name.endswith(".pdf"):
                            reader = pypdf.PdfReader(quiz_file)
                            extracted_text = "\n".join([p.extract_text() or "" for p in reader.pages[:10]])
                        elif quiz_file.name.endswith(".docx"):
                            doc = docx.Document(quiz_file)
                            extracted_text = "\n".join([p.text for p in doc.paragraphs])
                    
                    mcq_prompt = f"""
                    Create a structured Multiple Choice Questions (MCQs) test.
                    Topic/Subject: {quiz_topic}
                    Number of Questions: {num_mcqs}
                    Difficulty Level: {difficulty}
                    Source Text Context (if provided): {extracted_text[:6000]}

                    Format:
                    1. Question Statement
                       A) Option 1
                       B) Option 2
                       C) Option 3
                       D) Option 4
                    
                    At the end of all questions, provide an 'Answer Key & Detailed Explanations' section.
                    """
                    
                    res_text = call_gemini_ai(mcq_prompt)
                    st.session_state["res_tab6"] = res_text
            except Exception as e:
                st.error(f"Execution Error: {e}")

    if "res_tab6" in st.session_state:
        render_output_section(st.session_state["res_tab6"], f"{quiz_topic or 'Quiz'}_MCQs", "tab6")

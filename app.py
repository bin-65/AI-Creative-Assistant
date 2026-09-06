import streamlit as st
from google import genai
from google.genai.errors import APIError
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
        padding: 8px 18px !important;
    }
    .stDownloadButton>button {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
        color: #ffffff !important; border: none !important;
        border-radius: 8px !important; font-weight: 600 !important;
        padding: 10px 18px !important; width: 100% !important;
    }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px; padding: 8px 18px; background-color: #ffffff;
        color: #475569; border: 1px solid #e2e8f0;
    }
    .stTabs [aria-selected="true"] { background-color: #0284c7 !important; color: #ffffff !important; border: none !important; }
    </style>
""", unsafe_allow_html=True)

# Main Header
st.markdown("""
    <div class="header-box">
        <div class="main-title">⚡ Usman AI Studio</div>
        <div class="sub-title">Powered by Google Gemini AI</div>
    </div>
""", unsafe_allow_html=True)

# Safe Export Functions
def generate_pdf_bytes(title, text_content):
    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", 'B', 14)
        clean_title = title.encode('latin-1', 'ignore').decode('latin-1')
        pdf.cell(0, 10, clean_title if clean_title else "Document", ln=True)
        pdf.set_font("Arial", size=10)
        pdf.ln(5)
        clean_text = text_content.encode('latin-1', 'ignore').decode('latin-1')
        pdf.multi_cell(0, 7, clean_text)
        return bytes(pdf.output())
    except Exception:
        return text_content.encode("utf-8")

def generate_docx_bytes(title, text_content):
    try:
        doc = Document()
        doc.add_heading(title, 0)
        for paragraph in text_content.split("\n\n"):
            if paragraph.strip():
                doc.add_paragraph(paragraph.strip())
        bio = io.BytesIO()
        doc.save(bio)
        return bio.getvalue()
    except Exception:
        return text_content.encode("utf-8")

# Multi-Key Rotation Engine with Official Stable Model Names
def call_gemini_ai(prompt_text):
    raw_keys = st.secrets.get("GEMINI_API_KEY", "")
    if not raw_keys:
        raise Exception("Secrets mein GEMINI_API_KEY mojood nahi hai. Pehle st.secrets mein API Key add karein.")
    
    # Extract keys (split by comma if multiple keys are provided)
    keys_list = [k.strip().strip('"').strip("'") for k in str(raw_keys).split(",") if k.strip()]
    
    # Official stable models supported by google-genai SDK
    models_to_try = [
        "gemini-2.5-flash",
        "gemini-2.5-pro",
        "gemini-2.0-flash"
    ]
    
    last_error_msg = ""
    
    # Loop through each API Key
    for api_key in keys_list:
        client = genai.Client(api_key=api_key)
        
        # Loop through each Model for the current key
        for model_name in models_to_try:
            for attempt in range(1, 3):
                try:
                    response = client.models.generate_content(
                        model=model_name,
                        contents=prompt_text
                    )
                    if response and response.text:
                        return response.text
                except APIError as e:
                    last_error_msg = f"API Error ({model_name}): {e.message if hasattr(e, 'message') else str(e)}"
                    time.sleep(1)
                except Exception as e:
                    last_error_msg = f"Error ({model_name}): {str(e)}"
                    time.sleep(1)
                    
    raise Exception(f"Tamam API Keys aur Models busy hain. Thodi der baad try karein. Details: {last_error_msg}")

# Output Display Renderer
def render_output_section(result_text, doc_title, key_prefix, selected_format):
    st.markdown("---")
    st.success("🎉 Content Successfully Generate Ho Gaya!")
    
    if selected_format == "PDF Document (.pdf)":
        st.markdown("### 📥 Download PDF Document")
        pdf_bytes = generate_pdf_bytes(doc_title, result_text)
        st.download_button(
            label="📥 Download PDF (.pdf)",
            data=pdf_bytes,
            file_name=f"{doc_title}.pdf",
            mime="application/pdf",
            key=f"{key_prefix}_pdf_btn"
        )
        st.markdown("#### Preview:")
        st.markdown(result_text)

    elif selected_format == "MS Word (.docx)":
        st.markdown("### 📄 Download MS Word Document")
        docx_bytes = generate_docx_bytes(doc_title, result_text)
        st.download_button(
            label="📄 Download MS Word (.docx)",
            data=docx_bytes,
            file_name=f"{doc_title}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"{key_prefix}_docx_btn"
        )
        st.markdown("#### Preview:")
        st.markdown(result_text)

    elif selected_format == "On-Screen Text":
        st.markdown("### 📋 On-Screen Text & Copy Box")
        st.text_area(
            label="Copyable Text Box",
            value=result_text,
            height=220,
            key=f"{key_prefix}_copy_box"
        )
        st.markdown("#### Preview:")
        st.markdown(result_text)

# Sidebar
st.sidebar.markdown("### ⚙️ System Status")
st.sidebar.success("✨ **Usman AI Studio Active (Multi-Key Rotation Enabled)**")
st.sidebar.markdown("---")

# Main Navigation
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "✍️ Content Creator", 
    "🌐 Translator", 
    "⚡ Smart AI Workspace",
    "📚 Academic Writer",
    "📑 Advanced Doc Hub",
    "❓ MCQs & Quiz Generator"
])

EXPORT_OPTIONS = ["PDF Document (.pdf)", "MS Word (.docx)", "On-Screen Text"]

# TAB 1: CONTENT CREATOR
with tab1:
    st.subheader("✍️ Social Media Content Generator")
    c1, c2 = st.columns(2)
    with c1:
        platform = st.selectbox("Target Platform", ["LinkedIn", "Instagram", "Twitter / X", "Facebook"], key="t1_plat")
        content_type = st.selectbox("Content Style", ["Informational Post", "Promotional / Ad", "Storytelling"], key="t1_style")
    with c2:
        tone = st.selectbox("Tone & Persona", ["Professional", "Casual & Friendly", "Persuasive"], key="t1_tone")
        target_audience = st.text_input("Target Audience", value="student", key="t1_aud")
    
    export_fmt_1 = st.radio("Export Format Option", EXPORT_OPTIONS, horizontal=True, key="t1_fmt")
    topic_1 = st.text_area("Core Brief / Topic", placeholder="e.g. emission control", key="t1_topic")
    
    if st.button("🚀 Generate Content", key="t1_gen_btn"):
        if not topic_1.strip():
            st.warning("⚠️ Pehle topic brief likhein.")
        else:
            try:
                with st.spinner("AI Server respond kar raha hai, please wait..."):
                    res_text = call_gemini_ai(f"Platform: {platform}\nType: {content_type}\nTone: {tone}\nAudience: {target_audience}\nTopic: {topic_1}")
                st.session_state["out_tab1"] = res_text
                st.session_state["fmt_tab1"] = export_fmt_1
            except Exception as e:
                st.error(f"Execution Error: {e}")

    if "out_tab1" in st.session_state:
        render_output_section(st.session_state["out_tab1"], f"{platform}_Content", "tab1", st.session_state["fmt_tab1"])

# TAB 2: TRANSLATOR
with tab2:
    st.subheader("🌐 Global Multi-Language Translator")
    target_language = st.selectbox("Target Language", ["English", "Urdu", "Arabic", "Hindi", "Spanish", "French", "German"], key="t2_lang")
    export_fmt_2 = st.radio("Export Format Option", EXPORT_OPTIONS, horizontal=True, key="t2_fmt")
    input_text_2 = st.text_area("Source Text", height=150, key="t2_text")
    
    if st.button("🌐 Translate Content", key="t2_gen_btn"):
        if not input_text_2.strip():
            st.warning("⚠️ Translation ke liye text darj karein.")
        else:
            try:
                with st.spinner("Translate ho raha hai..."):
                    res_text = call_gemini_ai(f"Translate to {target_language}:\n\n{input_text_2}")
                st.session_state["out_tab2"] = res_text
                st.session_state["fmt_tab2"] = export_fmt_2
            except Exception as e:
                st.error(f"Execution Error: {e}")

    if "out_tab2" in st.session_state:
        render_output_section(st.session_state["out_tab2"], f"Translation_{target_language}", "tab2", st.session_state["fmt_tab2"])

# TAB 3: SMART WORKSPACE
with tab3:
    st.subheader("⚡ Smart File Workspace")
    user_prompt_3 = st.text_area("Analysis Prompt", height=100, key="t3_prompt")
    uploaded_files_3 = st.file_uploader("Upload Documents", type=["pdf", "docx", "txt"], accept_multiple_files=True, key="t3_files")
    export_fmt_3 = st.radio("Export Format Option", EXPORT_OPTIONS, horizontal=True, key="t3_fmt")
    
    if st.button("⚡ Process Query", key="t3_gen_btn"):
        try:
            with st.spinner("Files process ho rahi hain..."):
                extracted_text = ""
                if uploaded_files_3:
                    for file in uploaded_files_3:
                        if file.name.endswith(".pdf"):
                            reader = pypdf.PdfReader(file)
                            extracted_text += "\n".join([p.extract_text() or "" for p in reader.pages[:10]])
                        elif file.name.endswith(".docx"):
                            doc = docx.Document(file)
                            extracted_text += "\n".join([p.text for p in doc.paragraphs])
                res_text = call_gemini_ai(f"Extracted Text:\n{extracted_text[:8000]}\nUser Prompt: {user_prompt_3}")
            st.session_state["out_tab3"] = res_text
            st.session_state["fmt_tab3"] = export_fmt_3
        except Exception as e:
            st.error(f"Execution Error: {e}")

    if "out_tab3" in st.session_state:
        render_output_section(st.session_state["out_tab3"], "Workspace_Output", "tab3", st.session_state["fmt_tab3"])

# TAB 4: ACADEMIC WRITER
with tab4:
    st.subheader("📚 Academic & Assignment Writer")
    subject_topic_4 = st.text_input("Topic Header", key="t4_topic")
    academic_level_4 = st.selectbox("Academic Level", ["Undergraduate", "Postgraduate / PhD", "College"], key="t4_level")
    export_fmt_4 = st.radio("Export Format Option", EXPORT_OPTIONS, horizontal=True, key="t4_fmt")
    
    if st.button("✨ Generate Assignment", key="t4_gen_btn"):
        if not subject_topic_4.strip():
            st.warning("⚠️ Pehle topic header enter karein.")
        else:
            try:
                with st.spinner("Academic content tayyar ho raha hai..."):
                    res_text = call_gemini_ai(f"Write paper on '{subject_topic_4}' for {academic_level_4} level.")
                st.session_state["out_tab4"] = res_text
                st.session_state["fmt_tab4"] = export_fmt_4
            except Exception as e:
                st.error(f"Execution Error: {e}")

    if "out_tab4" in st.session_state:
        render_output_section(st.session_state["out_tab4"], f"{subject_topic_4}_Assignment", "tab4", st.session_state["fmt_tab4"])

# TAB 5: DOCUMENT HUB
with tab5:
    st.subheader("📑 Advanced Document Hub")
    doc_topic_5 = st.text_input("Document Topic", key="t5_topic")
    export_fmt_5 = st.radio("Export Format Option", EXPORT_OPTIONS, horizontal=True, key="t5_fmt")
    
    if st.button("✨ Build Document", key="t5_gen_btn"):
        if not doc_topic_5.strip():
            st.warning("⚠️ Pehle document topic enter karein.")
        else:
            try:
                with st.spinner("Document build ho raha hai..."):
                    res_text = call_gemini_ai(f"Create a detailed document on '{doc_topic_5}'.")
                st.session_state["out_tab5"] = res_text
                st.session_state["fmt_tab5"] = export_fmt_5
            except Exception as e:
                st.error(f"Execution Error: {e}")

    if "out_tab5" in st.session_state:
        render_output_section(st.session_state["out_tab5"], f"{doc_topic_5}_Document", "tab5", st.session_state["fmt_tab5"])

# TAB 6: MCQS & QUIZ GENERATOR
with tab6:
    st.subheader("❓ MCQs & Quiz Generator")
    quiz_topic_6 = st.text_input("Quiz Topic / Subject Header", placeholder="e.g. Thermodynamics, Machine Learning", key="t6_topic")
    c1, c2 = st.columns(2)
    with c1:
        num_mcqs_6 = st.slider("Number of Questions", min_value=5, max_value=30, value=10, key="t6_num")
    with c2:
        difficulty_6 = st.selectbox("Difficulty Level", ["Easy", "Medium", "Hard", "Mixed"], key="t6_diff")
    
    quiz_file_6 = st.file_uploader("Optional: Upload Source Document (PDF/DOCX)", type=["pdf", "docx", "txt"], key="t6_file")
    export_fmt_6 = st.radio("Export Format Option", EXPORT_OPTIONS, horizontal=True, key="t6_fmt")

    if st.button("🎯 Generate MCQs Test", key="t6_gen_btn"):
        if not quiz_topic_6.strip() and not quiz_file_6:
            st.warning("⚠️ Yahan topic enter karein ya document upload karein.")
        else:
            try:
                with st.spinner("Quiz generate ho rahi hai..."):
                    extracted_text = ""
                    if quiz_file_6:
                        if quiz_file_6.name.endswith(".pdf"):
                            reader = pypdf.PdfReader(quiz_file_6)
                            extracted_text = "\n".join([p.extract_text() or "" for p in reader.pages[:10]])
                        elif quiz_file_6.name.endswith(".docx"):
                            doc = docx.Document(quiz_file_6)
                            extracted_text = "\n".join([p.text for p in doc.paragraphs])
                    
                    mcq_prompt = f"""
                    Create a structured Multiple Choice Questions (MCQs) test.
                    Topic/Subject: {quiz_topic_6}
                    Number of Questions: {num_mcqs_6}
                    Difficulty Level: {difficulty_6}
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
                    st.session_state["out_tab6"] = res_text
                    st.session_state["fmt_tab6"] = export_fmt_6
            except Exception as e:
                st.error(f"Execution Error: {e}")

    if "out_tab6" in st.session_state:
        render_output_section(st.session_state["out_tab6"], f"{quiz_topic_6 or 'Quiz'}_MCQs", "tab6", st.session_state["fmt_tab6"])

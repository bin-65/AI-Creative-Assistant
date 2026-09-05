import streamlit as st
from google import genai
import pypdf
import docx
import pptx
from docx import Document
from pptx import Presentation
from fpdf import FPDF
import io

# Page configuration
st.set_page_config(
    page_title="AI Multi-Tool Studio", 
    page_icon="⚡", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling (Original Blue Gradient Theme)
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
        <div class="main-title">⚡ AI Multi-Tool Studio</div>
        <div class="sub-title">Powered by Google Gemini AI</div>
    </div>
""", unsafe_allow_html=True)

# Helper Export Functions
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
        doc.add_paragraph(paragraph)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# Gemini AI Engine Function (UPDATED MODEL NAME TO gemini-3.6-flash)
def call_gemini_ai(prompt_text):
    api_key = st.secrets.get("GEMINI_API_KEY", "")
    if not api_key:
        raise Exception("GEMINI_API_KEY Missing in Streamlit Secrets!")
    
    clean_key = str(api_key).strip().strip('"').strip("'")
    client = genai.Client(api_key=clean_key)
    
    try:
        response = client.models.generate_content(
            model='gemini-3.6-flash',
            contents=prompt_text,
        )
        return response.text
    except Exception as e:
        raise Exception(f"Gemini API Execution Failed: {str(e)}")

# System Status Sidebar
st.sidebar.markdown("### ⚙️ System Status")
st.sidebar.success("✨ **Google Gemini Active**")
st.sidebar.markdown("---")

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "✍️ Content Creator", 
    "🌐 Translator", 
    "⚡ Smart AI Workspace",
    "📚 Academic Writer",
    "📑 Advanced Doc Hub"
])

def render_output_section(result_text, doc_title, target_format, key_prefix):
    st.markdown("---")
    st.success("🎉 Generated Successfully!")
    
    try:
        pdf_data = generate_pdf_bytes(doc_title, result_text)
        docx_data = generate_docx_bytes(doc_title, result_text)
        
        col_pdf, col_doc = st.columns([1, 1])
        
        with col_pdf:
            st.download_button(
                label="📥 Download PDF (.pdf)", 
                data=pdf_data, 
                file_name=f"{doc_title}.pdf", 
                mime="application/pdf", 
                use_container_width=True,
                key=f"{key_prefix}_pdf"
            )
        with col_doc:
            st.download_button(
                label="📄 Download Word (.docx)", 
                data=docx_data, 
                file_name=f"{doc_title}.docx", 
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                use_container_width=True,
                key=f"{key_prefix}_docx"
            )
    except Exception as export_err:
        st.error(f"Error generating download files: {export_err}")
        
    st.markdown("### 📄 Result Output:")
    st.markdown(result_text)

# TAB 1: CONTENT ASSISTANT
with tab1:
    st.subheader("✍️ Social Media Content Generator")
    col1, col2 = st.columns(2)
    with col1:
        platform = st.selectbox("Target Platform", ["LinkedIn", "Instagram", "Twitter / X", "Facebook"])
        content_type = st.selectbox("Content Style", ["Informational Post", "Promotional / Ad", "Storytelling"])
    with col2:
        tone = st.selectbox("Tone & Persona", ["Professional", "Casual & Friendly", "Persuasive", "Inspirational"])
        target_audience = st.text_input("Target Audience", placeholder="e.g., Tech Founders, Students")
    
    output_format_1 = st.radio("Export Format Option", ["PDF Document (.pdf)", "MS Word (.docx)", "On-Screen Text"], horizontal=True, key="of1")
    topic = st.text_area("Core Brief / Topic", placeholder="Describe key talking points...")
    submit_btn = st.button("🚀 Generate Content", key="tab1_btn")

    if submit_btn:
        if not topic or not target_audience:
            st.warning("⚠️ Please complete required fields.")
        else:
            try:
                prompt = f"Platform: {platform}\nType: {content_type}\nTone: {tone}\nAudience: {target_audience}\nTopic: {topic}\nWrite a well-formatted post."
                with st.spinner("Processing with Gemini..."):
                    res_text = call_gemini_ai(prompt)
                render_output_section(res_text, f"{platform}_Content", output_format_1, "tab1")
            except Exception as e:
                st.error(f"Execution Error: {e}")

# TAB 2: TRANSLATOR
with tab2:
    st.subheader("🌐 Global Multi-Language Translator")
    languages_50 = ["English", "Urdu", "Arabic", "Hindi", "Pashto", "Punjabi", "Sindhi", "Spanish", "French", "German", "Chinese"]
    target_language = st.selectbox("Select Target Language", languages_50)
    output_format_2 = st.radio("Export Format Option", ["PDF Document (.pdf)", "MS Word (.docx)", "On-Screen Text"], horizontal=True, key="of2")
    input_text = st.text_area("Source Text", placeholder="Paste source text here...", height=150)
    translate_btn = st.button("🌐 Translate Content", key="tab2_btn")

    if translate_btn:
        if not input_text.strip():
            st.warning("⚠️ Please provide source text.")
        else:
            try:
                translation_prompt = f"Automatically detect source language and translate accurately to {target_language}:\n\n{input_text}"
                with st.spinner("Translating with Gemini..."):
                    res_text = call_gemini_ai(translation_prompt)
                render_output_section(res_text, f"Translation_{target_language}", output_format_2, "tab2")
            except Exception as e:
                st.error(f"Execution Error: {e}")

# TAB 3: SMART WORKSPACE
with tab3:
    st.subheader("⚡ Smart File Workspace")
    output_format_3 = st.radio("Select Preferred Export File", ["PDF Document (.pdf)", "MS Word (.docx)", "On-Screen Text"], horizontal=True, key="of3")
    user_prompt = st.text_area("Analysis Prompt", placeholder="e.g., 10 mcqs with answers...", height=100)
    uploaded_files = st.file_uploader(
        "Upload Documents", 
        type=["pdf", "docx", "pptx", "txt"], 
        accept_multiple_files=True
    )
    process_btn = st.button("⚡ Process Query", key="tab3_btn")

    if process_btn:
        if not user_prompt.strip() and not uploaded_files:
            st.warning("⚠️ Enter instructions or upload files.")
        else:
            try:
                with st.spinner("Processing with Gemini..."):
                    extracted_text_from_docs = ""

                    if uploaded_files:
                        for file in uploaded_files:
                            filename = file.name.lower()
                            if filename.endswith(".pdf"):
                                pdf_reader = pypdf.PdfReader(file)
                                pdf_text = "\n".join([page.extract_text() or "" for page in pdf_reader.pages[:15]])
                                extracted_text_from_docs += f"\n--- [PDF: {file.name}] ---\n{pdf_text}\n"
                            elif filename.endswith(".docx"):
                                doc_file = docx.Document(file)
                                docx_text = "\n".join([p.text for p in doc_file.paragraphs])
                                extracted_text_from_docs += f"\n--- [Word: {file.name}] ---\n{docx_text}\n"
                            elif filename.endswith(".pptx"):
                                prs_file = pptx.Presentation(file)
                                pptx_text = ""
                                for slide in prs_file.slides:
                                    for shape in slide.shapes:
                                        if hasattr(shape, "text"): pptx_text += shape.text + "\n"
                                extracted_text_from_docs += f"\n--- [PPT: {file.name}] ---\n{pptx_text}\n"
                            elif filename.endswith(".txt"):
                                extracted_text_from_docs += f"\n--- [TXT: {file.name}] ---\n{file.read().decode('utf-8', errors='ignore')}\n"

                    final_instruction = ""
                    if extracted_text_from_docs:
                        final_instruction += f"=== EXTRACTED DOCUMENTS CONTENT ===\n{extracted_text_from_docs[:10000]}\n"
                    if user_prompt.strip():
                        final_instruction += f"=== USER INSTRUCTION ===\n{user_prompt}"

                    res_text = call_gemini_ai(final_instruction)

                render_output_section(res_text, "Workspace_Output", output_format_3, "tab3")
            except Exception as e:
                st.error(f"Execution Error: {e}")

# TAB 4: ACADEMIC WRITER
with tab4:
    st.subheader("📚 Academic & Assignment Writer")
    subject_topic = st.text_input("Topic / Subject Header", placeholder="e.g., Deep Learning & Neural Networks")
    academic_level = st.selectbox("Academic Target Level", ["School Level", "High School / College", "Undergraduate", "Postgraduate / PhD"])
    output_format_4 = st.radio("Export Assignment Format", ["PDF Document (.pdf)", "MS Word (.docx)", "On-Screen Text"], horizontal=True, key="of4")
    assign_submit = st.button("✨ Generate Assignment", key="tab4_btn")

    if assign_submit:
        if not subject_topic.strip():
            st.warning("⚠️ Topic is required.")
        else:
            try:
                prompt = f"Write a structured academic paper on '{subject_topic}' for {academic_level} level."
                with st.spinner("Generating academic assignment with Gemini..."):
                    res_text = call_gemini_ai(prompt)
                render_output_section(res_text, f"{subject_topic}_Assignment", output_format_4, "tab4")
            except Exception as e:
                st.error(f"Execution Error: {e}")

# TAB 5: DOCUMENT HUB
with tab5:
    st.subheader("📑 Advanced Document & MCQ Hub")
    doc_sub_tab1, doc_sub_tab2 = st.tabs(["📝 Create Document / Manual", "📤 Upload & Extract MCQs"])

    with doc_sub_tab1:
        doc_topic = st.text_input("Document Topic", placeholder="e.g., Fundamentals of Cybersecurity")
        export_format = st.selectbox("Format Structure", ["MS Word (.docx)", "PDF Document (.pdf)", "PowerPoint Presentation (.pptx)"])
        doc_length = st.selectbox("Scope", ["Detailed Notes (~500 words)", "Full Chapter (~1500 words)", "Full Manual (~3000+ words)"])

        create_doc_btn = st.button("✨ Build Document", key="tab5_sub1_btn")

        if create_doc_btn:
            if not doc_topic.strip():
                st.warning("⚠️ Topic required.")
            else:
                try:
                    with st.spinner("Building document with Gemini..."):
                        gen_prompt = f"Create a structured document on '{doc_topic}' with length {doc_length}."
                        generated_text = call_gemini_ai(gen_prompt)
                    render_output_section(generated_text, f"{doc_topic}_Document", export_format, "tab5_1")
                except Exception as e:
                    st.error(f"Execution Error: {e}")

    with doc_sub_tab2:
        uploaded_docs = st.file_uploader("Upload Course Material (PDF, Word, PPT, TXT)", type=["pdf", "docx", "pptx", "txt"], accept_multiple_files=True)
        mcq_count = st.selectbox("Question Volume", ["10 MCQs", "20 MCQs", "30 MCQs", "50 MCQs", "100 MCQs"])
        task_type = st.selectbox("Extraction Goal", ["Generate MCQs with Answer Key", "Summarize Key Concepts", "Extract Formulas"])
        output_format_5 = st.radio("Export File Option", ["PDF Document (.pdf)", "MS Word (.docx)", "On-Screen Text"], horizontal=True, key="of5")

        process_upload_btn = st.button("🚀 Process & Extract", key="tab5_sub2_btn")

        if process_upload_btn:
            if not uploaded_docs:
                st.warning("⚠️ Upload files first.")
            else:
                try:
                    with st.spinner("Extracting content with Gemini..."):
                        extracted_full_text = ""
                        for file in uploaded_docs:
                            extracted_full_text += f"\n--- FILE: {file.name} ---\n"
                            if file.name.endswith(".pdf"):
                                pdf_reader = pypdf.PdfReader(file)
                                for page in pdf_reader.pages[:15]: extracted_full_text += (page.extract_text() or "") + "\n"
                            elif file.name.endswith(".docx"):
                                doc = docx.Document(file)
                                for p in doc.paragraphs: extracted_full_text += p.text + "\n"
                            elif file.name.endswith(".pptx"):
                                prs = pptx.Presentation(file)
                                for slide in prs.slides:
                                    for shape in slide.shapes:
                                        if hasattr(shape, "text"): extracted_full_text += shape.text + "\n"
                            elif file.name.endswith(".txt"):
                                extracted_full_text += file.read().decode("utf-8", errors="ignore")

                        mcq_prompt = f"Task: {task_type}\nQuantity: {mcq_count}\nSource Text:\n{extracted_full_text[:10000]}"
                        output_text = call_gemini_ai(mcq_prompt)

                    render_output_section(output_text, "Extracted_MCQs", output_format_5, "tab5_2")
                except Exception as e:
                    st.error(f"Execution Error: {e}")

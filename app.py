import streamlit as st
from google import genai
import pypdf
import docx
import pptx
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
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
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px; padding: 8px 18px; background-color: #ffffff;
        color: #475569; border: 1px solid #e2e8f0;
    }
    .stTabs [aria-selected="true"] { background-color: #0284c7 !important; color: #ffffff !important; border: none !important; }
    </style>
""", unsafe_allow_html=True)

# Header Bar
st.markdown("""
    <div class="header-box">
        <div class="main-title">⚡ Usman AI Studio</div>
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
        if paragraph.strip():
            doc.add_paragraph(paragraph.strip())
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# Advanced PPT Generator supporting 10 structural PowerPoint layouts
def generate_pptx_bytes(title, text_content, template_style="Corporate Banner Blue"):
    prs = Presentation()
    blank_layout = prs.slide_layouts[6] # Blank slide for full custom shape control

    # Palette & Layout Configurations
    configs = {
        "Corporate Banner Blue": {"primary": RGBColor(3, 105, 161), "accent": RGBColor(186, 230, 253), "bg": RGBColor(255, 255, 255), "text": RGBColor(30, 41, 59)},
        "Modern Side Border": {"primary": RGBColor(124, 58, 237), "accent": RGBColor(221, 214, 254), "bg": RGBColor(255, 255, 255), "text": RGBColor(15, 23, 42)},
        "Dark Tech Neon": {"primary": RGBColor(6, 182, 212), "accent": RGBColor(30, 41, 59), "bg": RGBColor(15, 23, 42), "text": RGBColor(241, 245, 249)},
        "Executive Burgundy Header": {"primary": RGBColor(159, 18, 57), "accent": RGBColor(254, 205, 211), "bg": RGBColor(255, 255, 255), "text": RGBColor(30, 41, 59)},
        "Engineering Steel Grid": {"primary": RGBColor(217, 119, 6), "accent": RGBColor(254, 243, 199), "bg": RGBColor(248, 250, 252), "text": RGBColor(30, 41, 59)},
        "Academic Emerald Clean": {"primary": RGBColor(5, 150, 105), "accent": RGBColor(167, 243, 208), "bg": RGBColor(255, 255, 255), "text": RGBColor(30, 41, 59)},
        "Startup Pitch Card": {"primary": RGBColor(79, 70, 229), "accent": RGBColor(224, 231, 255), "bg": RGBColor(249, 250, 251), "text": RGBColor(17, 24, 39)},
        "Medical Healthcare Soft": {"primary": RGBColor(13, 148, 136), "accent": RGBColor(204, 251, 241), "bg": RGBColor(255, 255, 255), "text": RGBColor(30, 41, 59)},
        "Creative Vibrant Pink": {"primary": RGBColor(219, 39, 119), "accent": RGBColor(252, 231, 243), "bg": RGBColor(255, 255, 255), "text": RGBColor(30, 41, 59)},
        "Minimal Gray Structure": {"primary": RGBColor(71, 85, 105), "accent": RGBColor(226, 232, 240), "bg": RGBColor(255, 255, 255), "text": RGBColor(30, 41, 59)}
    }
    
    cfg = configs.get(template_style, configs["Corporate Banner Blue"])

    # Slide 1: Title Slide
    slide1 = prs.slides.add_slide(blank_layout)
    
    # Background Shape
    bg_shape = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(10), Inches(7.5))
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = cfg["bg"]
    bg_shape.line.fill.background()

    # Decorative Header Card
    hdr_box = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(2), Inches(8), Inches(3))
    hdr_box.fill.solid()
    hdr_box.fill.fore_color.rgb = cfg["primary"]
    hdr_box.line.fill.background()

    tf1 = hdr_box.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = title
    p1.font.bold = True
    p1.font.size = Pt(32)
    p1.font.color.rgb = RGBColor(255, 255, 255)
    p1.alignment = PP_ALIGN.CENTER

    p2 = tf1.add_paragraph()
    p2.text = f"\nTemplate Style: {template_style}\nGenerated by Usman AI Studio"
    p2.font.size = Pt(14)
    p2.font.color.rgb = cfg["accent"]
    p2.alignment = PP_ALIGN.CENTER

    # Content Slides Creation
    paragraphs = [p.strip() for p in text_content.split("\n\n") if p.strip()]
    chunk_size = 3
    
    for i in range(0, len(paragraphs), chunk_size):
        chunk = paragraphs[i:i + chunk_size]
        slide = prs.slides.add_slide(blank_layout)

        # Slide Background
        s_bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(10), Inches(7.5))
        s_bg.fill.solid()
        s_bg.fill.fore_color.rgb = cfg["bg"]
        s_bg.line.fill.background()

        # Top Banner / Side Structural Element
        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(10), Inches(1.1))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = cfg["primary"]
        top_bar.line.fill.background()

        # Title Text in Banner
        tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.8))
        p = tb.text_frame.paragraphs[0]
        p.text = f"{title} - Slide {(i // chunk_size) + 1}"
        p.font.bold = True
        p.font.size = Pt(22)
        p.font.color.rgb = RGBColor(255, 255, 255)

        # Decorative Side Accent Bar
        side_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.5), Inches(0.15), Inches(5.2))
        side_bar.fill.solid()
        side_bar.fill.fore_color.rgb = cfg["primary"]
        side_bar.line.fill.background()

        # Content Box
        cnt_box = slide.shapes.add_textbox(Inches(0.9), Inches(1.5), Inches(8.5), Inches(5.2))
        tf = cnt_box.text_frame
        tf.word_wrap = True

        for idx, para in enumerate(chunk):
            p_cnt = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            p_cnt.text = f"•  {para}"
            p_cnt.font.size = Pt(15)
            p_cnt.font.color.rgb = cfg["text"]
            p_cnt.space_after = Pt(14)

    bio = io.BytesIO()
    prs.save(bio)
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

# Sidebar
st.sidebar.markdown("### ⚙️ System Status")
st.sidebar.success("✨ **Usman AI Studio Active**")
st.sidebar.markdown("---")

# Main Tabs Navigation
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Presentation Studio",
    "✍️ Content Creator", 
    "🌐 Translator", 
    "⚡ Smart AI Workspace",
    "📚 Academic Writer",
    "📑 Advanced Doc Hub"
])

def render_output_section(result_text, doc_title, key_prefix, show_ppt=True, template_style="Corporate Banner Blue"):
    st.markdown("---")
    st.success("🎉 Generated Successfully!")
    
    try:
        pdf_data = generate_pdf_bytes(doc_title, result_text)
        docx_data = generate_docx_bytes(doc_title, result_text)
        pptx_data = generate_pptx_bytes(doc_title, result_text, template_style)
        
        if show_ppt:
            c1, c2, c3 = st.columns([1, 1, 1])
            with c1: st.download_button("📥 Download PDF (.pdf)", pdf_data, f"{doc_title}.pdf", "application/pdf", use_container_width=True, key=f"{key_prefix}_pdf")
            with c2: st.download_button("📄 Download Word (.docx)", docx_data, f"{doc_title}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, key=f"{key_prefix}_docx")
            with c3: st.download_button("📊 Download Presentation (.pptx)", pptx_data, f"{doc_title}.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation", use_container_width=True, key=f"{key_prefix}_pptx")
        else:
            c1, c2 = st.columns([1, 1])
            with c1: st.download_button("📥 Download PDF (.pdf)", pdf_data, f"{doc_title}.pdf", "application/pdf", use_container_width=True, key=f"{key_prefix}_pdf")
            with c2: st.download_button("📄 Download Word (.docx)", docx_data, f"{doc_title}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, key=f"{key_prefix}_docx")
    except Exception as export_err:
        st.error(f"Error generating download files: {export_err}")
        
    st.markdown("### 📄 Result Output:")
    st.markdown(result_text)

# TAB 1: PRESENTATION STUDIO (10 PPT SLIDE TEMPLATES)
with tab1:
    st.markdown("## Create Slides with AI & Edit with Full Control")
    ppt_mode = st.radio("Creation Mode", ["✨ From Topic", "📄 From Document", "📑 From Outline"], horizontal=True)
    st.markdown("---")
    
    if "From Topic" in ppt_mode:
        topic_input = st.text_area("Slide Topic / Title", value="Four stroke engine", height=100)
        
        c1, c2, c3 = st.columns([2, 2, 2])
        with c1: use_search = st.checkbox("🔍 Enable Web Search", value=True)
        with c2: deep_think = st.checkbox("🧠 DeepThink Reasoning", value=True)
        with c3: slide_count = st.slider("Target Slide Count", 5, 20, 8)

        st.markdown("##### Quick Sample Topics:")
        q1, q2, q3, q4 = st.columns(4)
        if q1.button("Health Management"): topic_input = "Health Management Program"
        if q2.button("Summary Report"): topic_input = "Summary Report"
        if q3.button("Mental Health"): topic_input = "Mental Health Report"
        if q4.button("Data Analysis"): topic_input = "Data Analysis Report"

        st.markdown("---")
        st.markdown("### 🎨 Choose PPT Slide Template Design (10 PowerPoint Layouts)")
        
        # 10 Unique Structural PPT Templates List with previews
        templates_list = [
            ("Corporate Banner Blue", "https://images.unsplash.com/photo-1557804506-669a67965ba0?w=400&q=80"),
            ("Modern Side Border", "https://images.unsplash.com/photo-1460925895917-afdab827c52f?w=400&q=80"),
            ("Dark Tech Neon", "https://images.unsplash.com/photo-1526374965328-7f61d4dc18c5?w=400&q=80"),
            ("Executive Burgundy Header", "https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?w=400&q=80"),
            ("Engineering Steel Grid", "https://images.unsplash.com/photo-1581092160607-ee22621dd758?w=400&q=80"),
            ("Academic Emerald Clean", "https://images.unsplash.com/photo-1516321318423-f06f85e504b3?w=400&q=80"),
            ("Startup Pitch Card", "https://images.unsplash.com/photo-1559136555-9303baea8ebd?w=400&q=80"),
            ("Medical Healthcare Soft", "https://images.unsplash.com/photo-1576091160399-112ba8d25d1d?w=400&q=80"),
            ("Creative Vibrant Pink", "https://images.unsplash.com/photo-1533750349088-cd871a92f312?w=400&q=80"),
            ("Minimal Gray Structure", "https://images.unsplash.com/photo-1551288049-bebda4e38f71?w=400&q=80"),
        ]

        # Displaying 10 Templates Grid (5 cols x 2 rows)
        row1_cols = st.columns(5)
        for idx, (t_name, t_img) in enumerate(templates_list[:5]):
            with row1_cols[idx]:
                st.image(t_img, caption=t_name, use_container_width=True)
                if st.button(f"Select Layout {idx+1}", key=f"tpl_btn_{idx}"):
                    st.session_state["selected_tpl"] = t_name

        row2_cols = st.columns(5)
        for idx, (t_name, t_img) in enumerate(templates_list[5:]):
            with row2_cols[idx]:
                st.image(t_img, caption=t_name, use_container_width=True)
                if st.button(f"Select Layout {idx+6}", key=f"tpl_btn_{idx+5}"):
                    st.session_state["selected_tpl"] = t_name

        active_tpl = st.session_state.get("selected_tpl", "Corporate Banner Blue")
        st.info(f"Selected PowerPoint Template: **{active_tpl}**")

        if st.button("🚀 Generate Presentation Slides", key="ppt_gen_btn"):
            if not topic_input.strip():
                st.warning("⚠️ Topic field empty.")
            else:
                try:
                    with st.spinner(f"Creating PPT Deck with layout '{active_tpl}'..."):
                        ppt_prompt = f"Create a structured {slide_count}-slide presentation deck on '{topic_input}'. Template Style: {active_tpl}. Format slide by slide with titles and detailed bullet points."
                        res_text = call_gemini_ai(ppt_prompt)
                    render_output_section(res_text, f"{topic_input}_Slides", "tab_ppt", show_ppt=True, template_style=active_tpl)
                except Exception as e:
                    st.error(f"Execution Error: {e}")

    elif "From Document" in ppt_mode:
        ppt_file = st.file_uploader("Upload PDF or Word Document", type=["pdf", "docx", "txt"])
        if st.button("🚀 Convert Document to Presentation", key="ppt_doc_btn") and ppt_file:
            try:
                with st.spinner("Converting Document..."):
                    extracted_text = ""
                    if ppt_file.name.endswith(".pdf"):
                        reader = pypdf.PdfReader(ppt_file)
                        extracted_text = "\n".join([page.extract_text() or "" for page in reader.pages[:10]])
                    elif ppt_file.name.endswith(".docx"):
                        doc = docx.Document(ppt_file)
                        extracted_text = "\n".join([p.text for p in doc.paragraphs])
                    res_text = call_gemini_ai(f"Convert this document to slides:\n\n{extracted_text[:8000]}")
                render_output_section(res_text, "Document_Slides", "tab_ppt_doc", show_ppt=True)
            except Exception as e:
                st.error(f"Execution Error: {e}")

    elif "From Outline" in ppt_mode:
        outline_text = st.text_area("Paste Outline", height=150)
        if st.button("🚀 Build Slides from Outline", key="ppt_out_btn") and outline_text.strip():
            try:
                with st.spinner("Building Slides..."):
                    res_text = call_gemini_ai(f"Expand outline to slide deck:\n{outline_text}")
                render_output_section(res_text, "Outline_Slides", "tab_ppt_out", show_ppt=True)
            except Exception as e:
                st.error(f"Execution Error: {e}")

# TAB 2: CONTENT ASSISTANT
with tab2:
    st.subheader("✍️ Social Media Content Generator")
    c1, c2 = st.columns(2)
    with c1:
        platform = st.selectbox("Target Platform", ["LinkedIn", "Instagram", "Twitter / X", "Facebook"])
        content_type = st.selectbox("Content Style", ["Informational Post", "Promotional / Ad", "Storytelling"])
    with c2:
        tone = st.selectbox("Tone & Persona", ["Professional", "Casual & Friendly", "Persuasive"])
        target_audience = st.text_input("Target Audience", placeholder="e.g., Engineers, Students")
    
    topic = st.text_area("Core Brief / Topic")
    if st.button("🚀 Generate Content", key="tab2_btn") and topic and target_audience:
        try:
            res_text = call_gemini_ai(f"Platform: {platform}\nType: {content_type}\nTone: {tone}\nAudience: {target_audience}\nTopic: {topic}")
            render_output_section(res_text, f"{platform}_Content", "tab2", show_ppt=False)
        except Exception as e:
            st.error(f"Execution Error: {e}")

# TAB 3: TRANSLATOR
with tab3:
    st.subheader("🌐 Global Multi-Language Translator")
    target_language = st.selectbox("Target Language", ["English", "Urdu", "Arabic", "Hindi", "Spanish", "French", "German"])
    input_text = st.text_area("Source Text", height=150)
    if st.button("🌐 Translate Content", key="tab3_btn") and input_text.strip():
        try:
            res_text = call_gemini_ai(f"Translate to {target_language}:\n\n{input_text}")
            render_output_section(res_text, f"Translation_{target_language}", "tab3", show_ppt=False)
        except Exception as e:
            st.error(f"Execution Error: {e}")

# TAB 4: SMART WORKSPACE
with tab4:
    st.subheader("⚡ Smart File Workspace")
    user_prompt = st.text_area("Analysis Prompt", height=100)
    uploaded_files = st.file_uploader("Upload Documents", type=["pdf", "docx", "pptx", "txt"], accept_multiple_files=True)
    if st.button("⚡ Process Query", key="tab4_btn"):
        try:
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
            render_output_section(res_text, "Workspace_Output", "tab4", show_ppt=True)
        except Exception as e:
            st.error(f"Execution Error: {e}")

# TAB 5: ACADEMIC WRITER
with tab5:
    st.subheader("📚 Academic & Assignment Writer")
    subject_topic = st.text_input("Topic Header")
    academic_level = st.selectbox("Academic Level", ["Undergraduate", "Postgraduate / PhD", "College"])
    if st.button("✨ Generate Assignment", key="tab5_btn") and subject_topic.strip():
        try:
            res_text = call_gemini_ai(f"Write paper on '{subject_topic}' for {academic_level} level.")
            render_output_section(res_text, f"{subject_topic}_Assignment", "tab5", show_ppt=True)
        except Exception as e:
            st.error(f"Execution Error: {e}")

# TAB 6: DOCUMENT HUB
with tab6:
    st.subheader("📑 Advanced Document Hub")
    doc_topic = st.text_input("Document Topic")
    if st.button("✨ Build Document", key="tab6_btn") and doc_topic.strip():
        try:
            res_text = call_gemini_ai(f"Create a detailed document on '{doc_topic}'.")
            render_output_section(res_text, f"{doc_topic}_Document", "tab6", show_ppt=True)
        except Exception as e:
            st.error(f"Execution Error: {e}")

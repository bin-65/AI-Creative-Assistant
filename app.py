# TAB 1: PRESENTATION STUDIO
with tab1:
    st.markdown("## Create Slides with AI & Edit with Full Control")
    
    # Mode Selector
    ppt_mode = st.radio("Creation Mode", ["✨ From Topic", "📄 From Document", "📑 From Outline"], horizontal=True)
    
    st.markdown("---")
    
    if "From Topic" in ppt_mode:
        
        # Topic Input Box
        topic_input = st.text_area("Slide Topic / Title", value="Four stroke engine", height=100)
        
        # Options & Controls
        c1, c2, c3 = st.columns([2, 2, 2])
        with c1:
            use_search = st.checkbox("🔍 Enable Web Search Context", value=True)
        with c2:
            deep_think = st.checkbox("🧠 DeepThink Reasoning", value=True)
        with c3:
            slide_count = st.slider("Target Slide Count", 5, 20, 8)

        st.markdown("##### Quick Sample Topics:")
        q1, q2, q3, q4 = st.columns(4)
        if q1.button("Health Management"): topic_input = "Health Management Program"
        if q2.button("Summary Report"): topic_input = "Summary Report"
        if q3.button("Mental Health"): topic_input = "Mental Health Report"
        if q4.button("Data Analysis"): topic_input = "Data Analysis Report"

        st.markdown("---")
        st.markdown("### 🎨 Select Template Design")
        
        # Template Category Filters
        cat_col1, cat_col2, cat_col3, cat_col4 = st.columns(4)
        selected_category = "Featured"
        
        # Template Cards Grid with Visual Thumbnails
        t_col1, t_col2, t_col3 = st.columns(3)
        
        with t_col1:
            st.image("https://images.unsplash.com/photo-1557804506-669a67965ba0?w=500&q=80", caption="Corporate Annual Summary", use_container_width=True)
            t1_selected = st.button("Use 'Corporate Blue' Template", key="t1")
            
        with t_col2:
            st.image("https://images.unsplash.com/photo-1516321318423-f06f85e504b3?w=500&q=80", caption="Education & Training", use_container_width=True)
            t2_selected = st.button("Use 'Academic Minimal' Template", key="t2")
            
        with t_col3:
            st.image("https://images.unsplash.com/photo-1460925895917-afdab827c52f?w=500&q=80", caption="Marketing & Pitch Deck", use_container_width=True)
            t3_selected = st.button("Use 'Modern Pitch' Template", key="t3")

        # Set Selected Template state
        chosen_template = "Corporate Blue"
        if t2_selected:
            chosen_template = "Academic Minimal"
        elif t3_selected:
            chosen_template = "Modern Pitch"

        st.info(f" Selected Template Style: **{chosen_template}**")

        gen_ppt_btn = st.button("🚀 Generate Presentation Slides", key="ppt_gen_btn")

        if gen_ppt_btn:
            if not topic_input.strip():
                st.warning("⚠️ Topic field empty.")
            else:
                try:
                    with st.spinner("Generating Structured Presentation Deck..."):
                        ppt_prompt = f"Create a structured {slide_count}-slide presentation deck on '{topic_input}'. Design Template Style: {chosen_template}. Format slide by slide with titles and points."
                        res_text = call_gemini_ai(ppt_prompt)
                    render_output_section(res_text, f"{topic_input}_Slides", "tab_ppt", show_ppt=True)
                except Exception as e:
                    st.error(f"Execution Error: {e}")

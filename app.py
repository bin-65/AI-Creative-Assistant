def call_groq_ai(prompt_text):
    api_key = st.secrets.get("GROQ_API_KEY", "")
    if not api_key:
        raise Exception("GROQ_API_KEY Missing in Streamlit Secrets!")
    
    clean_key = str(api_key).strip().strip('"').strip("'")
    client = Groq(api_key=clean_key)
    
    # Active & Supported Groq Models
    models_to_try = [
        "llama-3.3-70b-versatile",
        "llama-3.1-8b-instant",
        "llama3-70b-8192",
        "llama3-8b-8192"
    ]
    
    last_err = ""
    for model_name in models_to_try:
        try:
            response = client.chat.completions.create(
                messages=[{"role": "user", "content": prompt_text}],
                model=model_name,
            )
            return response.choices[0].message.content
        except Exception as e:
            last_err = str(e)
            continue

    raise Exception(f"Groq API Execution Failed: {last_err}")

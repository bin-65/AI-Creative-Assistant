def call_groq_ai(prompt_text):
    # Retrieve key from Secrets or direct fallback
    api_key = st.secrets.get("GROQ_API_KEY", "")
    
    if not api_key:
        raise Exception("GROQ_API_KEY Streamlit Secrets mein nahi mili! Secrets settings check karein.")
    
    # Strip any spaces or unexpected quotes
    clean_key = str(api_key).strip().strip('"').strip("'")
    
    client = Groq(api_key=clean_key)
    
    # Currently Active Supported Models
    models_to_try = [
        "llama-3.3-70b-versatile",
        "llama-3.1-8b-instant",
        "llama3-70b-8192",
        "mixtral-8x7b-32768"
    ]
    
    last_error = ""
    for model_name in models_to_try:
        try:
            response = client.chat.completions.create(
                messages=[{"role": "user", "content": prompt_text}],
                model=model_name,
            )
            return response.choices[0].message.content
        except Exception as e:
            last_error = str(e)
            continue

    raise Exception(f"Groq API Error: {last_error}")

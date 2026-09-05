# Ultra-Stable Groq Call (Updated Models)
def call_groq_ai(prompt_text):
    api_key = st.secrets.get("GROQ_API_KEY")
    if not api_key:
        raise Exception("GROQ_API_KEY Missing in Streamlit Secrets!")
    
    client = Groq(api_key=api_key.strip())
    
    # Active & Verified Groq Models
    models_to_try = [
        "llama-3.1-8b-instant",
        "llama3-70b-8192",
        "llama3-8b-8192",
        "mixtral-8x7b-32768"
    ]
    
    for model_name in models_to_try:
        try:
            response = client.chat.completions.create(
                messages=[{"role": "user", "content": prompt_text}],
                model=model_name,
            )
            return response.choices[0].message.content
        except Exception:
            continue

    raise Exception("Groq API Call Failed on all available models. Please verify API key.")

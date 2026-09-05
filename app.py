def call_groq_ai(prompt_text):
    # Retrieve Key
    api_key = st.secrets.get("GROQ_API_KEY", "")
    
    if not api_key:
        raise Exception("Streamlit Secrets mein 'GROQ_API_KEY' nahi mila. Settings check karein.")
    
    clean_key = str(api_key).strip().strip('"').strip("'")
    
    try:
        client = Groq(api_key=clean_key)
        # Using verified active model
        response = client.chat.completions.create(
            messages=[{"role": "user", "content": prompt_text}],
            model="llama-3.1-8b-instant",
        )
        return response.choices[0].message.content
    except Exception as err:
        raise Exception(f"Groq Rejected Key: {err}")

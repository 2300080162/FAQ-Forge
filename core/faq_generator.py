import google.generativeai as genai

genai.configure(api_key="AIzaSyBLNyZUlIjOQdVD9EcGk4TWAA_v8XHvYA4")

def generate(context, topic, difficulty, n):
    prompt = f"""
Using ONLY this content:

{context}

Generate {n} FAQs about "{topic}"
Difficulty: {difficulty}

Format:
Q:
A:
"""

    model = genai.GenerativeModel("gemini-1.5-flash")
    response = model.generate_content(prompt)
    return response.text
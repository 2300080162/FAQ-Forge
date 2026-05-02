import google.generativeai as genai
import numpy as np

# 🔑 Paste your REAL Gemini API key
genai.configure(api_key="AIzaSyBLNyZUlIjOQdVD9EcGk4TWAA_v8XHvYA4")

def embed(text):
    result = genai.embed_content(
        model="embedding-001",
        content=text
    )
    return np.array(result["embedding"], dtype="float32")
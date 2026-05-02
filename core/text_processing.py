import re

def clean_and_chunk(text, size=500):
    text = re.sub(r'\s+', ' ', text)
    return [text[i:i+size] for i in range(0, len(text), size)]
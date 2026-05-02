from pypdf import PdfReader

def load_text(uploaded_file, pasted_text):
    if uploaded_file:
        reader = PdfReader(uploaded_file)
        return " ".join(page.extract_text() for page in reader.pages)
    return pasted_text
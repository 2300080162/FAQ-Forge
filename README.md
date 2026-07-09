# 🤖 IntelliFAQ – AI-Powered Automated FAQ Builder

An AI-powered FAQ generation system that automatically creates structured Frequently Asked Questions (FAQs) from product documentation using **Retrieval-Augmented Generation (RAG)**, **Google Gemini API**, and **FAISS**.

The application enables organizations to quickly build knowledge bases, reduce manual documentation effort, and improve customer support through intelligent document analysis.

---

## 📌 Features

- 📄 Upload product documentation in PDF format
- 🤖 AI-powered FAQ generation using Gemini API
- 🔍 RAG-based document retrieval with FAISS
- 🎯 Customizable FAQ generation
  - Difficulty Level
  - Number of Questions
  - FAQ Topic
- 📚 Intelligent document chunking and semantic search
- 📑 Certificate PDF analysis
- 📅 Automatic certificate validity calculation
- 📊 Certificate progress tracking
- 💻 Interactive Streamlit Web Interface

---

## 🏗️ System Architecture

```
User Upload
      │
      ▼
Document Processing
      │
      ▼
Text Chunking
      │
      ▼
Sentence Embeddings
      │
      ▼
FAISS Vector Database
      │
      ▼
RAG Retrieval
      │
      ▼
Gemini AI
      │
      ▼
FAQ Generation
      │
      ▼
Cleaned FAQ Output
```

---

## 🚀 Technologies Used

| Category | Technologies |
|----------|--------------|
| Programming Language | Python |
| Frontend | Streamlit |
| AI Model | Google Gemini API |
| Vector Database | FAISS |
| Embedding Model | Sentence Transformers |
| NLP | Natural Language Processing (NLP) |
| Data Processing | Pandas |
| PDF Processing | PyPDF2 |
| Document Processing | python-docx |

---

## ⚙️ Installation

Clone the repository

```bash
git clone https://github.com/yourusername/IntelliFAQ.git
```

Move into the project

```bash
cd IntelliFAQ
```

Create Virtual Environment

```bash
python -m venv venv
```

Activate Virtual Environment

### Windows

```bash
venv\Scripts\activate
```

### Linux / macOS

```bash
source venv/bin/activate
```

Install dependencies

```bash
pip install -r requirements.txt
```

---

## 🔑 Configure Gemini API

Create a `.env` file

```env
GEMINI_API_KEY=YOUR_API_KEY
```

---

## ▶️ Run the Application

```bash
streamlit run app.py
```

---

## 📂 Project Structure

```
IntelliFAQ
│
├── app.py
├── rag_engine.py
├── generator.py
├── cleaner.py
├── certificate_validator.py
├── requirements.txt
├── README.md
└── assets
```

---

## 📊 Project Workflow

1. Upload Product Documentation
2. Extract Document Text
3. Create Text Embeddings
4. Store Embeddings in FAISS
5. Retrieve Relevant Content using RAG
6. Generate FAQs with Gemini API
7. Clean and Format Output
8. Display Results in Streamlit
9. Analyze Certificate PDF
10. Calculate Certificate Validity & Progress

---

## 🎯 Applications

- Customer Support
- Knowledge Base Generation
- SaaS Documentation
- Product Documentation
- Employee Training
- Technical Documentation
- Certificate Management

---

## 🌟 Future Enhancements

- Multi-language FAQ Generation
- DOCX, PPTX & HTML Support
- Chatbot Integration
- Cloud Deployment
- Voice-based FAQ Assistant
- Admin Dashboard
- Authentication System

---

## 👨‍💻 Author

**Danda Ramakanth Reddy**

B.Tech – Artificial Intelligence & Data Science

Interested in:

- Artificial Intelligence
- Machine Learning
- Deep Learning
- Natural Language Processing
- Generative AI
- Computer Vision

---

## ⭐ Support

If you found this project helpful,

⭐ Star this repository

🍴 Fork the repository

💡 Feel free to contribute and improve the project.

---

## 📜 License

This project is developed for educational and research purposes.

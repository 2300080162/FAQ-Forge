from core.embeddings import embed

def retrieve(topic, db, k=8):
    q = genai.embed_content(
        model="models/text-embedding-004",
        content=topic,
        task_type="retrieval_query"
    )
    query_vector = np.array(q["embedding"], dtype="float32").reshape(1, -1)
    return db.search(query_vector, k)
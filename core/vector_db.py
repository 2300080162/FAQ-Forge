import faiss

class VectorDB:
    def __init__(self, dim):
        self.index = faiss.IndexFlatL2(dim)
        self.data = []

    def add(self, embeddings, chunks):
        self.index.add(embeddings)
        self.data.extend(chunks)

    def search(self, query, k):
        _, idx = self.index.search(query, k)
        return [self.data[i] for i in idx[0]]
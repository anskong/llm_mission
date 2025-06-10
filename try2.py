from sentence_transformers import SentenceTransformer
if __name__ == "__main__":	
    model = SentenceTransformer("jhgan/ko-sroberta-multitask")
    sentences = [
        "The weather is lovely today.",
        "It's so sunny outside!",
        "He drove to the stadium."
    ]
    embeddings = model.encode(sentences)

    similarities = model.similarity(embeddings, embeddings)
    print(similarities.shape)
# [3, 3]
    # import langchain_community.document_loaders as loaders
    # print(dir(loaders))

# pip install langchain-huggingface sentence-transformers
# from langchain_huggingface.embeddings import HuggingFaceEmbeddings
# hf = HuggingFaceEmbeddings(model_name="sentence-transformers/all-MiniLM-L6-v2")



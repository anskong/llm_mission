from sentence_transformers import SentenceTransformer
if __name__ == "__main__2":	
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

if __name__ == "__main__1":
    from langchain_huggingface import HuggingFaceEmbeddings

    ko_embedding = HuggingFaceEmbeddings(
        model_name=r"C:\ai_dev\ko-sroberta-multitask"
    )

    embedding = ko_embedding.embed_query("안녕하세요, 임베딩 테스트입니다.")
    print(embedding[:10])  # 일부만 출력

if __name__ == "__main__":
    from sentence_transformers import SentenceTransformer

    model = SentenceTransformer('jhgan/ko-sroberta-multitask')  # 인터넷 가능 환경
    model.save("C:/ai_dev/ko-sroberta-multitask")  # 로컬에 저장

    from langchain_huggingface import HuggingFaceEmbeddings

    ko_embedding = HuggingFaceEmbeddings(
        model_name=r"C:\ai_dev\ko-sroberta-multitask"
    )

    embedding = ko_embedding.embed_query("안녕하세요, 임베딩 테스트입니다.")
    print(embedding[:10])  # 일부만 출력
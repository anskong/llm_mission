if __name__ == "__main__":
    import langchain_community.document_loaders as loaders
    print(dir(loaders))

# pip install langchain-huggingface sentence-transformers
# from langchain_huggingface.embeddings import HuggingFaceEmbeddings
# hf = HuggingFaceEmbeddings(model_name="sentence-transformers/all-MiniLM-L6-v2")
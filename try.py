
import numpy as np
from tkinter import Tk, filedialog, messagebox
import json
#코사인 유사도 (두 벡터 간의 유사도 )
def cos_similarty(a,b):
    result = np.dot(a,b)
    return result

# pip install langchain-core langchain-community
# pip install faiss-cpu
# pip install openai  # or use other embedding models
# pip install python-pptx openpyxl PyMuPDF
# pip install "unstructured[all-docs]"  # for PPT/Excel loader
from tkinter import Tk, filedialog
import os

from langchain_community.document_loaders import (
    PyPDFLoader,
    UnstructuredPowerPointLoader,
    UnstructuredExcelLoader,
)
from langchain_community.vectorstores import FAISS
from langchain_openai import OpenAIEmbeddings

from langchain.text_splitter import RecursiveCharacterTextSplitter

# ✅ GUI로 여러 파일 선택
def select_multiple_files():
    root = Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(
        title="PDF, PPTX, XLSX 파일 선택",
        filetypes=[
            ("Supported files", "*.pdf *.pptx *.xlsx"),
            ("All files", "*.*"),
        ],
    )
    return list(file_paths)

# ✅ 확장자에 따른 LangChain Loader 선택
def load_documents(file_paths):
    all_docs = []
    for path in file_paths:
        ext = os.path.splitext(path)[1].lower()
        if ext == ".pdf":
            loader = PyPDFLoader(path)
        elif ext == ".pptx":
            loader = UnstructuredPowerPointLoader(path)
        elif ext == ".xlsx":
            loader = UnstructuredExcelLoader(path)
        else:
            print(f"⚠️ 지원하지 않는 형식: {ext}")
            continue
        docs = loader.load()
        all_docs.extend(docs)
    return all_docs

# ✅ 문서 분할
def split_documents(documents):
    splitter = RecursiveCharacterTextSplitter(chunk_size=1000, chunk_overlap=200)
    return splitter.split_documents(documents)

# ✅ 벡터스토어에 저장 (FAISS)
def embed_and_store(documents, persist_path="faiss_index"):
    embeddings = OpenAIEmbeddings()  # OpenAI API 키 필요 (환경변수 OPENAI_API_KEY)
    vectorstore = FAISS.from_documents(documents, embeddings)
    vectorstore.save_local(persist_path)
    print(f"✅ FAISS 벡터스토어 저장 완료: {persist_path}")
    return vectorstore

# ✅ 전체 워크플로우
if __name__ == "__main__mm":
    file_paths = select_multiple_files()
    if not file_paths:
        print("❗파일이 선택되지 않았습니다.")
        exit()

    print(f"📂 선택된 파일: {file_paths}")
    documents = load_documents(file_paths)
    print(f"📄 총 로딩된 문서 수: {len(documents)}")

    split_docs = split_documents(documents)
    print(f"✂️ 분할된 청크 수: {len(split_docs)}")

    embed_and_store(split_docs)



def open_file():
    # 파일 선택기 열기
    root = Tk()
    root.withdraw()

    contents_path = filedialog.askopenfilename(title="TXT 파일 선택", filetypes=[("TXT files", "*.txt")])

    if not contents_path:
        messagebox.showwarning("경고", "목차 파일을 선택하지 않았습니다.")
        return None
    else:
        print(f"선택한 파일 명 :::: {contents_path}")
        with open(contents_path, 'r', encoding='utf-8') as file:
            contents = json.load(file)
            print(f"목차 내용 ::: {contents}")
            # return contents
    
    file_paths = filedialog.askopenfilenames(
        title="PDF, PPTX, XLSX 파일 선택",
        filetypes=[
            ("Supported files", "*.pdf *.pptx *.xlsx *.docx"),
            ("All files", "*.*"),
        ],
    )

    print("선택 된 파일 경로들:")
    for file_path in file_paths:
        print(file_path)

    return
    
if __name__ == "__main__1":
    import langchain_community.document_loaders as loaders
#   open_file()

if __name__ == "__main__":
    from langchain_openai import OpenAIEmbeddings, ChatOpenAI
    llm = ChatOpenAI(
        openai_api_base="http://localhost:1234/v1",
        openai_api_key="lm-studio",
        model_name="exaone-3.5-2.4b-instruct",
        temperature=0.7,
    )

    result = llm.invoke("너는 누구니??")
    print(result.content)
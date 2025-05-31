# pdf_full_loader.py

from typing import List
from langchain_core.document_loaders import BaseLoader
from langchain_core.documents import Document
from langchain.text_splitter import RecursiveCharacterTextSplitter
import pdfplumber
import fitz
from pdf2image import convert_from_path
import pytesseract

class PDFFullLoader(BaseLoader):
    def __init__(self, pdf_path: str, lang: str = "eng+kor", chunk_size=500, chunk_overlap=50):
        self.pdf_path = pdf_path
        self.lang = lang
        self.chunk_size = chunk_size
        self.chunk_overlap = chunk_overlap

    def load(self) -> List[Document]:
        raw_docs = []

        # 텍스트 + 표 추출
        with pdfplumber.open(self.pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                text = page.extract_text() or ""
                tables = page.extract_tables()
                if tables:
                    table_text = ""
                    for table in tables:
                        for row in table:
                            row_text = "\t".join([cell or "" for cell in row])
                            table_text += row_text + "\n"
                    text += "\n[표 내용]\n" + table_text

                if text.strip():
                    raw_docs.append(Document(page_content=text.strip(), metadata={"page": i + 1, "type": "text+table"}))

        # 이미지 OCR
        # images = convert_from_path(self.pdf_path)
        # for i, image in enumerate(images):
        #     ocr_text = pytesseract.image_to_string(image, lang=self.lang)
        #     if ocr_text.strip():
        #         raw_docs.append(Document(page_content="[OCR 이미지 텍스트]\n" + ocr_text.strip(), metadata={"page": i + 1, "type": "image_ocr"}))

        # 문서 분할
        splitter = RecursiveCharacterTextSplitter(chunk_size=self.chunk_size, chunk_overlap=self.chunk_overlap)
        return splitter.split_documents(raw_docs)
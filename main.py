from abc import ABC, abstractmethod
import os
import pdfplumber
import docx
import pptx
import mysql.connector
import csv
from PIL import Image
from io import BytesIO
import json
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml.ns import qn
from docx.document import Document as DocxDocument
import fitz  # PyMuPDF for PDF and font extraction

# Abstract Class: FileLoader
class FileLoader(ABC):
    def __init__(self, file_path):
        self.file_path = file_path
        self.validate_file()

    @abstractmethod
    def validate_file(self):
        pass

    @abstractmethod
    def load_file(self):
        pass

# Concrete Class: PDFLoader
class PDFLoader(FileLoader):
    def validate_file(self):
        if not self.file_path.endswith('.pdf'):
            raise ValueError("Invalid file type. Expected a PDF file.")
    
    def load_file(self):
        return pdfplumber.open(self.file_path)

# Concrete Class: DOCXLoader
class DOCXLoader(FileLoader):
    def validate_file(self):
        if not self.file_path.endswith('.docx'):
            raise ValueError("Invalid file type. Expected a DOCX file.")
    
    def load_file(self):
        return docx.Document(self.file_path)

# Concrete Class: PPTLoader
class PPTLoader(FileLoader):
    def validate_file(self):
        if not self.file_path.endswith('.pptx'):
            raise ValueError("Invalid file type. Expected a PPTX file.")
    
    def load_file(self):
        return pptx.Presentation(self.file_path)



class DataExtractor:
    def __init__(self, loader):
        self.loader = loader.load_file()
        self.file_path = loader.file_path
        self.file_name = os.path.basename(self.file_path)
        self.metadata = self.get_metadata()

    def get_metadata(self):
        file_stats = os.stat(self.file_path)
        return {
            "file_size": file_stats.st_size,
            "creation_time": file_stats.st_ctime,
            "modification_time": file_stats.st_mtime
        }

    def extract_text(self):
        results = []

        def is_heading(font_size, bold):
            try:
                fs = float(font_size)
            except (ValueError, TypeError):
                fs = 0
            return bold or fs > 12  # adjust threshold as needed

        # PDF Extraction using fitz
        if self.file_path.endswith('.pdf'):
            doc = fitz.open(self.file_path)
            for i, page in enumerate(doc):
                blocks = page.get_text("dict")["blocks"]
                for block in blocks:
                    if block["type"] == 0:  # text block
                        for line in block["lines"]:
                            for span in line["spans"]:
                                text = span.get('text', '').strip()
                                if not text:
                                    continue
                                font_name = span.get("font", "Default")
                                font_size = span.get("size", 0)
                                bold = "Bold" in font_name
                                italic = False  # fitz does not provide italic info directly
                                data_type = "heading" if is_heading(font_size, bold) else "text"
                                results.append((i + 1, text, data_type, font_name, font_size, bold, italic))
            return results

        # DOCX Extraction using python-docx
        elif self.file_path.endswith('.docx'):
            doc = self.loader  # already loaded as a Document
            for para in doc.paragraphs:
                for run in para.runs:
                    text = run.text.strip()
                    if not text:
                        continue
                    font_name = run.font.name if run.font.name else "Default"
                    font_size = run.font.size.pt if run.font.size else 0
                    bold = run.bold if run.bold is not None else False
                    italic = run.italic if run.italic is not None else False
                    data_type = "heading" if is_heading(font_size, bold) else "text"
                    results.append((1, text, data_type, font_name, font_size, bold, italic))
            return results

        # PPTX Extraction using python-pptx
        elif self.file_path.endswith('.pptx'):
            pres = self.loader  # already loaded as a Presentation
            for i, slide in enumerate(pres.slides):
                for shape in slide.shapes:
                    if hasattr(shape, "text_frame") and shape.text_frame:
                        for para in shape.text_frame.paragraphs:
                            for run in para.runs:
                                text = run.text.strip()
                                if not text:
                                    continue
                                font = run.font
                                font_name = font.name if font.name else "Default"
                                font_size = font.size.pt if font.size else 0
                                bold = font.bold if font.bold is not None else False
                                italic = font.italic if font.italic is not None else False
                                data_type = "heading" if is_heading(font_size, bold) else "text"
                                results.append((i + 1, text, data_type, font_name, font_size, bold, italic))
            return results

        return results

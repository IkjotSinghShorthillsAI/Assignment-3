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
    def extract_links(self):
            links = []
            
            if isinstance(self.loader, pdfplumber.PDF):
                for i, page in enumerate(self.loader.pages):
                    if hasattr(page, 'annots') and page.annots:
                        for annot in page.annots:
                            uri = annot.get("uri")
                            if uri:
                                links.append((i + 1, uri))
            
            elif isinstance(self.loader, DocxDocument):
                for para in self.loader.paragraphs:
                    for rel in self.loader.part.rels.values():
                        if "hyperlink" in rel.reltype:
                            if para.text and rel.target_ref in para.text:
                                links.append((1, rel.target_ref))
            
            elif hasattr(self.loader, "slides"):  # Check if it's a Presentation object
                for slide_num, slide in enumerate(self.loader.slides):
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            for paragraph in shape.text_frame.paragraphs:
                                for run in paragraph.runs:
                                    if run.hyperlink and run.hyperlink.address:
                                        links.append((slide_num + 1, run.hyperlink.address))
            
            return links

    def extract_images(self):
        images = []
        
        if self.file_path.endswith(".pdf"):
            doc = fitz.open(self.file_path)
            for page_num in range(len(doc)):
                for img_index, img in enumerate(doc[page_num].get_images(full=True)):
                    xref = img[0]
                    base_image = doc.extract_image(xref)
                    image_bytes = base_image["image"]
                    img_obj = Image.open(BytesIO(image_bytes))
                    images.append((page_num + 1, "PNG", img_obj.size, img_obj))
        
        elif self.file_path.endswith(".docx"):
            for i, rel in enumerate(self.loader.part.rels):
                if "image" in self.loader.part.rels[rel].target_ref:
                    image_data = self.loader.part.rels[rel].target_part.blob
                    img_obj = Image.open(BytesIO(image_data))
                    images.append((i + 1, "PNG", img_obj.size, img_obj))
        
        elif self.file_path.endswith(".pptx"):
            for slide_num, slide in enumerate(self.loader.slides):
                for shape in slide.shapes:
                    if shape.shape_type == 13:  # Picture shape type
                        image = shape.image
                        img_obj = Image.open(BytesIO(image.blob))
                        images.append((slide_num + 1, "PNG", img_obj.size, img_obj))
        
        return images

    def extract_tables(self):
        tables = []
        if isinstance(self.loader, pdfplumber.PDF):
            for i, page in enumerate(self.loader.pages):
                extracted_tables = page.extract_tables()
                for table in extracted_tables:
                    tables.append((i + 1, len(table), len(table[0]) if table else 0, table))
        
        elif isinstance(self.loader, DocxDocument):
            for i, table in enumerate(self.loader.tables):
                extracted_table = [[cell.text.strip() for cell in row.cells] for row in table.rows]
                tables.append((i + 1, len(extracted_table), len(extracted_table[0]) if extracted_table else 0, extracted_table))
        
        elif hasattr(self.loader, "slides"):  # Check if it's a Presentation object
            for slide_num, slide in enumerate(self.loader.slides):
                for shape in slide.shapes:
                    if shape.has_table:
                        table = shape.table
                        extracted_table = [[cell.text.strip() for cell in row.cells] for row in table.rows]
                        tables.append((slide_num + 1, len(extracted_table), len(extracted_table[0]) if extracted_table else 0, extracted_table))
        
        return tables
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

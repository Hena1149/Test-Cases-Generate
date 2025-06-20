import os
import re
import logging
from io import BytesIO
from typing import List, Optional
from zipfile import ZipFile, BadZipFile
from pypdf import PdfReader  # Remplacement de fitz
import pandas as pd
from docx import Document
from openpyxl import Workbook
from xlsxwriter.workbook import Workbook as XlsxWorkbook

# Configuration du logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def extract_text_from_pdf(file_path: str) -> str:
    """Extrait le texte d'un PDF avec PyPDF (solution plus stable que fitz)."""
    try:
        text = ""
        reader = PdfReader(file_path)
        for page in reader.pages:
            text += page.extract_text() or ""
        return text.strip()
    except Exception as e:
        logger.error(f"Erreur d'extraction PDF : {str(e)}")
        raise ValueError(f"Impossible de lire le PDF : {str(e)}")

def is_valid_docx(file_path: str) -> bool:
    """Vérifie si un fichier DOCX est valide."""
    try:
        with ZipFile(file_path) as z:
            return 'word/document.xml' in z.namelist()
    except (BadZipFile, KeyError):
        return False

def extract_text_from_docx(file_path: str) -> str:
    """Extrait le texte d'un DOCX avec gestion d'erreurs améliorée."""
    if not is_valid_docx(file_path):
        raise ValueError("Fichier DOCX corrompu ou invalide")
    
    try:
        doc = Document(file_path)
        return "\n".join(para.text for para in doc.paragraphs if para.text.strip())
    except Exception as e:
        logger.error(f"Erreur DOCX : {str(e)}")
        raise ValueError(f"Erreur de lecture DOCX : {str(e)}")

def extract_text_from_txt(file_path: str) -> str:
    """Lit un fichier texte avec gestion d'encodage."""
    encodings = ['utf-8', 'latin-1']
    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    raise ValueError("Échec de décodage du fichier texte")

def process_uploaded_file(file_path: str) -> str:
    """
    Traite un fichier uploadé (PDF/DOCX/TXT) et retourne son texte.
    Lève une ValueError en cas d'échec.
    """
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext == ".pdf":
            return extract_text_from_pdf(file_path)
        elif ext == ".docx":
            return extract_text_from_docx(file_path)
        elif ext == ".txt":
            return extract_text_from_txt(file_path)
        else:
            raise ValueError("Format non supporté. Utilisez PDF, DOCX ou TXT.")
    except Exception as e:
        logger.error(f"Erreur avec {file_path} : {str(e)}")
        raise

def export_to_excel(data: List[str], sheet_name: str = "Data") -> BytesIO:
    """Exporte une liste de textes vers un fichier Excel en mémoire."""
    output = BytesIO()
    df = pd.DataFrame(data, columns=["Contenu"])
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        worksheet = writer.sheets[sheet_name]
        worksheet.set_column(0, 0, 50)  # Ajuste la largeur de colonne
    
    output.seek(0)
    return output

def export_test_cases_to_excel(test_cases: List[str]) -> BytesIO:
    """Exporte des cas de test structurés vers Excel."""
    structured_data = []
    pattern = re.compile(r'###\s*(.*?)\s*\n(.*?)(?=###|$)', re.DOTALL)
    
    for i, test_case in enumerate(test_cases, 1):
        sections = {k.lower(): v.strip() for k, v in pattern.findall(test_case)}
        structured_data.append({
            "ID": f"TEST-{i}",
            "Titre": sections.get("titre", ""),
            "Préconditions": sections.get("préconditions", ""),
            "Étapes": sections.get("étapes", ""),
            "Résultat attendu": sections.get("résultat attendu", "")
        })
    
    output = BytesIO()
    df = pd.DataFrame(structured_data)
    df.to_excel(output, index=False, engine='xlsxwriter')
    output.seek(0)
    return output
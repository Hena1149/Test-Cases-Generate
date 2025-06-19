import fitz  # PyMuPDF
import docx
from typing import List, Optional
import pandas as pd
from io import BytesIO
import re
import os
from zipfile import ZipFile, BadZipFile
import tempfile
import logging

# Configuration du logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def extract_text_from_pdf(file_path: str) -> str:
    """Extrait le texte d'un fichier PDF de manière robuste."""
    try:
        text = ""
        with fitz.open(file_path) as doc:
            for page in doc:
                text += page.get_text() or ""  # Gestion des pages sans texte
        return text.strip()
    except Exception as e:
        logger.error(f"Erreur d'extraction PDF: {str(e)}")
        raise ValueError(f"Impossible de lire le fichier PDF: {str(e)}")

def is_valid_docx(file_path: str) -> bool:
    """Vérifie si un fichier DOCX est valide."""
    try:
        with ZipFile(file_path) as z:
            if 'word/document.xml' not in z.namelist():
                return False
            # Test de lecture du contenu
            with z.open('word/document.xml') as f:
                f.read(100)  # Lecture des 100 premiers octets
        return True
    except (BadZipFile, KeyError, IOError):
        return False

def extract_text_from_docx(file_path: str) -> str:
    """Extrait le texte d'un fichier Word avec gestion des erreurs améliorée."""
    try:
        # Vérification préalable du fichier
        if not is_valid_docx(file_path):
            raise BadZipFile("Fichier DOCX corrompu ou invalide")

        # Tentative d'extraction standard
        doc = docx.Document(file_path)
        text = "\n".join([para.text for para in doc.paragraphs if para.text.strip()])
        
        if not text.strip():  # Si le document semble vide
            text = backup_docx_extraction(file_path) or text
            
        return text.strip()
    except Exception as e:
        logger.error(f"Erreur d'extraction DOCX: {str(e)}")
        # Tentative de récupération alternative
        backup_text = backup_docx_extraction(file_path)
        if backup_text:
            return backup_text
        raise ValueError(f"Impossible de lire le fichier DOCX: {str(e)}")

def backup_docx_extraction(file_path: str) -> Optional[str]:
    """Méthode alternative pour extraire le texte des DOCX corrompus."""
    try:
        from xml.etree import ElementTree as ET
        
        with ZipFile(file_path) as z:
            with z.open('word/document.xml') as f:
                tree = ET.parse(f)
                root = tree.getroot()
                
                # Namespace pour les documents Word
                ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                text_elements = root.findall('.//w:t', ns)
                
                return ' '.join([t.text for t in text_elements if t.text]).strip()
    except Exception:
        return None

def extract_text_from_txt(file_path: str) -> str:
    """Extrait le texte d'un fichier TXT avec gestion d'encodage."""
    try:
        with open(file_path, "r", encoding='utf-8') as f:
            return f.read()
    except UnicodeDecodeError:
        try:
            with open(file_path, "r", encoding='latin-1') as f:
                return f.read()
        except Exception as e:
            raise ValueError(f"Impossible de lire le fichier texte: {str(e)}")

def process_uploaded_file(file_path: str) -> str:
    """
    Traite le fichier uploadé selon son type avec gestion robuste des erreurs.
    
    Args:
        file_path: Chemin vers le fichier temporaire
        
    Returns:
        str: Texte extrait du fichier
        
    Raises:
        ValueError: Si le fichier est corrompu ou le format non supporté
    """
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError("Fichier temporaire introuvable")
        
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext == ".pdf":
            return extract_text_from_pdf(file_path)
        elif file_ext == ".docx":
            return extract_text_from_docx(file_path)
        elif file_ext == ".txt":
            return extract_text_from_txt(file_path)
        else:
            raise ValueError("Type de fichier non supporté. Formats acceptés: PDF, DOCX, TXT")
    except Exception as e:
        logger.error(f"Erreur de traitement du fichier {file_path}: {str(e)}")
        raise ValueError(f"Erreur de traitement du fichier: {str(e)}")

def export_to_excel(data: List[str], sheet_name: str = "Data") -> BytesIO:
    """Convertit une liste de textes en fichier Excel avec gestion des erreurs."""
    try:
        if not data:
            raise ValueError("Aucune donnée à exporter")
            
        output = BytesIO()
        df = pd.DataFrame(data, columns=["Contenu"])
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            
            # Formatage automatique des colonnes
            for idx, col in enumerate(df.columns):
                max_len = max(
                    df[col].astype(str).map(len).max(),  # Longueur max des données
                    len(col)  # Longueur de l'en-tête
                ) + 2  # Marge
                worksheet.set_column(idx, idx, min(max_len, 50))  # Limite à 50 caractères
        
        output.seek(0)
        return output
    except Exception as e:
        logger.error(f"Erreur d'export Excel: {str(e)}")
        raise ValueError(f"Erreur lors de la génération du fichier Excel: {str(e)}")

def export_test_cases_to_excel(test_cases: List[str]) -> BytesIO:
    """Exporte les cas de test structurés vers Excel avec gestion robuste."""
    try:
        if not test_cases:
            raise ValueError("Aucun cas de test à exporter")
            
        output = BytesIO()
        data = []
        
        # Expression régulière optimisée pour l'extraction des sections
        section_pattern = re.compile(
            r'###\s*(Titre|Préconditions|Données d\'entrée|Étapes|Résultat attendu)\s*\n(.*?)(?=###|$)', 
            re.DOTALL
        )
        
        for i, test_case in enumerate(test_cases, 1):
            case_data = {
                "ID": f"TEST-{i}",
                "Titre": "",
                "Préconditions": "",
                "Données d'entrée": "",
                "Étapes": "",
                "Résultat attendu": ""
            }
            
            # Extraction des sections avec la regex
            sections = section_pattern.findall(test_case)
            for section in sections:
                section_name, content = section
                content = content.strip()
                
                if section_name == "Titre":
                    case_data["Titre"] = content
                elif section_name == "Préconditions":
                    case_data["Préconditions"] = content
                elif section_name == "Données d'entrée":
                    case_data["Données d'entrée"] = content
                elif section_name == "Étapes":
                    case_data["Étapes"] = content
                elif section_name == "Résultat attendu":
                    case_data["Résultat attendu"] = content
            
            data.append(case_data)
        
        # Création du DataFrame avec vérification des données
        df = pd.DataFrame(data)
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name="Cas_de_test", index=False)
            
            # Formatage avancé
            workbook = writer.book
            worksheet = writer.sheets["Cas_de_test"]
            
            # Format des cellules
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#4472C4',
                'font_color': 'white',
                'border': 1
            })
            
            # Appliquer le format aux en-têtes
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Ajustement automatique des largeurs
            for idx, col in enumerate(df.columns):
                max_len = max(
                    df[col].astype(str).map(len).max(),
                    len(col)
                ) + 2
                worksheet.set_column(idx, idx, min(max_len, 100))  # Limite à 100 caractères
        
        output.seek(0)
        return output
        
    except Exception as e:
        logger.error(f"Erreur d'export des cas de test: {str(e)}")
        raise ValueError(f"Erreur lors de la génération des cas de test: {str(e)}")

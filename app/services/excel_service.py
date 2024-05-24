import logging
import unicodedata
import pandas as pd
from fastapi import HTTPException
from app.core.config import settings
from app.services.word_service import generate_word_document
import os

# Configure logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

def normalize_string(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn').lower()


def process_excel_file(file_path: str, output_dir: str) -> list:
    try:
        logger.debug("Chargement du fichier Excel.")
        df_titles = pd.read_excel(file_path, header=None)
        titles_row = df_titles.iloc[0, 2:22].tolist()  # Adjusted to C1 to V1
        
        df_students = pd.read_excel(file_path, header=1)
        df_students = df_students.rename(columns={
            'DatedeNaissance': 'Date de Naissance',
            'NomSite': 'Nom Site',
            'CodeGroupe': 'Code Groupe',
            'NomGroupe': 'Nom Groupe',
            'EtenduGroupe': 'Étendu Groupe',
            'ABSjustifiées': 'ABS justifiées',
            'ABSinjustifiées': 'ABS injustifiées',
        })
        logger.debug(f"{len(df_students)} étudiants trouvés dans le fichier.")
        
        templates = {
            "MAPI": settings.M1_S1_MAPI_TEMPLATE_WORD,
            "MAGI": settings.M1_S1_MAGI_TEMPLATE_WORD,
            "MEFIM": settings.M1_S1_MEFIM_TEMPLATE_WORD
        }
        
        matching_values = {
            "MAPI": ['UE 1 – Economie & Gestion', 'Stratégie et Solutions Immobilières', 'Finance Immobilière', 'Economie Immobilière I', 'UE 2 – Droit', 'Droit des Affaires et des Contrats', 'UE 3 – Aménagement & Urbanisme', 'Ville et Développements Urbains', "Politique de l'Habitat", 'UE 4 – Compétences Professionnalisantes', 'Real Estate English', "Les Rencontres de l'Immobilier", 'ESPI Career Services', 'ESPI Inside', 'Immersion Professionnelle', 'Projet Voltaire', 'UE SPE – MAPI', 'Etude Foncière', "Montage d'une Opération de Promotion Immobilière", 'Acquisition et Dissociation du Foncier'],
            "MAGI": ['UE 1 – Economie & Gestion', 'Stratégie et Solutions Immobilières', 'Finance Immobilière', 'Economie Immobilière I', 'UE 2 – Droit', 'Droit des Affaires et des Contrats', 'UE 3 – Aménagement & Urbanisme', 'Ville et Développements Urbains', "Politique de l'Habitat", 'UE 4 – Compétences Professionnalisantes', 'Real Estate English', "Les Rencontres de l'Immobilier", 'ESPI Career Services', 'ESPI Inside', 'Immersion Professionnelle', 'Projet Voltaire', 'UE SPE – MAGI', 'Baux Commerciaux et Gestion Locative', 'Actifs Tertiaires en Copropriété', 'Techniques du Bâtiment'],
            "MEFIM": ['UE 1 – Economie & Gestion', 'Stratégie et Solutions Immobilières', 'Finance Immobilière', 'Economie Immobilière I', 'UE 2 – Droit', 'Droit des Affaires et des Contrats', 'UE 3 – Aménagement & Urbanisme', 'Ville et Développements Urbains', "Politique de l'Habitat", 'UE 4 – Compétences Professionnalisantes', 'Real Estate English', "Les Rencontres de l'Immobilier", 'ESPI Career Services', 'ESPI Inside', 'Immersion Professionnelle', 'Projet Voltaire', 'UE SPE – MEFIM', "Les Fondamentaux de l'Evaluation", 'Analyse et Financement Immobilier', 'Modélisation Financière'],
        }
        
        # Normalize titles for comparison
        normalized_titles_row = [normalize_string(title) for title in titles_row]

        # Identify the template based on the normalized titles_row
        template_key = None
        for key, values in matching_values.items():
            normalized_values = [normalize_string(value) for value in values]
            if normalized_titles_row == normalized_values:
                template_key = key
                break

        if template_key is None:
            raise HTTPException(status_code=400, detail="No matching template found")

        template_path = templates[template_key]
        logger.debug(f"Using template: {template_path}")

        bulletin_paths = []
        for index, student_data in df_students.iterrows():
            bulletin_path = generate_word_document(student_data, titles_row, template_path, output_dir)
            bulletin_paths.append(bulletin_path)
            logger.debug(f"Bulletin généré pour {student_data.get('Nom', 'N/A')}: {bulletin_path}")

        return bulletin_paths
    except Exception as e:
        logger.error("Erreur lors du traitement du fichier Excel", exc_info=True)
        raise HTTPException(status_code=400, detail=f"Error processing Excel file: {e}")
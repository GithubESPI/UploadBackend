import logging
from fastapi import APIRouter, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse
import os
import openpyxl
from app.core.config import settings
from app.services.api_service import fetch_api_data
from app.utils.date_utils import sum_durations, format_minutes_to_duration
from app.services.excel_service import process_excel_file

# Configure the logger
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

router = APIRouter()

def normalize_title(title):
    import re
    if not isinstance(title, str):
        title = str(title)
    return re.sub(r'\W+', '', title).lower()

async def fetch_api_data_for_template(headers):
    # Fetch API data for students
    api_url = f"{settings.YPAERO_BASE_URL}/r/v1/formation-longue/apprenants?codesPeriode=2"
    api_data = await fetch_api_data(api_url, headers)

    # Fetch API data for groups
    groupes_api_url = f"{settings.YPAERO_BASE_URL}/r/v1/formation-longue/groupes"
    groupes_data = await fetch_api_data(groupes_api_url, headers)

    # Fetch API data for absences
    absences_api_url = f"{settings.YPAERO_BASE_URL}/r/v1/absences/01-01-2023/31-12-2024"
    absences_data = await fetch_api_data(absences_api_url, headers)

    return api_data, groupes_data, absences_data

async def process_file(uploaded_wb, template_path, columns_config):
    template_wb = openpyxl.load_workbook(template_path, data_only=True)
    uploaded_ws = uploaded_wb.active
    template_ws = template_wb.active

    header_row_uploaded = 4
    header_row_template = 1

    uploaded_titles = {normalize_title(uploaded_ws.cell(row=header_row_uploaded, column=col).value): col 
    for col in range(1, uploaded_ws.max_column + 1) 
    if uploaded_ws.cell(row=header_row_uploaded, column=col).value is not None}

    template_titles = {normalize_title(template_ws.cell(row=header_row_template, column=col).value): col 
    for col in range(1, template_ws.max_column + 1) 
    if template_ws.cell(row=header_row_template, column=col).value is not None}

    matching_columns = {uploaded_title: (uploaded_titles[uploaded_title], template_titles[template_title]) 
                        for uploaded_title in uploaded_titles 
                        for template_title in template_titles 
                        if uploaded_title == template_title}

    if not matching_columns:
        return JSONResponse(content={"message": "No matching columns found, leaving new table empty."})

    # Set "Nom" in cell B2 of the template (corrected cell location)
    template_ws.cell(row=header_row_template + 1, column=columns_config['name_column_index_template']).value = "Nom"

    headers = {
        'X-Auth-Token': settings.YPAERO_API_TOKEN,
        'Content-Type': 'application/json'
    }

    api_data, groupes_data, absences_data = await fetch_api_data_for_template(headers)

    # Ensure api_data, groupes_data, and absences_data are dictionaries
    if not isinstance(api_data, dict) or not isinstance(groupes_data, dict) or not isinstance(absences_data, dict):
        raise HTTPException(status_code=500, detail="Unexpected API response format")

    # Debug logging
    logger.debug(f"API Data: {api_data}")
    logger.debug(f"Groupes Data: {groupes_data}")
    logger.debug(f"Absences Data: {absences_data}")

    # Create dictionaries from API data for quick lookup
    api_dict = {normalize_title(apprenant['nomApprenant'] + apprenant['prenomApprenant']): apprenant for key, apprenant in api_data.items()}
    groupes_dict = {groupe['codeGroupe']: groupe for groupe in groupes_data.values()}
    absences_summary = {}
    for absence in absences_data.values():
        apprenant_id = absence.get('codeApprenant')
        duration = int(absence.get('duree', 0))

        if apprenant_id not in absences_summary:
            absences_summary[apprenant_id] = {'justified': [], 'unjustified': [], 'delays': []}

        if absence.get('isJustifie'):
            absences_summary[apprenant_id]['justified'].append(duration)
        elif absence.get('isRetard'):
            absences_summary[apprenant_id]['delays'].append(duration)
        else:
            absences_summary[apprenant_id]['unjustified'].append(duration)

    logger.debug(f"API Dictionary: {api_dict}")
    logger.debug(f"Groupes Dictionary: {groupes_dict}")
    logger.debug(f"Absences Summary: {absences_summary}")

    exclude_phrase = 'moyennedugroupe'  # Normalized exclude phrase
    for row in range(header_row_uploaded + 1, uploaded_ws.max_row + 1):
        if any(exclude_phrase in normalize_title(uploaded_ws.cell(row=row, column=col).value or '') for col in range(1, uploaded_ws.max_column + 1)):
            continue  # Skip rows that contain 'Moyenne du groupe'

        # Copy names from uploaded file to template
        uploaded_name = uploaded_ws.cell(row=row, column=columns_config['name_column_index_uploaded']).value
        template_row = row - header_row_uploaded + header_row_template + 1
        template_ws.cell(row=template_row, column=columns_config['name_column_index_template']).value = uploaded_name

        # Normalize the name for comparison
        normalized_name = normalize_title(uploaded_name)

        # Debug: Check if normalized name is found in the API data
        if (apprenant_info := api_dict.get(normalized_name)):
            logger.debug(f"Found API data for: {uploaded_name}, {apprenant_info}")

            # Add API data to the template if columns exist
            template_ws.cell(row=template_row, column=columns_config['code_apprenant_column_index_template']).value = apprenant_info.get('codeApprenant', 'N/A')
            template_ws.cell(row=template_row, column=columns_config['date_naissance_column_index_template']).value = apprenant_info.get('dateNaissance', 'N/A')
            if 'inscriptions' in apprenant_info and apprenant_info['inscriptions']:
                template_ws.cell(row=template_row, column=columns_config['nom_site_column_index_template']).value = apprenant_info['inscriptions'][0]['site'].get('nomSite', 'N/A')
            
            code_groupe = apprenant_info.get('informationsCourantes', {}).get('codeGroupe', None)
            if code_groupe and code_groupe in groupes_dict:
                groupe_info = groupes_dict[code_groupe]
                template_ws.cell(row=template_row, column=columns_config['code_groupe_column_index_template']).value = groupe_info.get('codeGroupe', 'N/A')
                template_ws.cell(row=template_row, column=columns_config['nom_groupe_column_index_template']).value = groupe_info.get('nomGroupe', 'N/A')
                template_ws.cell(row=template_row, column=columns_config['etendu_groupe_column_index_template']).value = groupe_info.get('etenduGroupe', 'N/A')

            # Add absence data to the template if columns exist
            apprenant_id = apprenant_info.get('codeApprenant')
            abs_info = absences_summary.get(apprenant_id, {'justified': [], 'unjustified': [], 'delays': []})
            
            # Debugging the values before writing them to the sheet
            justified_duration = sum_durations(abs_info['justified']) or "00h00"
            unjustified_duration = sum_durations(abs_info['unjustified']) or "00h00"
            delays_duration = sum_durations(abs_info['delays']) or "00h00"
            logger.debug(f"Row {row} - Justified: {justified_duration}, Unjustified: {unjustified_duration}, Delays: {delays_duration}")

            template_ws.cell(row=template_row, column=columns_config['duree_justifie_column_index_template']).value = justified_duration
            template_ws.cell(row=template_row, column=columns_config['duree_non_justifie_column_index_template']).value = unjustified_duration
            template_ws.cell(row=template_row, column=columns_config['duree_retard_column_index_template']).value = delays_duration

        for uploaded_title, (src_col, dest_col) in matching_columns.items():
            src_cell = uploaded_ws.cell(row=row, column=src_col)
            dest_cell = template_ws.cell(row=template_row, column=dest_col)
            dest_cell.value = src_cell.value

    # Remove duplicates
    for col in range(1, template_ws.max_column + 1):
        if template_ws.cell(row=header_row_template + 1, column=col).value == template_ws.cell(row=header_row_template, column=col).value:
            template_ws.cell(row=header_row_template + 1, column=col).value = None

    # Remove duplicate "Note" in row 3
    for col in range(1, template_ws.max_column + 1):
        if template_ws.cell(row=header_row_template + 2, column=col).value == "Note":
            template_ws.cell(row=header_row_template + 2, column=col).value = None

    # After all processing and before saving, remove the target phrase
    target_phrase = "* Attention, le total des absences prend en compte toutes les absences aux séances sur la période concernée. S'il existe des absences sur des matières qui ne figurent pas dans le relevé, elles seront également comptabilisées."
    for row in template_ws.iter_rows():
        for cell in row:
            if cell.value == target_phrase:
                cell.value = None

    return template_wb

@router.post("/upload-and-integrate-excel")
async def upload_and_integrate_excel(file: UploadFile = File(...)):
    try:
        temp_file_path = os.path.join(settings.UPLOAD_DIR, file.filename)
        with open(temp_file_path, 'wb') as temp_file:
            temp_file.write(await file.read())

        uploaded_wb = openpyxl.load_workbook(temp_file_path, data_only=True)
        uploaded_ws = uploaded_wb.active

        # Extract specific cells values for template matching
        uploaded_values = []
        for cell in ['C4', 'F4', 'I4', 'L4', 'O4', 'R4', 'U4', 'X4', 'AA4', 'AD4', 'AG4', 'AJ4', 'AM4', 'AP4', 'AS4', 'AV4', 'AY4', 'BB4', 'BE4', 'BH4']:
            cell_value = uploaded_ws[cell].value
            if cell_value is not None:
                uploaded_values.append(cell_value)

        # Log the extracted values for debugging
        logger.debug(f"Extracted values from uploaded file: {uploaded_values}")

        # Define the template paths
        templates = {
            "MAPI": settings.M1_S1_MAPI_TEMPLATE,
            "MAGI": settings.M1_S1_MAGI_TEMPLATE,
            "MEFIM": settings.M1_S1_MEFIM_TEMPLATE,
            "MAPI_S2": settings.M1_S2_MAPI_TEMPLATE,
            "MAGI_S2": settings.M1_S2_MAGI_TEMPLATE,
            "MEFIM_S2": settings.M1_S2_MEFIM_TEMPLATE,
            "MAPI_S3": settings.M2_S3_MAPI_TEMPLATE,
            "MAGI_S3": settings.M2_S3_MAGI_TEMPLATE,
            "MEFIM_S3": settings.M2_S3_MEFIM_TEMPLATE,
            "MAPI_S4": settings.M2_S4_MAPI_TEMPLATE,
            "MAGI_S4": settings.M2_S4_MAGI_TEMPLATE,
            "MEFIM_S4": settings.M2_S4_MEFIM_TEMPLATE,
        }

        # Define the matching values for each template
        matching_values = {
            "MAPI": ['UE 1 – Economie & Gestion', 'Stratégie et Solutions Immobilières', 'Finance Immobilière', 'Economie Immobilière I', 'UE 2 – Droit', 'Droit des Affaires et des Contrats', 'UE 3 – Aménagement & Urbanisme', 'Ville et Développements Urbains', "Politique de l'Habitat", 'UE 4 – Compétences Professionnalisantes', 'Real Estate English', "Rencontres de l'Immobilier", 'ESPI Career Services', 'ESPI Inside', 'Immersion Professionnelle', 'Projet Voltaire', 'UE SPE – MAPI', 'Etude Foncière', "Montage d'une Opération de Promotion Immobilière", 'Acquisition et Dissociation du Foncier'],
            "MAGI": ['UE 1 – Economie & Gestion', 'Stratégie et Solutions Immobilières', 'Finance Immobilière', 'Economie Immobilière I', 'UE 2 – Droit', 'Droit des Affaires et des Contrats', 'UE 3 – Aménagement & Urbanisme', 'Ville et Développements Urbains', "Politique de l'Habitat", 'UE 4 – Compétences Professionnalisantes', 'Real Estate English', "Rencontres de l'Immobilier", 'ESPI Career Services', 'ESPI Inside', 'Immersion Professionnelle', 'Projet Voltaire', 'UE SPE – MAGI', 'Baux Commerciaux et Gestion Locative', 'Actifs Tertiaires en Copropriété', 'Techniques du Bâtiment'],
            "MEFIM": ['UE 1 – Economie & Gestion', 'Stratégie et Solutions Immobilières', 'Finance Immobilière', 'Economie Immobilière I', 'UE 2 – Droit', 'Droit des Affaires et des Contrats', 'UE 3 – Aménagement & Urbanisme', 'Ville et Développements Urbains', "Politique de l'Habitat", 'UE 4 – Compétences Professionnalisantes', 'Real Estate English', "Rencontres de l'Immobilier", 'ESPI Career Services', 'ESPI Inside', 'Immersion Professionnelle', 'Projet Voltaire', 'UE SPE – MEFIM', "Les Fondamentaux de l'Evaluation", 'Analyse et Financement Immobilier', 'Modélisation Financière'],
            
            "MAPI_S2": ['UE 1 – Economie & Gestion', "Marketing de l'Immobilier", 'Investissement et Financiarisation', 'Fiscalité', 'UE 2 – Droit', "Droit de l'Urbanisme et de la Construction", "Déontologie en France et à l'International", 'UE 4 – Compétences Professionnalisantes', 'Immersion Professionnelle', 'Real Estate English', 'Atelier Méthodologie de la Recherche', 'Techniques de Négociation', "Rencontres de l'Immobilier", 'ESPI Inside', 'Projet Voltaire', 'UE SPE – MAPI', "Droit de la Promotion Immobilière", "Montage d'une Opération de Logement", 'Financement des Opérations de Promotion Immobilière', "Logement Social et Accession Sociale"],
            "MAGI_S2": ['UE 1 – Economie & Gestion', "Marketing de l'Immobilier", 'Investissement et Financiarisation', 'Fiscalité', 'UE 2 – Droit', "Droit de l'Urbanisme et de la Construction", "Déontologie en France et à l'International", 'UE 4 – Compétences Professionnalisantes', 'Immersion Professionnelle', 'Real Estate English', 'Atelier Méthodologie de la Recherche', 'Techniques de Négociation', "Rencontres de l'Immobilier", 'ESPI Inside', 'Projet Voltaire', 'UE SPE – MAGI', "Budget d'Exploitation et de Travaux", 'Développement et Stratégie Commerciale', 'Technique et Conformité des Immeubles', "Gestion de l'Immobilier - Logistique et Data Center"],
            "MEFIM_S2": ['UE 1 – Economie & Gestion', "Marketing de l'Immobilier", 'Investissement et Financiarisation', 'Fiscalité', 'UE 2 – Droit', "Droit de l'Urbanisme et de la Construction", "Déontologie en France et à l'International", 'UE 4 – Compétences Professionnalisantes', 'Immersion Professionnelle', 'Real Estate English', 'Atelier Méthodologie de la Recherche', 'Techniques de Négociation', "Rencontres de l'Immobilier", 'ESPI Inside', 'Projet Voltaire', 'UE SPE – MEFIM', "Marché d'Actifs Immobiliers", "Baux Commerciaux", 'Evaluation des Actifs Résidentiels', "Audit et Gestion des Immeubles"],
            
            "MAPI_S3": ['UE 1 – Economie & Gestion', "PropTech et Innovation", 'Economie Immobilière II', 'UE 3 – Aménagement & Urbanisme', "Stratégies et Aménagement des Territoires I", "UE 4 – Compétences Professionnalisantes", 'Communication Digitale, Ecrite et Orale', 'Immersion Professionnelle', 'Real Estate English', 'Méthodologie de la Recherche', "Rencontres de l'Immobilier", 'ESPI Inside', 'UE SPE – MAPI', "Acquisition et Dissociation du Foncier", "Montage des Opérations Tertiaires", "Aménagement et Commande Publique", "Techniques du Bâtiment", "Réhabilitation et Pathologies du Bâtiment"],
            "MAGI_S3": ['UE 1 – Economie & Gestion', "PropTech et Innovation", 'Economie Immobilière II', 'UE 3 – Aménagement & Urbanisme', "Stratégies et Aménagement des Territoires I", "UE 4 – Compétences Professionnalisantes", 'Communication Digitale, Ecrite et Orale', 'Immersion Professionnelle', 'Real Estate English', 'Méthodologie de la Recherche', "Rencontres de l'Immobilier", 'ESPI Inside', 'UE SPE – MAGI', "Rénovation Energétique des Actifs Tertiaires", "Arbitrage, Optimisation et Valorisation des Actifs Tertiaires", 'Maintenance et Facility Management', "Réhabilitation et Pathologies du Bâtiment"],
            "MEFIM_S3": ['UE 1 – Economie & Gestion', "PropTech et Innovation", 'Economie Immobilière II', 'UE 3 – Aménagement & Urbanisme', "Stratégies et Aménagement des Territoires I", "UE 4 – Compétences Professionnalisantes", 'Communication Digitale, Ecrite et Orale', 'Immersion Professionnelle', 'Real Estate English', 'Méthodologie de la Recherche', "Rencontres de l'Immobilier", 'ESPI Inside', 'UE SPE – MEFIM', "Droit des Suretés et de la Transmission", 'Due Diligence', "Evaluation d'Actifs Tertiaires et Industriels", "Gestion de Patrimoine"],
            
            "MAPI_S4": ['UE 1 – Economie & Gestion', "Economie de l'Environnement", 'UE 3 – Aménagement & Urbanisme', "Normalisation, Labellisation", "Stratégies et Aménagement des Territoires II", 'UE 4 – Compétences Professionnalisantes', 'Real Estate English', 'Mémoire de Recherche', "Rencontres de l'Immobilier", 'ESPI Career Services', 'Immersion Professionnelle', 'UE SPE – MAPI', "Business Game Aménagement et Promotion Immobilière", "Fiscalité et Promotion Immobilière", "Contentieux de l'Urbanisme"],
            "MAGI_S4": ['UE 1 – Economie & Gestion', "Economie de l'Environnement", 'UE 3 – Aménagement & Urbanisme', "Normalisation, Labellisation", "Stratégies et Aménagement des Territoires II", 'UE 4 – Compétences Professionnalisantes', 'Real Estate English', 'Mémoire de Recherche', "Rencontres de l'Immobilier", 'ESPI Career Services', 'Immersion Professionnelle', 'UE SPE – MAGI', "Business Game Property Management", "Gestion des Centres Commerciaux", "Gestion de Contentieux et Recouvrement"],
            "MEFIM_S4": ['UE 1 – Economie & Gestion', "Economie de l'Environnement", 'UE 3 – Aménagement & Urbanisme', "Normalisation, Labellisation", "Stratégies et Aménagement des Territoires II", 'UE 4 – Compétences Professionnalisantes', 'Real Estate English', 'Mémoire de Recherche', "Rencontres de l'Immobilier", 'ESPI Career Services', 'Immersion Professionnelle', 'UE SPE – MEFIM', "Business Game Arbitrage et Stratégies d'Investissement", "Fiscalité du Patrimoine", "Fintech et Blockchain"],
        }

        # Define the column configurations for each template
        column_configs = {
            "MAPI": {
                'name_column_index_uploaded': 2,
                'name_column_index_template': 2,
                'code_apprenant_column_index_template': 1,
                'date_naissance_column_index_template': 23,
                'nom_site_column_index_template': 24,
                'code_groupe_column_index_template': 25,
                'nom_groupe_column_index_template': 26,
                'etendu_groupe_column_index_template': 27,
                'duree_justifie_column_index_template': 28,
                'duree_non_justifie_column_index_template': 29,
                'duree_retard_column_index_template': 30,
            },
            "MAGI": {
                'name_column_index_uploaded': 2,
                'name_column_index_template': 2,
                'code_apprenant_column_index_template': 1,
                'date_naissance_column_index_template': 23,
                'nom_site_column_index_template': 24,
                'code_groupe_column_index_template': 25,
                'nom_groupe_column_index_template': 26,
                'etendu_groupe_column_index_template': 27,
                'duree_justifie_column_index_template': 28,
                'duree_non_justifie_column_index_template': 29,
                'duree_retard_column_index_template': 30,
            },
            "MEFIM": {
                'name_column_index_uploaded': 2,
                'name_column_index_template': 2,
                'code_apprenant_column_index_template': 1,
                'date_naissance_column_index_template': 23,
                'nom_site_column_index_template': 24,
                'code_groupe_column_index_template': 25,
                'nom_groupe_column_index_template': 26,
                'etendu_groupe_column_index_template': 27,
                'duree_justifie_column_index_template': 28,
                'duree_non_justifie_column_index_template': 29,
                'duree_retard_column_index_template': 30,
            },
            "MAPI_S2": {
                'name_column_index_uploaded': 2,
                'name_column_index_template': 2,
                'code_apprenant_column_index_template': 1,
                'date_naissance_column_index_template': 23,
                'nom_site_column_index_template': 24,
                'code_groupe_column_index_template': 25,
                'nom_groupe_column_index_template': 26,
                'etendu_groupe_column_index_template': 27,
                'duree_justifie_column_index_template': 28,
                'duree_non_justifie_column_index_template': 29,
                'duree_retard_column_index_template': 30,
            },
            "MAGI_S2": {
                'name_column_index_uploaded': 2,
                'name_column_index_template': 2,
                'code_apprenant_column_index_template': 1,
                'date_naissance_column_index_template': 24,
                'nom_site_column_index_template': 25,
                'code_groupe_column_index_template': 26,
                'nom_groupe_column_index_template': 27,
                'etendu_groupe_column_index_template': 28,
                'duree_justifie_column_index_template': 29,
                'duree_non_justifie_column_index_template': 30,
                'duree_retard_column_index_template': 31,
            },
            "MEFIM_S2": {
                'name_column_index_uploaded': 2,
                'name_column_index_template': 2,
                'code_apprenant_column_index_template': 1,
                'date_naissance_column_index_template': 24,
                'nom_site_column_index_template': 25,
                'code_groupe_column_index_template': 26,
                'nom_groupe_column_index_template': 27,
                'etendu_groupe_column_index_template': 28,
                'duree_justifie_column_index_template': 29,
                'duree_non_justifie_column_index_template': 30,
                'duree_retard_column_index_template': 31,
            },
            "MAPI_S3": {
                'name_column_index_uploaded': 2,
                'name_column_index_template': 2,
                'code_apprenant_column_index_template': 1,
                'date_naissance_column_index_template': 21,
                'nom_site_column_index_template': 22,
                'code_groupe_column_index_template': 23,
                'nom_groupe_column_index_template': 24,
                'etendu_groupe_column_index_template': 25,
                'duree_justifie_column_index_template': 26,
                'duree_non_justifie_column_index_template': 27,
                'duree_retard_column_index_template': 28,
            },
            "MAGI_S3": {
                'name_column_index_uploaded': 2,
                'name_column_index_template': 2,
                'code_apprenant_column_index_template': 1,
                'date_naissance_column_index_template': 20,
                'nom_site_column_index_template': 21,
                'code_groupe_column_index_template': 22,
                'nom_groupe_column_index_template': 23,
                'etendu_groupe_column_index_template': 24,
                'duree_justifie_column_index_template': 25,
                'duree_non_justifie_column_index_template': 26,
                'duree_retard_column_index_template': 27,
            },
            "MEFIM_S3": {
                'name_column_index_uploaded': 2,
                'name_column_index_template': 2,
                'code_apprenant_column_index_template': 1,
                'date_naissance_column_index_template': 20,
                'nom_site_column_index_template': 21,
                'code_groupe_column_index_template': 22,
                'nom_groupe_column_index_template': 23,
                'etendu_groupe_column_index_template': 24,
                'duree_justifie_column_index_template': 25,
                'duree_non_justifie_column_index_template': 26,
                'duree_retard_column_index_template': 27,
            },
            "MAPI_S4": {
                'name_column_index_uploaded': 2,
                'name_column_index_template': 2,
                'code_apprenant_column_index_template': 1,
                'date_naissance_column_index_template': 18,
                'nom_site_column_index_template': 19,
                'code_groupe_column_index_template': 20,
                'nom_groupe_column_index_template': 21,
                'etendu_groupe_column_index_template': 22,
                'duree_justifie_column_index_template': 23,
                'duree_non_justifie_column_index_template': 24,
                'duree_retard_column_index_template': 25,
            },
            "MAGI_S4": {
                'name_column_index_uploaded': 2,
                'name_column_index_template': 2,
                'code_apprenant_column_index_template': 1,
                'date_naissance_column_index_template': 18,
                'nom_site_column_index_template': 19,
                'code_groupe_column_index_template': 20,
                'nom_groupe_column_index_template': 21,
                'etendu_groupe_column_index_template': 22,
                'duree_justifie_column_index_template': 23,
                'duree_non_justifie_column_index_template': 24,
                'duree_retard_column_index_template': 25,
            },
            "MEFIM_S4": {
                'name_column_index_uploaded': 2,
                'name_column_index_template': 2,
                'code_apprenant_column_index_template': 1,
                'date_naissance_column_index_template': 18,
                'nom_site_column_index_template': 19,
                'code_groupe_column_index_template': 20,
                'nom_groupe_column_index_template': 21,
                'etendu_groupe_column_index_template': 22,
                'duree_justifie_column_index_template': 23,
                'duree_non_justifie_column_index_template': 24,
                'duree_retard_column_index_template': 25,
            },
        }

        template_to_use = None
        columns_config = None
        for template, values in matching_values.items():
            if uploaded_values[:len(values)] == values:
                template_to_use = templates[template]
                columns_config = column_configs[template]
                break

        if not template_to_use or not columns_config:
            raise HTTPException(status_code=400, detail="No matching template found")

        template_wb = await process_file(uploaded_wb, template_to_use, columns_config)

        output_path = os.path.join(settings.OUTPUT_DIR, 'integrated_data.xlsx')
        template_wb.save(output_path)

    except Exception as e:
        logger.error("Failed to process the file", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))

    return JSONResponse(content={"message": "Integration réussie"})

@router.post("/generate-bulletins")
async def generate_bulletins(file: UploadFile = File(...)):
    try:
        temp_file_path = os.path.join(settings.UPLOAD_DIR, file.filename)
        with open(temp_file_path, 'wb') as temp_file:
            temp_file.write(await file.read())

        output_dir = os.path.join(settings.OUTPUT_DIR, 'bulletins')
        os.makedirs(output_dir, exist_ok=True)

        bulletin_paths = process_excel_file(temp_file_path, output_dir)

        return JSONResponse(content={"message": "Bulletins générés avec succès", "bulletins": bulletin_paths})

    except Exception as e:
        logger.error("Failed to generate bulletins", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))
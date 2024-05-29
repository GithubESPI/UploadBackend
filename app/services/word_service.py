import logging
import json
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
from app.core.config import settings
import os
import unicodedata

# Configure logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

def read_ects_config():
    with open(settings.ECTS_JSON_PATH, 'r') as file:
        data = json.load(file)
    return data

def normalize_string(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn').lower()

def extract_grades_and_coefficients(grade_str):
    grades_coefficients = []
    if not grade_str.strip():
        return grades_coefficients  # Return empty list if string is empty
    parts = grade_str.split(" - ")
    for part in parts:
        if "Absent au devoir" in part:
            continue
        try:
            if "(" in part:
                grade_part, coefficient_part = part.rsplit("(", 1)
                coefficient_part = coefficient_part.rstrip(")")
            else:
                grade_part = part
                coefficient_part = "1.0"
            grade = grade_part.replace(",", ".").strip()
            coefficient = coefficient_part.replace(",", ".").strip()
            
            # Replace 'CCHM' with 1
            if grade == 'CCHM':
                grade = '1'
            grades_coefficients.append((float(grade), float(coefficient)))
        except ValueError:
            # Skip values that cannot be converted to float or are not in the expected format
            continue
    return grades_coefficients

def calculate_weighted_average(notes, ects):
    if not notes or not ects:
        return 0.0
    total_grade = sum(note * ects for note, ects in zip(notes, ects))
    total_ects = sum(ects)
    return total_grade / total_ects if total_ects != 0 else 0

def generate_word_document(student_data, case_config, template_path, output_dir):
    ects_config = read_ects_config()
    current_date = datetime.now().strftime("%d/%m/%Y")
    group_name = student_data["Nom Groupe"]
    is_relevant_group = group_name in settings.RELEVANT_GROUPS
    logger.debug("Processing document for group: %s", group_name)

    placeholders = {
        "nomApprenant": student_data["Nom"],
        "etendugroupe": student_data["Étendu Groupe"],
        "dateNaissance": student_data["Date de Naissance"],
        "groupe": student_data["Nom Groupe"],
        "campus": student_data["Nom Site"],
        "justifiee": student_data["ABS justifiées"],
        "injustifiee": student_data["ABS injustifiées"],
        "retard": student_data["Retards"],
        "datedujour": current_date,
    }

    # Ajouter les placeholders pour les titres et matières dynamiquement
    for idx, title in enumerate(case_config["titles_row"]):
        placeholders[f"UE{idx//3 + 1}_Title" if idx % 3 == 0 else f"matiere{idx}"] = title

    total_ects = 0  # Initialize total ECTS

    for i, col_index in enumerate(case_config["grade_column_indices"], start=1):
        grade_str = str(student_data.iloc[col_index]).strip() if pd.notna(student_data.iloc[col_index]) else ""
        if grade_str and grade_str != 'Note':
            grades_coefficients = extract_grades_and_coefficients(grade_str)
            individual_average = calculate_weighted_average([g[0] for g in grades_coefficients], [g[1] for g in grades_coefficients])
            placeholders[f"note{i}"] = f"{individual_average:.2f}" if individual_average else ""
            if individual_average >= 8:
                ects_value = int(ects_config.get(f"ECTS{i}", 0))
                placeholders[f"ECTS{i}"] = ects_value
            else:
                placeholders[f"ECTS{i}"] = 0
        else:
            placeholders[f"note{i}"] = ""
            placeholders[f"ECTS{i}"] = 0

    for ue, indices in case_config["ects_sum_indices"].items():
        sum_values = sum(float(placeholders[f"note{index}"]) for index in indices if placeholders[f"note{index}"] != "")
        sum_ects = sum(placeholders[f"ECTS{index}"] for index in indices)
        count_valid_notes = sum(1 for index in indices if placeholders[f"note{index}"] != "")
        average_ue = round(sum_values / count_valid_notes, 2) if count_valid_notes > 0 else 0
        placeholders[f"moy{ue}"] = average_ue
        placeholders[f"ECTS{ue}"] = sum_ects
        total_ects += sum_ects

    placeholders["moyenneECTS"] = total_ects

    total_notes = sum(placeholders[f"moy{ue}"] * placeholders[f"ECTS{ue}"] for ue in case_config["ects_sum_indices"])
    total_ects = sum(placeholders[f"ECTS{ue}"] for ue in case_config["ects_sum_indices"])
    placeholders["moyenne"] = round(total_notes / total_ects, 2) if total_ects else 0

    logger.debug(f"Placeholders: {placeholders}")  # Log the placeholders to check their values

    doc = DocxTemplate(template_path)
    doc.render(placeholders)
    output_filename = f"{student_data['Nom']}_bulletin.docx"
    output_filepath = os.path.join(output_dir, output_filename)
    doc.save(output_filepath)
    return output_filepath

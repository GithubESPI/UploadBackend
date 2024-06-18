import logging
import json
from datetime import datetime
import pandas as pd
from docxtpl import DocxTemplate
from app.core.config import settings
import os
import unicodedata
import math
import docx

# Configure logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

def read_ects_config():
    with open(settings.ECTS_JSON_PATH, 'r') as file:
        data = json.load(file)
    return data

def normalize_string(s):
    if not isinstance(s, str):
        s = str(s)
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

    # Filter out notes and ects where ects is zero
    filtered_notes = [note for note, ect in zip(notes, ects) if ect != 0]
    filtered_ects = [ect for ect in ects if ect != 0]

    # If no valid notes remain after filtering, return 0.0
    if not filtered_notes or not filtered_ects:
        return 0.0

    total_grade = sum(note * ect for note, ect in zip(filtered_notes, filtered_ects))
    total_ects = sum(filtered_ects)
    
    return total_grade / total_ects if total_ects != 0 else 0

def generate_placeholders(titles_row, case_config, student_data, current_date, ects_data):
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
        "appreciations": student_data["Appreciations"]
    }

    if case_config["key"] == "M1_S1":
        placeholders.update({
            "UE1_Title": titles_row[0],
            "matiere1": titles_row[1],
            "matiere2": titles_row[2],
            "matiere3": titles_row[3],
            "UE2_Title": titles_row[4],
            "matiere4": titles_row[5],
            "matiere5": titles_row[6],
            "UE3_Title": titles_row[7],
            "matiere6": titles_row[8],
            "matiere7": titles_row[9],
            "matiere8": titles_row[10],
            "UE4_Title": titles_row[11],
            "matiere9": titles_row[12],
            "matiere10": titles_row[13],
            "matiere11": titles_row[14],
            "matiere12": titles_row[15],
            "UESPE_Title": titles_row[16],
            "matiere13": titles_row[17],
            "matiere14": titles_row[18],
            "matiere15": titles_row[19],
        })
    elif case_config["key"] == "M1_S2":
        placeholders.update({
            "UE1_Title": titles_row[0],
            "matiere1": titles_row[1],
            "matiere2": titles_row[2],
            "matiere3": titles_row[3],
            "UE2_Title": titles_row[4],
            "matiere4": titles_row[5],
            "matiere5": titles_row[6],
            "UE3_Title": titles_row[7],
            "matiere6": titles_row[8],
            "matiere7": titles_row[9],
            "matiere8": titles_row[10],
            "matiere9": titles_row[11],
            "matiere10": titles_row[12],
            "matiere11": titles_row[13],
            "matiere12": titles_row[14],
            "UESPE_Title": titles_row[15],
            "matiere13": titles_row[16],
            "matiere14": titles_row[17],
            "matiere15": titles_row[18],
            "matiere16": titles_row[19],
        })
    elif case_config["key"] == "M2_S3_MAGI":
        placeholders.update({
            "UE1_Title": titles_row[0],
            "matiere1": titles_row[1],
            "matiere2": titles_row[2],
            "UE2_Title": titles_row[3],
            "matiere3": titles_row[4],
            "UE3_Title": titles_row[5],
            "matiere4": titles_row[6],
            "matiere5": titles_row[7],
            "matiere6": titles_row[8],
            "matiere7": titles_row[9],
            "matiere8": titles_row[10],
            "matiere9": titles_row[11],
            "UESPE_Title": titles_row[12],
            "matiere10": titles_row[13],
            "matiere11": titles_row[14],
            "matiere12": titles_row[15],
            "matiere13": titles_row[16],
        })
    elif case_config["key"] == "M2_S3_MEFIM":
        placeholders.update({
            "UE1_Title": titles_row[0],
            "matiere1": titles_row[1],
            "matiere2": titles_row[2],
            "UE2_Title": titles_row[3],
            "matiere3": titles_row[4],
            "UE3_Title": titles_row[5],
            "matiere4": titles_row[6],
            "matiere5": titles_row[7],
            "matiere6": titles_row[8],
            "matiere7": titles_row[9],
            "matiere8": titles_row[10],
            "matiere9": titles_row[11],
            "UESPE_Title": titles_row[12],
            "matiere10": titles_row[13],
            "matiere11": titles_row[14],
            "matiere12": titles_row[15],
            "matiere13": titles_row[16],
        })
    elif case_config["key"] == "M2_S3_MAPI":
        placeholders.update({
            "UE1_Title": titles_row[0],
            "matiere1": titles_row[1],
            "matiere2": titles_row[2],
            "UE2_Title": titles_row[3],
            "matiere3": titles_row[4],
            "UE3_Title": titles_row[5],
            "matiere4": titles_row[6],
            "matiere5": titles_row[7],
            "matiere6": titles_row[8],
            "matiere7": titles_row[9],
            "matiere8": titles_row[10],
            "matiere9": titles_row[11],
            "UESPE_Title": titles_row[12],
            "matiere10": titles_row[13],
            "matiere11": titles_row[14],
            "matiere12": titles_row[15],
            "matiere13": titles_row[16],
            "matiere14": titles_row[17],
        })
    elif case_config["key"] == "M2_S4":
        placeholders.update({
            "UE1_Title": titles_row[0],
            "matiere1": titles_row[1],
            "UE2_Title": titles_row[2],
            "matiere2": titles_row[3],
            "matiere3": titles_row[4],
            "UE3_Title": titles_row[5],
            "matiere4": titles_row[6],
            "matiere5": titles_row[7],
            "matiere6": titles_row[8],
            "matiere7": titles_row[9],
            "matiere8": titles_row[10],
            "UESPE_Title": titles_row[11],
            "matiere9": titles_row[12],
            "matiere10": titles_row[13],
            "matiere11": titles_row[14],
        })

    # Add ECTS values to placeholders, hiding specified ones
    for i in range(1, 17):
        if i not in case_config["hidden_ects"]:
            placeholders[f"ECTS{i}"] = ects_data.get(f"ECTS{i}", 0)

    return placeholders

def generate_word_document(student_data, case_config, template_path, output_dir):
    ects_config = read_ects_config()
    current_date = datetime.now().strftime("%d/%m/%Y")
    group_name = student_data["Nom Groupe"]
    is_relevant_group = group_name in settings.RELEVANT_GROUPS
    logger.debug("Processing document for group: %s", group_name)

    # Correct case key if necessary
    corrected_key = case_config["key"].replace("_", "-")

    ects_data_key = corrected_key
    if corrected_key == "M2_S3_MAGI_MEFIM":
        if "MAGI" in student_data["Nom Groupe"]:
            ects_data_key = "M2-S3-MAGI"
        elif "MEFIM" in student_data["Nom Groupe"]:
            ects_data_key = "M2-S3-MEFIM"

    ects_data = ects_config.get(ects_data_key, [{}])[0]
    logger.debug(f"ECTS data for {corrected_key}: {ects_data}")

    placeholders = generate_placeholders(case_config["titles_row"], case_config, student_data, current_date, ects_data)

    total_ects = 0  # Initialize total ECTS

    for i, col_index in enumerate(case_config["grade_column_indices"], start=1):
        grade_str = str(student_data.iloc[col_index]).strip() if pd.notna(student_data.iloc[col_index]) else ""
        if grade_str and grade_str != 'Note':
            grades_coefficients = extract_grades_and_coefficients(grade_str)
            individual_average = calculate_weighted_average([g[0] for g in grades_coefficients], [g[1] for g in grades_coefficients])
            placeholders[f"note{i}"] = f"{individual_average:.2f}" if individual_average else ""
            if individual_average > 8 and i not in case_config["hidden_ects"]:
                ects_value = int(ects_data.get(f"ECTS{i}", 1))  # Use default coefficient 1 for hidden ECTS and convert to int
                placeholders[f"ECTS{i}"] = ects_value
            elif individual_average > 0:
                placeholders[f"ECTS{i}"] = 0
            else:
                placeholders[f"ECTS{i}"] = ""
        else:
            placeholders[f"note{i}"] = ""
            placeholders[f"ECTS{i}"] = ""

    for ue, indices in case_config["ects_sum_indices"].items():
        ue_sum = 0
        ue_ects = 0

        for index in indices:
            note = float(placeholders[f"note{index}"]) if placeholders[f"note{index}"] not in ["", None] else 0
            ects = int(placeholders[f"ECTS{index}"]) if placeholders[f"ECTS{index}"] not in ["", None] else 0

            if ects != 0:
                ue_sum += note * ects
                ue_ects += ects

        count_valid_notes = len([index for index in indices if placeholders[f"note{index}"] not in ["", None]])
        average_ue = math.ceil(ue_sum / ue_ects * 100) / 100 if ue_ects > 0 else 0
        placeholders[f"moy{ue}"] = f"{average_ue:.2f}" if average_ue else ""
        placeholders[f"ECTS{ue}"] = ue_ects if ue_ects else ""
        total_ects += ue_ects

    placeholders["moyenneECTS"] = total_ects

    # Calcul de la moyenne générale en fonction des moyennes des UE
    total_ue_notes = sum(
        float(placeholders[f"moy{ue}"]) * int(placeholders[f"ECTS{ue}"])
        for ue in case_config["ects_sum_indices"].keys()
        if placeholders[f"moy{ue}"] not in ["", None] and placeholders[f"ECTS{ue}"] not in ["", 0, None]
    )
    total_ue_ects = sum(
        int(placeholders[f"ECTS{ue}"])
        for ue in case_config["ects_sum_indices"].keys()
        if placeholders[f"ECTS{ue}"] not in ["", 0, None]
    )

    # Calcul de la moyenne générale arrondie au centième près
    placeholders["moyenne"] = f"{math.ceil(total_ue_notes / total_ue_ects * 100) / 100:.2f}" if total_ue_ects else 0

    # Remove placeholders for hidden ECTS from the final document
    for hidden_ects in case_config["hidden_ects"]:
        placeholders.pop(f"ECTS{hidden_ects}", None)

    logger.debug(f"Placeholders: {placeholders}")  # Log the placeholders to check their values

    doc = DocxTemplate(template_path)
    doc.render(placeholders)
    output_filename = f"{normalize_string(student_data['Nom'])}_bulletin.docx"
    output_filepath = os.path.join(output_dir, output_filename)
    doc.save(output_filepath)
    return output_filepath
import pandas as pd
from classes import Class, Subject
import config
from typing import List, Dict, Optional
import helper


def process_class_sheet(
        xls,
        sheet_name,
        all_subjects_dict: Dict[str, Dict[str, Subject]]
) -> Optional[Class]:
    print(f"\n# --- Configuration for Class: {sheet_name} ---")
    df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
    if len(df.columns) < 4:
        print(f"# Skipping sheet '{sheet_name}' - it does not have the expected format.")
        return None

    student_col_name = df.columns[1]
    quarter_col_name = df.columns[2]
    subject_col_names = df.columns[3:]
    df[student_col_name] = df[student_col_name].ffill()
    data_df = df.dropna(subset=[student_col_name, quarter_col_name]).copy()
    student_list = data_df[student_col_name].unique().tolist()
    if not student_list:
        return None

    print("\n# List of student names")
    print(f"student_names_{sheet_name.replace(' ', '_')} = {student_list}")
    subjects_grades_dict = {}
    for subject in subject_col_names:
        if 'Unnamed' in str(subject):
            continue

        normalized_subject = str(subject).strip().lower()
        grade_series = data_df[subject]
        grade_string = "".join(grade_series.apply(helper.clean_grade))
        subjects_grades_dict[normalized_subject] = grade_string

    clean_class = Class(sheet_name, student_list)

    clean_class.is_kz = any(sheet_name.endswith(c) for c in ('A', 'a', '8B', '8b'))

    subject_template_key = sheet_name
    if not clean_class.is_kz:
        subject_template_key = sheet_name[:-1]

    if subject_template_key in all_subjects_dict:
        clean_class.subjects = all_subjects_dict[subject_template_key].copy()
    else:
        print(f"# WARNING: No subject template found for key '{subject_template_key}'. Class '{sheet_name}' will have no subjects.")
        return clean_class

    print("\n# Dictionary of subjects and their grade strings")
    print(f"subjects_{sheet_name.replace(' ', '_')} = {{")
    for subject, grades in subjects_grades_dict.items():
        if subject in clean_class.subjects:
            print(f"    '{subject}':\n        \"{grades}\",")
            clean_class.subjects[subject].grades = grades
        else:
            print(f"# WARNING: Grade data found for subject '{subject}', but it is not in the subject template for this class.")
    print("}")
    return clean_class


def extract_grades_and_classes(
        all_subjects_dict: Dict[str, Dict[str, Subject]],
        filepath=config.grades_path
) -> Dict[str, Optional[Class]]:
    try:
        xls = pd.ExcelFile(filepath)
    except FileNotFoundError:
        print(f"Error: The file '{filepath}' was not found.")
        return {}
    all_classes = {}
    for sheet_name in xls.sheet_names:
        all_classes[sheet_name] = process_class_sheet(xls, sheet_name, all_subjects_dict)
    return all_classes

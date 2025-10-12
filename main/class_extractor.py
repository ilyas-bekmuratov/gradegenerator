import pandas as pd
from classes import Class
import config
from typing import Dict
import helper


def process_class_sheet(
        xls,
        sheet_name,
        all_classes_dict: Dict[str, Class],
        target_class: str = ""
):
    if sheet_name != target_class and target_class != "":
        return

    print(f"\n# --- Configuration for Class: {sheet_name} ---")

    df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
    if len(df.columns) < 4:
        print(f"# Skipping sheet '{sheet_name}' - it does not have the expected format.")
        return

    student_col_name = df.columns[1]
    quarter_col_name = df.columns[2]
    subject_col_names = df.columns[3:]
    df[student_col_name] = df[student_col_name].ffill()
    data_df = df.dropna(subset=[student_col_name, quarter_col_name]).copy()
    student_list = data_df[student_col_name].unique().tolist()
    if not student_list:
        return

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

    clean_class = all_classes_dict[sheet_name]
    clean_class.students = student_list
    clean_class.is_kz = any(sheet_name.endswith(c) for c in ('A', 'a', '8B', '8b'))

    print("\n# Dictionary of subjects and their grade strings")
    print(f"subjects_{sheet_name.replace(' ', '_')} = {{")
    for subject, grades in subjects_grades_dict.items():
        if subject in clean_class.subjects:
            clean_class.subjects[subject].has_exam = check_exam_grade(grades)
            if not clean_class.subjects[subject].has_exam:
                grades = remove_6th_and_7th_chars(grades)
            clean_class.subjects[subject].grades = grades
            print(f"    '{subject}':\n        \"{grades}\",")
        else:
            print(f"# WARNING: Grades found for subject '{subject}', but subject is missing from class.")
    print("}")
    return clean_class


def check_exam_grade(grades: str):
    if grades[6] != '0':
        print("has exam")
        return True
    else:
        return False


def remove_6th_and_7th_chars(input_string: str):
    """
    A more concise version using a list comprehension.
    """
    # Create a list of 5-character slices for every 7-character step
    parts = [input_string[i:i+5] for i in range(0, len(input_string), 7)]
    return "".join(parts)


def extract_grades_and_classes(
        all_classes_dict: Dict[str, Class],
        filepath=config.grades_path,
        class_name: str = ""
):
    try:
        xls = pd.ExcelFile(filepath)
    except FileNotFoundError:
        print(f"Error: The file '{filepath}' was not found.")
        return
    for sheet_name in xls.sheet_names:
        all_classes_dict[sheet_name] = process_class_sheet(xls, sheet_name, all_classes_dict, target_class=class_name)
    return

import pandas as pd
from classes import Class
import config
from typing import Dict
import helper
import re


def process_class_sheet(
        xls,
        sheet_name,
        all_classes_dict: Dict[str, Class],
        target_class: str = "",
):
    if sheet_name != target_class and target_class != "":
        return

    print(f"\n# --- Configuration for Class: {sheet_name} ---")

    df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
    if len(df.columns) < 4:
        print(f"# Skipping sheet '{sheet_name}' - it does not have the expected format.")
        return

    genders_col_name = df.columns[0]
    student_col_name = df.columns[1]
    subject_col_names = df.columns[3:]
    df[student_col_name] = df[student_col_name].ffill()
    data_df = df.dropna(subset=[student_col_name]).copy()

    unique_student_df = data_df.drop_duplicates(subset=[student_col_name])
    student_list = unique_student_df[student_col_name].tolist()
    if not student_list:
        return

    gender_list = unique_student_df[genders_col_name].notna().tolist()
    print("\n# List of student names")
    print(f"student_names_{sheet_name.replace(' ', '_')} = {student_list}")
    print(f"genders_{sheet_name.replace(' ', '_')} = {gender_list}")
    subjects_grades_dict = {}
    for subject in subject_col_names:
        if 'Unnamed' in str(subject):
            continue

        normalized_subject = str(subject).strip().lower()
        grade_series = data_df[subject]
        grade_string = "".join(grade_series.fillna('').apply(helper.clean_grade))
        subjects_grades_dict[normalized_subject] = grade_string

    if sheet_name not in all_classes_dict:
        print(f"# WARNING: Class '{sheet_name}' from grades file not found in timetable data. Skipping.")
        return

    clean_class = all_classes_dict[sheet_name]
    clean_class.students = student_list
    clean_class.genders = gender_list

    print("\n# Dictionary of subjects and their grade strings")
    print(f"subjects_{sheet_name.replace(' ', '_')} = {{")
    class_number_str = re.match(r'^\d+', clean_class.name).group(0)
    class_number = int(class_number_str)
    for subject, grades in subjects_grades_dict.items():
        if subject in clean_class.subjects:
            clean_class.subjects[subject].has_exam = check_exam_grade(grades, sheet_name) and class_number >= 5
            if not clean_class.subjects[subject].has_exam and class_number >= 5:
                grades = remove_6th_and_7th_chars(grades)
            clean_class.subjects[subject].grades = grades
            print(f"    '{subject}':\n        \"{grades}\",")
        else:
            print(f"# WARNING: Grades found for subject '{subject}', but subject is missing from class.")
    print("}")
    return clean_class


def check_exam_grade(grades: str, class_name):
    class_number_str = re.match(r'^\d+', class_name).group(0)
    class_number = int(class_number_str)
    for grade in grades:
        if grade == '1':
            return False
    if grades[6] != '0' and class_number >= 5:
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

import pandas as pd
from class_class import Class, Subject
import config
from typing import List, Dict, Optional
from pathlib import Path


def clean_grade(grade):
    """
    Cleans and standardizes a single grade value.
    - Converts numbers like 4.0 to '4'.
    - Converts empty cells or non-numeric text (e.g., 'зачет') to '0'.
    """
    if pd.isna(grade) or str(grade).strip() == '':
        return '0'
    try:
        return str(int(float(grade)))
    except (ValueError, TypeError):
        return '0'


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
        grade_series = data_df[subject]
        grade_string = "".join(grade_series.apply(clean_grade))
        subjects_grades_dict[subject] = grade_string

    clean_class = Class(sheet_name, student_list)
    clean_class.is_kz = any(sheet_name.endswith(c) for c in ('A', 'a'))

    # Determine which subject template to use
    subject_template_key = sheet_name
    if not clean_class.is_kz:
        # russian classes have no letter in templates (e.g., 8 instead of 8B)
        subject_template_key = sheet_name[:-1]

    if subject_template_key in all_subjects_dict:
        # Create a deep copy to avoid modifying the original template
        clean_class.subjects = {name: Subject(s.name, s.teacher, s.hours) for name, s in all_subjects_dict[subject_template_key].items()}
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


def extract_subjects(filepath=config.subjects_path) -> Dict[str, Dict[str, Subject]]:
    all_class_subjects = {}
    try:
        xls = pd.ExcelFile(filepath)
    except FileNotFoundError:
        print(f"Error: The file '{filepath}' was not found.")
        return all_class_subjects
    for sheet_name in xls.sheet_names:
        try:
            print(f"\n--- Processing Class Sheet: {sheet_name} ---")
            subjects_in_class = process_subject_sheet(xls, sheet_name)
            if subjects_in_class:
                all_class_subjects[sheet_name] = subjects_in_class
        except Exception as e:
            print(f"# ERROR: Could not process sheet '{sheet_name}'. Reason: {e}")
    return all_class_subjects


def process_subject_sheet(xls, sheet_name) -> Optional[Dict[str, Subject]]:
    df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
    if len(df.columns) < 3:
        print(f"# WARNING: Skipping '{sheet_name}'. It has fewer than 3 columns.")
        return None
    df.dropna(how='all', inplace=True)
    subjects_list: List[str] = df.iloc[:, 0].tolist()
    teachers_list = df.iloc[:, 1].tolist()
    hours_list = df.iloc[:, 2].tolist()
    subjects_in_class = {}

    for subject_name, teacher, hours in zip(subjects_list, teachers_list, hours_list):
        if pd.isna(subject_name):
            continue  # Skip empty rows
        try:
            num_hours = int(hours) if not pd.isna(hours) else 0
        except (ValueError, TypeError):
            num_hours = 0

        subject = Subject(
            name=str(subject_name),
            teacher=str(teacher),
            hours=num_hours
        )
        subjects_in_class[subject.name] = subject

    print(f"  -> Successfully created object for '{sheet_name}' with {len(subjects_in_class)} subjects.")
    return subjects_in_class


def extract_topics_and_hw(
        all_class_subjects_dict: Dict[str, Dict[str, Subject]],
        is_kaz
):
    folder_path_str = config.kaz_topics_path if is_kaz else config.rus_topics_path
    if not folder_path_str:
        return

    path = Path(folder_path_str)
    if not path.is_dir():
        print(f"Error: The folder '{folder_path_str}' was not found.")
        return

    for file_path in path.glob('*.xlsx'):

        class_name = file_path.name[:2]
        if is_kaz:
            class_name = class_name + "A"  # e.g., "10A"

        subjects_for_this_class = all_class_subjects_dict.get(class_name)
        if not subjects_for_this_class:
            print(f"# WARNING: Class '{class_name}' not found for topics file '{file_path.name}'. Skipping.")
            continue

        try:
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                add_topics_and_hw(xls, sheet_name, subjects_for_this_class)
            print(f"  -> Successfully processed {len(xls.sheet_names)} subjects for class '{class_name}'.")
        except Exception as e:
            print(f"# ERROR: Could not process file '{file_path.name}'. Reason: {e}")


def add_topics_and_hw(
        xls,
        sheet_name,
        subjects_in_class: Dict[str, Subject]
):
    df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
    if len(df.columns) < 2:
        print(f"# WARNING: Skipping subject sheet '{sheet_name}'. It has fewer than 2 columns.")
        return

    topics: List[str] = [str(item) for item in df.iloc[:, 0].dropna().tolist()]
    homework: List[str] = [str(item) for item in df.iloc[:, 1].dropna().tolist()]

    subject_obj = subjects_in_class.get(sheet_name)
    if subject_obj:
        subject_obj.topics = topics
        subject_obj.homework = homework
    else:
        print(f"# ERROR: Could not find a subject named '{sheet_name}' for this class.")


def extract_all_data():
    subjects_per_class = extract_subjects()
    extract_topics_and_hw(subjects_per_class, True)
    extract_topics_and_hw(subjects_per_class, False)
    classes = extract_grades_and_classes(subjects_per_class)
    return classes

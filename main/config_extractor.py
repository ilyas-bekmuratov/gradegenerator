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
    # Check for empty or NaN values first
    if pd.isna(grade) or str(grade).strip() == '':
        return '0'

    try:
        # Convert float (e.g., 4.0) to an integer, then to a string
        return str(int(float(grade)))
    except (ValueError, TypeError):
        # If conversion fails, it's non-numeric text like 'зачет'
        return '0'


def process_class_sheet(
        xls,
        sheet_name,
        all_subjects_dict: Dict[str, Dict[str, Subject]]
) -> Optional[Class]:
    print(f"\n# --- Configuration for Class: {sheet_name} ---")

    df = pd.read_excel(xls, sheet_name=sheet_name, header=0)

    # We expect at least 4 columns: Index, Name, Quarter, Subject1
    if len(df.columns) < 4:
        print(f"# Skipping sheet '{sheet_name}' - it does not have the expected format.")
        return

    student_col_name = df.columns[1]  # Column B: 'ФИО обучающегося'
    quarter_col_name = df.columns[2]  # Column C: 'Четверть'
    subject_col_names = df.columns[3:]  # Columns D onwards

    # --- 1. Extract Student Names ---
    # Forward-fill handles merged cells for student names
    df[student_col_name] = df[student_col_name].ffill()

    # Filter for rows that contain actual grade data
    # This removes headers and empty rows from the data section
    data_df = df.dropna(subset=[student_col_name, quarter_col_name]).copy()
    student_list = data_df[student_col_name].unique().tolist()

    if student_list.count() == 0:
        return

    print("\n# List of student names")
    # Printing with a placeholder variable name for easy copy-pasting
    print(f"student_names_{sheet_name.replace(' ', '_')} = {student_list}")

    # --- 2. Extract Subjects and Grades ---
    subjects_grades_dict = {}
    for subject in subject_col_names:
        if 'Unnamed' in str(subject):
            continue
        grade_series = data_df[subject]
        grade_string = "".join(grade_series.apply(clean_grade))
        subjects_grades_dict[subject] = grade_string

    clean_class = Class(
        sheet_name,
        student_list
    )

    clean_class.is_kz = sheet_name.endswith(('A', 'a'))
    if clean_class.is_kz:
        clean_class.subjects = all_subjects_dict[sheet_name].copy()
    else:
        clean_class.subjects = all_subjects_dict[sheet_name[:-1]].copy()  # russian classes have no letter in templates

    # --- 3. Print the Subjects Dictionary ---
    print("\n# Dictionary of subjects and their grade strings")
    print(f"subjects_{sheet_name.replace(' ', '_')} = {{")
    for subject, grades in subjects_grades_dict.items():
        print(f"    '{subject}':\n        \"{grades}\",")
        clean_class.subjects[subject].grades = grades
    print("}")

    return clean_class


def extract_grades_and_classes(
        all_subjects_dict: Dict[str, Dict[str, Subject]],
        filepath=config.grades_path
) -> Optional[Dict[str, Class]]:
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
    """
    Reads the subjects.xlsx file, creates a Class object for each sheet,
    and populates it with subjects, teachers, and hours.

    Args:
        filepath (str): The path to the subjects Excel file.

    Returns:
        dict: A dictionary where keys are class names (e.g., "8A") and
              values are the corresponding populated Class objects.
    """
    all_class_subjects = {}

    try:
        xls = pd.ExcelFile(filepath)
    except FileNotFoundError:
        print(f"Error: The file '{filepath}' was not found.")
        return all_class_subjects  # Return an empty dictionary

    # Process each sheet in the Excel file
    for sheet_name in xls.sheet_names:
        try:
            print(f"\n--- Processing Class Sheet: {sheet_name} ---")

            # Add the fully populated object to our dictionary
            all_class_subjects[sheet_name] = process_subject_sheet(xls, sheet_name)

        except Exception as e:
            print(f"# ERROR: Could not process sheet '{sheet_name}'. Reason: {e}")

    return all_class_subjects


def process_subject_sheet(xls, sheet_name) -> Optional[Dict[str, Subject]]:
    df = pd.read_excel(xls, sheet_name=sheet_name, header=0)

    # Check for required columns
    if len(df.columns) < 3:
        print(f"# WARNING: Skipping '{sheet_name}'. It has fewer than 3 columns.")
        return {}

    # Remove any rows that are completely empty
    df.dropna(how='all', inplace=True)

    # Extract data from the first three columns into lists
    subjects_list: List[str] = df.iloc[:, 0].tolist()
    teachers_list = df.iloc[:, 1].tolist()
    hours_list = df.iloc[:, 2].tolist()

    subjects_in_class = {}

    for subject_name, teacher, hours in subjects_list, teachers_list, hours_list:
        # Create a new Class instance with the extracted data
        subject = Subject(
            name=subject_name,
            teacher=teacher,
            hours=hours_list
        )
        subjects_in_class[subject_name] = subject

    print(f"  -> Successfully created object for '{sheet_name}' with {len(subjects_list)} subjects.")
    return subjects_in_class


def extract_topics_and_hw(
        all_class_subjects_dict: Dict[str, Dict[str, Subject]],
        is_kaz
):
    if is_kaz:
        folder_path = config.kaz_topics_path,
    else:
        folder_path = config.rus_topics_path

    path = Path(folder_path)
    if not path.is_dir():
        print(f"Error: The folder '{folder_path}' was not found.")
        return

    for file_path in path.glob('*.xlsx'):
        file_name = file_path.name  # e.g., "10 Grade Topics.xlsx"

        class_name = file_name[:2]
        if is_kaz:
            class_name = class_name + "A"  # e.g., "10A"

        subjects_for_this_class = all_class_subjects_dict.get(class_name)
        if not subjects_for_this_class:
            print(f"# WARNING: Class '{class_name}' not found in the main dictionary. Skipping file.")
            continue

        try:
            xls = pd.ExcelFile(file_path)
            # Each sheet in the file corresponds to a subject
            for sheet_name in xls.sheet_names:
                add_topics_and_hw(xls, sheet_name, subjects_for_this_class)
            print(f"  -> Successfully processed {len(xls.sheet_names)} subjects for class '{class_name}'.")
        except Exception as e:
            print(f"# ERROR: Could not process file '{file_name}'. Reason: {e}")

    return


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

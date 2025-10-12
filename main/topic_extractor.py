import pandas as pd
from classes import Class
import config
from typing import Dict
from pathlib import Path
import re


def extract_all_topics_and_hw(
        all_classes_dict: Dict[str, Class],
        class_name: str = ""
):
    extract_topics_and_hw(all_classes_dict, True, target_class=class_name)
    extract_topics_and_hw(all_classes_dict, False, target_class=class_name)


def extract_topics_and_hw(
        all_classes_dict: Dict[str, Class],
        is_kaz: bool,
        target_class: str = ""
):
    folder_path_str = config.kaz_topics_path if is_kaz else config.rus_topics_path
    if not folder_path_str:
        return

    path = Path(folder_path_str)
    if not path.is_dir():
        print(f"Error: The folder '{folder_path_str}' was not found.")
        return

    print(f"\n--- Extracting topics/homework from {folder_path_str} ---")
    for file_path in path.glob('*.xlsx'):
        try:
            filename_stem = file_path.stem  # "5 Алгебра"

            # --- 1. Extract subject and class number from filename ---
            match = re.match(r'^(\d+)\s+(.+)', filename_stem)
            if not match:
                print(f"# WARNING: Skipping topics file with unexpected name format: '{file_path.name}'")
                continue

            class_num_str, subject_from_filename = match.groups()
            normalized_subject_name = subject_from_filename.strip().lower()

            # --- 2. Find the correct class.subjects dictionary to add topics to ---
            # Find a class that starts with the number and matches the language context (Kaz/Rus)
            for class_name_key, class_object in all_classes_dict.items():
                if class_name_key.startswith(class_num_str):
                    if target_class != "" and not target_class.startswith(class_num_str):
                        continue
                    is_class_key_kaz = any(class_name_key.endswith(c) for c in ('A', 'a', '8B', '8b'))

                    if (is_kaz and is_class_key_kaz) or (not is_kaz and not is_class_key_kaz):
                        set_data_to_subject(
                            class_object.subjects,
                            file_path,
                            normalized_subject_name,
                            class_name_key)

        except Exception as e:
            print(f"# ERROR: Could not process file '{file_path.name}'. Reason: {e}")


def set_data_to_subject(
        subjects_for_this_class,
        file_path,
        normalized_subject_name,
        target_class_name
):
    if not subjects_for_this_class:
        print(f"# WARNING: Could not find a matching class for topics file '{file_path.name}'. Skipping.")
        return

    # --- 3. Find the subject object within that class ---
    subject_obj = subjects_for_this_class.get(normalized_subject_name)
    if not subject_obj:
        print(f"# WARNING: Subject '{normalized_subject_name}' from file not found for class '{target_class_name}'. Skipping.")
        return

    # --- 4. Aggregate topics and homework from ALL sheets in the file ---
    xls = pd.ExcelFile(file_path)
    all_topics = []
    all_homework = []
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
        if len(df.columns) < 1:
            continue  # Skip empty sheets

        all_topics.extend(df.iloc[:, 0].dropna().astype(str).tolist())
        if len(df.columns) > 1:
            all_homework.extend(df.iloc[:, 1].dropna().astype(str).tolist())

    subject_obj.topics = all_topics
    subject_obj.homework = all_homework
    print(f"  -> class '{target_class_name}':'{normalized_subject_name}': {len(all_topics)} topics and {len(all_homework)} homeworks.")

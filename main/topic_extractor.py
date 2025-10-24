import pandas as pd
from classes import Class
import config
from typing import Dict
from pathlib import Path
import re
import timetable_extractor
import helper


def extract_all_topics_and_hw(
        all_classes_dict: Dict[str, Class],
        class_name: str = "",
        is_dod=False
):
    extract_topics_and_hw(all_classes_dict, True, target_class=class_name, is_dod=is_dod)
    extract_topics_and_hw(all_classes_dict, False, target_class=class_name, is_dod=is_dod)


def extract_topics_and_hw(
        all_classes_dict: Dict[str, Class],
        is_kaz: bool,
        target_class: str = "",
        is_dod=False
):
    if not is_dod:
        if is_kaz:
            folder_path_index = 0
        else:
            folder_path_index = 1
    else:
        if is_kaz:
            folder_path_index = 2
        else:
            folder_path_index = 3

    folder_path_str = config.topic_paths[folder_path_index]
    if not folder_path_str:
        return

    path = Path(folder_path_str)
    if not path.is_dir():
        print(f"Error: The folder '{folder_path_str}' was not found.")
        return

    print(f"\n--- Extracting topics/homework from {folder_path_str} ---")
    for file_path in path.glob('*.xlsx'):
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
                        class_name_key,
                        is_dod)


def set_data_to_subject(
        subjects_for_this_class,
        file_path,
        normalized_subject_name,
        target_class_name,
        is_dod: bool = False
):
    if not subjects_for_this_class:
        print(f"# WARNING: Could not find a matching class for topics file '{file_path.name}'. Skip.")
        return

    # --- 3. Find the subject object within that class ---
    subject_obj = subjects_for_this_class.get(normalized_subject_name)
    if not subject_obj:
        print(f"# WARNING: Subject '{normalized_subject_name}' from file not found for class '{target_class_name}'. Skip.")
        return

    # --- 4. Aggregate topics and homework from ALL sheets in the file ---
    xls = pd.ExcelFile(file_path)
    all_topics = []
    all_homework = []

    start_row_index = 8 if is_dod else 4  # Excel row 5 is 0-indexed as 4

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

        if len(df.columns) < 4:
            print(f"  # WARNING: Sheet '{sheet_name}' in '{file_path.name}' has fewer than 4 columns. Skipping sheet.")
            continue

        if len(df) <= start_row_index:
            print(f"  # WARNING: Sheet '{sheet_name}' in '{file_path.name}' has no data from row 5 onwards. Skipping.")
            continue

        for index, row in df.iloc[start_row_index:].iterrows():
            if not is_dod:
                topic_val = row.iloc[1]  # Column B (index 1) = Topic
                hw_val = row.iloc[2]  # Column C (index 2) = Homework
                hours_val = row.iloc[3]  # Column D (index 3) = Hours
            else:
                topic_val = row.iloc[2]  # Column B (index 1) = Topic
                hw_val = row.iloc[3]  # Column C (index 2) = Homework
                hours_val = row.iloc[0]  # Column D (index 3) = Hours

            # If the topic cell is empty, we assume it's the end of the list
            if pd.isna(topic_val) or str(topic_val).strip() == "":
                continue

            # Clean the topic and homework values
            topic = str(topic_val).strip()
            # Handle empty homework cells gracefully
            homework = str(hw_val).strip() if not pd.isna(hw_val) else ""

            hours = 1
            try:
                parsed_hours = int(float(hours_val))
                if parsed_hours > 1:
                    hours = parsed_hours
            except (ValueError, TypeError):
                hours = 1  # Default to 1 if cell is empty, text, or invalid

            # Add the topic and homework 'hours' number of times
            for _ in range(hours):
                all_topics.append(topic)
                all_homework.append(homework)

    subject_obj.topics = all_topics
    subject_obj.homework = all_homework
    print(f"  -> class '{target_class_name}':'{normalized_subject_name}': {len(all_topics)} topics and {len(all_homework)} homeworks.")
    total = 0
    for q in range(1, 5):
        total += len(helper.get_days_this_quarter(subject_obj, q))
    if is_dod and normalized_subject_name in config.two_per_month:
        total = total //2
    print(f"  -> in total has {total} hours this year.")


def test():
    class_str = "10D"
    is_dod = True
    all_classes_dict = timetable_extractor.extract_class_subjects(class_name=class_str, is_dod=is_dod)
    extract_all_topics_and_hw(all_classes_dict, class_name=class_str, is_dod=is_dod)
    for subject_name, subject in all_classes_dict[class_str].subjects.items():
        print(f"subject \'{subject_name}\' has topics: {subject.topics}")


if __name__ == "__main__":
    test()

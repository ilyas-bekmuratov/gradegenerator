"""
Generates plausible midterm and final exam scores for a given final grade mark
by reverse-engineering the grading process.
### How to Use the Script
0.  **Install libraries:** If you don't have them, open your terminal or command prompt and run:
    `pip install pandas numpy odfpy openpyxl`
1.  **Run the script:** Execute the file from your terminal while inside the folder with these scripts:
     `python main.py`
2. You can change the input by modifying the `config.py` file.
"""

import pandas as pd
import os
import config
import grade_generator as gg
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import column_index_from_string
import class_extractor
import topic_extractor
import timetable_extractor
import re
from collections import defaultdict
import random
import helper
import writer
from typing import List, Dict
from classes import Class, Subject


def extract_all_data():
    all_classes_dict = timetable_extractor.extract_class_subjects()
    topic_extractor.extract_all_topics_and_hw(all_classes_dict)
    class_extractor.extract_grades_and_classes(all_classes_dict)
    return all_classes_dict


def main():
    all_days_in_year = timetable_extractor.extract_days()
    all_classes_dict = extract_all_data()

    # --- Group classes by parallel (grade level) ---
    grouped_classes = defaultdict(list)
    for class_name, class_obj in all_classes_dict.items():
        if class_obj is None:
            continue
        match = re.match(r'^\d+', class_name)
        if match:
            parallel = match.group(0)
            grouped_classes[parallel].append(class_obj)

    os.makedirs(config.output_dir, exist_ok=True)

    # --- Loop through each parallel group and create a separate file ---
    for parallel, classes_in_parallel in grouped_classes.items():
        output_filename = f"journal {parallel}.xlsx"
        filepath = os.path.join(config.output_dir, output_filename)
        print(f"\n{'='*20} PROCESSING PARALLEL {parallel} {'='*20}")

        workbook = None
        try:
            if os.path.exists(filepath):
                workbook = openpyxl.load_workbook(filepath)
                print(f"Successfully loaded existing report from '{filepath}'.")
            else:
                workbook = openpyxl.load_workbook(config.template_path)
                print(f"Creating new report for parallel {parallel} from template.")

        except FileNotFoundError:
            print(f"Error: Template file not found at '{config.template_path}'.")
            continue
        except Exception as e:
            print(f"An error occurred while loading the workbook for parallel {parallel}: {e}")
            continue

        for current_class in classes_in_parallel:
            process_class(workbook, current_class, all_days_in_year)

        try:
            print("\nCleaning up final workbook...")
            for sheet_name in list(workbook.sheetnames):
                if sheet_name != config.template_sheet_name:
                    workbook.remove(workbook[sheet_name])

            workbook.save(filepath)
            print(f"\nSuccessfully saved the complete report to '{filepath}'.")
        except Exception as e:
            print(f"\nAn error occurred while saving the file '{filepath}': {e}")


def process_class(workbook, current_class: Class, all_days_in_year: Dict[int, List[str]]):
    for subject_name, subject in current_class.subjects.items():
        print(f"\n--- Processing Subject: {subject_name} ({subject.hours}h/w) for class {current_class.name} ---")

        class_number_str = re.match(r'^\d+', current_class.name).group(0)
        class_number = int(class_number_str)

        split = 7 if class_number >= 5 else 5
        split_grades: list[list[int]] = helper.split_string_by_pattern(subject.grades, split)

        for i in range(4):
            quarter_num = i + 1
            quarter(workbook, current_class, quarter_num, subject, split_grades, all_days_in_year)


def quarter(
        workbook,
        current_class: Class,
        quarter_num: int,
        subject: Subject,
        split_grades: list[list[int]],
        all_days_in_year: Dict[int, List[str]]
):
    chrome_length = len(f"{current_class.name} -  - Q{quarter_num}")
    max_subject_len = 31 - chrome_length
    short_subject_name = subject.name[:max_subject_len] if len(subject.name) > max_subject_len else subject.name
    output_sheet_name = f"{current_class.name} - {short_subject_name} - Q{quarter_num}"

    quarter_grades = split_grades[quarter_num - 1]

    if not any(quarter_grades):
        print(f"  -> Skipping Quarter {quarter_num} (no grades).")
        return

    print(f"  -> Generating data for Quarter {quarter_num}'...")
    results = []

    num_midterms_for_df = 1 if subject.hours == 1 else config.num_midterms

    for grade in quarter_grades:
        if grade in [0, 1]:  # Handle blank and pass/fail
            pass_fail_text = ""
            if grade == 1:
                pass_fail_text = "есп" if current_class.is_kz else "зач"

            blank_data = {
                "Input Grade": pass_fail_text,
                "СОр Scores (Midterms)": [''] * num_midterms_for_df,
                "СОч Score (Final)": '', "Adjusted СОр %": '', "Actual СОч %": '',
                "Generated Total %": '', "Penalty/Bonus Applied": 0
            }
            results.append(blank_data)
        elif grade in config.grade_bands:
            generated_data = gg.generate_plausible_grades(grade, current_class, subject, quarter_num)
            results.append(generated_data)

    if not results:
        return

    df = pd.DataFrame(results)

    actual_midterm_cols = [f'СОр {j+1}' for j in range(num_midterms_for_df)]
    midterm_df = pd.DataFrame(df['СОр Scores (Midterms)'].tolist(), index=df.index)
    midterm_df.columns = actual_midterm_cols

    template_midterm_cols = [f'СОр {j+1}' for j in range(config.max_midterms)]
    midterm_df = midterm_df.reindex(columns=template_midterm_cols)

    df = pd.concat([midterm_df, df.drop(columns=['СОр Scores (Midterms)'])], axis=1)

    max_sop_weight = config.weights['sop']
    max_so4_weight = config.weights['so4']
    final_df = df.rename(columns={
        'СОч Score (Final)': 'Балл СО за четв.',
        'Adjusted СОр %': f'% СОр (макс. {max_sop_weight}%)',
        'Actual СОч %': f'% СОч (макс. {max_so4_weight}%)',
        'Generated Total %': 'Сумма %', 'Input Grade': 'Оценка за четверть'
    })
    column_order = (
            template_midterm_cols +
            ['Балл СО за четв.', f'% СОр (макс. {max_sop_weight}%)',
             f'% СОч (макс. {max_so4_weight}%)', 'Сумма %', 'Оценка за четверть']
    )
    final_df = final_df[column_order]

    if output_sheet_name in workbook.sheetnames:
        sheet = workbook[output_sheet_name]
        print(f"  -> Found existing sheet: '{output_sheet_name}'. Overwriting data.")
    else:
        if config.template_sheet_name not in workbook.sheetnames:
            print(f"  -> ERROR: Template sheet '{config.template_sheet_name}' not found. Skipping.")
            return
        template_sheet = workbook[config.template_sheet_name]
        sheet = workbook.copy_worksheet(template_sheet)
        sheet.title = output_sheet_name
        print(f"  -> Created sheet '{output_sheet_name}' from template '{config.template_sheet_name}'.")

    [student_start_row, student_start_col] = config.student_name_cell
    for idx, student_name in enumerate(current_class.students):
        sheet.cell(row=student_start_row + idx, column=student_start_col, value=student_name)

    [subject_teacher_cell_row, subject_teacher_cell_col] = config.subject_teacher_cell
    title = f"Наименование предмета: {subject.name.capitalize()} Преподователь: {subject.teacher}"
    sheet.cell(row=subject_teacher_cell_row, column=subject_teacher_cell_col, value=title)

    rows = dataframe_to_rows(final_df, index=False, header=False)

    for r_idx, row_data in enumerate(rows, config.start_row):
        for c_idx, value in enumerate(row_data, ):
            sheet.cell(row=r_idx, column=c_idx, value=value if not pd.isna(value) else None)
    print(f"  -> Wrote main grade data for {len(final_df)} students.")

    total_hours_this_quarter = helper.get_hours_this_quarter(subject, quarter_num, all_days_in_year)

    topics_start_col = column_index_from_string(config.topic_col)
    quarter_start_index = helper.get_quarter_start_index(subject, quarter_num, total_hours_this_quarter)

    # --- Topic and Homework Distribution Logic ---
    print("  -> Placing dates, topics, homework")
    quarter_topics = subject.topics[quarter_start_index:quarter_start_index+total_hours_this_quarter]
    quarter_hw = subject.homework[quarter_start_index:quarter_start_index+total_hours_this_quarter]
    quarter_dates = helper.get_days_this_quarter(subject, quarter_num, all_days_in_year)

    for idx, date in enumerate(quarter_dates):
        sheet.cell(row=config.start_row + idx, column=topics_start_col, value=date)

    for idx, topic in enumerate(quarter_topics):
        sheet.cell(row=config.start_row + idx, column=topics_start_col, value=topic)

    for idx, hw in enumerate(quarter_hw):
        sheet.cell(row=config.start_row + idx, column=topics_start_col+1, value=hw)

    # --- Daily Grade Generation Logic ---
    writer.extend_day_columns(sheet, total_hours_this_quarter)

    num_grades_to_place = int(total_hours_this_quarter * config.daily_grade_density)

    quarter_grades_start_col = column_index_from_string(config.quarter_grade_col)
    daily_grades_start_col = column_index_from_string(config.daily_grade_col)

    available_cols = list(range(daily_grades_start_col, quarter_grades_start_col + total_hours_this_quarter))

    for idx, row in df.iterrows():
        student_row = student_start_row + idx
        bonus = row['Penalty/Bonus Applied']

        if bonus == 0:
            continue  # Skip for blank/pass-fail students

        distribution = config.get_daily_grade_distribution(bonus)
        grades, weights = zip(*distribution.items())
        cols_to_fill = random.sample(available_cols, num_grades_to_place)

        for col in cols_to_fill:
            generated_grade = random.choices(grades, weights=weights, k=1)[0]
            sheet.cell(row=student_row, column=col, value=generated_grade)


if __name__ == "__main__":
    main()

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
from openpyxl.utils import column_index_from_string
from class_class import Class
import config_extractor
import re
from collections import defaultdict
import random


def main():
    all_classes = config_extractor.extract_all_data()
    # --- Group classes by parallel (grade level) ---
    grouped_classes = defaultdict(list)
    for class_name, class_obj in all_classes.items():
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
            process_class(workbook, current_class)

        try:
            print("\nCleaning up final workbook...")
            for sheet_name in list(workbook.sheetnames):
                if sheet_name in config.template_sheet_names:
                    workbook.remove(workbook[sheet_name])

            workbook.save(filepath)
            print(f"\nSuccessfully saved the complete report to '{filepath}'.")
        except Exception as e:
            print(f"\nAn error occurred while saving the file '{filepath}': {e}")


def process_class(workbook, current_class: Class):
    for subject_name, subject in current_class.subjects.items():
        print(f"\n--- Processing Subject: {subject_name} ({subject.hours}h/w) for class {current_class.name} ---")

        class_number_str = re.match(r'^\d+', current_class.name).group(0)
        class_number = int(class_number_str)

        split = 7 if class_number >= 5 else 5
        split_grades = split_string_by_pattern(subject.grades, split)

        for i in range(4):
            quarter_num = i + 1
            quarter(workbook, current_class, quarter_num, subject_name, subject, split_grades)


def split_string_by_pattern(data_string: str, grades_per_student=7) -> list[list[int]]:
    # Splits a string of grades into 7 lists for (Q1, Q2, Q3, Q4, Final, exam, total).
    result_lists = [[] for _ in range(grades_per_student)]
    for index, char in enumerate(data_string):
        result_lists[index % grades_per_student].append(int(char))
    return result_lists


def quarter(workbook, current_class: Class, quarter_num, subject_name, subject, split_grades):
    class_name = current_class.name
    chrome_length = len(f"{class_name} -  - Q{quarter_num}")
    max_subject_len = 31 - chrome_length
    short_subject_name = subject_name[:max_subject_len] if len(subject_name) > max_subject_len else subject_name
    output_sheet_name = f"{class_name} - {short_subject_name} - Q{quarter_num}"

    quarter_grades = split_grades[quarter_num-1]

    subject_hours = subject.hours
    template_info = config.TEMPLATE_MAPPINGS.get(subject_hours, {}).get(quarter_num)

    if not template_info:
        print(f"  -> Skipping Q{quarter_num}: No template mapping for {subject_hours}-hour subject.")
        return

    template_sheet_name, start_col_letter = template_info
    start_col = column_index_from_string(start_col_letter)

    if not any(quarter_grades):
        print(f"  -> Skipping Quarter {quarter_num} (no grades).")
        return

    print(f"  -> Generating data for Quarter {quarter_num} using template '{template_sheet_name}'...")
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
        if template_sheet_name not in workbook.sheetnames:
            print(f"  -> ERROR: Template sheet '{template_sheet_name}' not found. Skipping.")
            return
        template_sheet = workbook[template_sheet_name]
        sheet = workbook.copy_worksheet(template_sheet)
        sheet.title = output_sheet_name
        print(f"  -> Created sheet '{output_sheet_name}' from template '{template_sheet_name}'.")

    [student_start_row, student_start_col] = config.student_name_cell
    for idx, student_name in enumerate(current_class.students):
        sheet.cell(row=student_start_row + idx, column=student_start_col, value=student_name)
    print(f"  -> Wrote {len(current_class.students)} student names to the sheet.")

    [subject_name_row, subject_name_col] = config.subject_name_cell
    sheet.cell(row=subject_name_row, column=subject_name_col, value=f"Наименование предмета: {subject_name.capitalize()}")

    # --- Daily Grade Generation Logic ---
    print("  -> Generating and placing daily grades...")
    available_cols = list(range(config.daily_grades_start_col, start_col))

    if available_cols:
        num_grades_to_place = int(len(available_cols) * config.daily_grade_density)
        for idx, row in df.iterrows():
            student_row = student_start_row + idx
            bonus = row['Penalty/Bonus Applied']

            if bonus == 0:
                continue  # Skip for blank/pass-fail students

            distribution = config.get_daily_grade_distribution(bonus)
            grades, weights = zip(*distribution.items())

            cols_to_fill = random.sample(available_cols, num_grades_to_place)

            for col in cols_to_fill:
                # Choose a grade based on the weighted distribution
                generated_grade = random.choices(grades, weights=weights, k=1)[0]
                sheet.cell(row=student_row, column=col, value=generated_grade)

        # --- Topic and Homework Distribution Logic ---
        print("  -> Placing topics and homework...")
        topics_start = start_col + config.topics_start_after
        quarter_topics = subject.topics[topics_start:topics_start+len(available_cols)]
        quarter_hw = subject.homework[topics_start:topics_start+len(available_cols)]

        for idx, topic in enumerate(quarter_topics):
            sheet.cell(row=config.start_row + idx, column=topics_start, value=topic)

        for idx, hw in enumerate(quarter_hw):
            sheet.cell(row=config.start_row + idx, column=topics_start+1, value=hw)


if __name__ == "__main__":
    main()

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
from class_class import Class, Subject
import config_extractor


def main():
    all_classes = config_extractor.extract_all_data()

    output_filename = os.path.splitext(config.output_filename)[0] + ".xlsx"
    filepath = os.path.join(config.output_dir, output_filename)

    os.makedirs(config.output_dir, exist_ok=True)

    workbook = None
    template_workbook = None
    try:
        # We always need the template to copy sheets from it
        template_workbook = openpyxl.load_workbook(config.template_path)

        if os.path.exists(filepath):
            workbook = openpyxl.load_workbook(filepath)
            print(f"Successfully loaded existing report from '{filepath}'.")
        else:
            # If the output file doesn't exist, we start with a fresh copy of the template
            workbook = openpyxl.load_workbook(config.template_path)
            print(f"Creating new report from template '{config.template_path}'.")

    except FileNotFoundError:
        print(f"Error: Template file not found at '{config.template_path}'.")
        print("Please ensure your template file exists.")
        return
    except Exception as e:
        print(f"An error occurred while loading the workbook: {e}")
        return

    # Filter out any non-class objects that might have been returned on error
    valid_classes = [c for c in all_classes.values() if c is not None]

    for current_class in valid_classes:
        process_class(workbook, template_workbook, current_class)

    try:
        workbook.save(filepath)
        print(f"\nSuccessfully saved the complete report to '{filepath}'.")
    except Exception as e:
        print(f"\nAn error occurred while saving the file: {e}")


def split_string_by_pattern(data_string: str, grades_per_student=7) -> list[list[int]]:
    # Splits a string of grades into 7 lists for (Q1, Q2, Q3, Q4, Final, exam, total).
    result_lists = [[] for _ in range(grades_per_student)]
    for index, char in enumerate(data_string):
        result_lists[index % grades_per_student].append(int(char))
    return result_lists


def process_class(workbook, template_workbook, current_class: Class):
    for subject_name, subject in current_class.subjects.items():
        print(f"\n--- Processing Subject: {subject_name} ({subject.hours}h/w) ---")

        split = 5
        if int(current_class.name[0:2]) > 4:
            split = 7
        split_grades = split_string_by_pattern(subject.grades, split)

        for i in range(4):
            quarter_num = i + 1
            output_sheet_name = f"{subject_name} - Q{quarter_num}"
            quarter_grades = split_grades[i]

            # 1. Look up settings from the mapping
            subject_hours = subject.hours
            template_info = config.TEMPLATE_MAPPINGS.get(subject_hours, {}).get(quarter_num)

            if not template_info:
                print(f"  -> Skipping Q{quarter_num}: No template mapping found for a {subject_hours}-hour subject.")
                continue

            template_sheet_name, start_col_letter = template_info

            # 2. Set the start position for writing data
            start_row = 7  # Row is always 7
            start_col = column_index_from_string(start_col_letter)

            if not any(quarter_grades):
                print(f"  -> Skipping Quarter {quarter_num} (no grades).")
                continue

            print(f"  -> Generating data for Quarter {quarter_num} using template '{template_sheet_name}'...")
            results = []
            for grade in quarter_grades:
                if grade == 0:
                    blank_data = {
                        "Input Grade": '', "СОр Scores (Midterms)": [''] * config.num_midterms,
                        "СОч Score (Final)": '', "Adjusted СОр %": '', "Actual СОч %": '',
                        "Generated Total %": '',
                    }
                    results.append(blank_data)
                elif grade in config.grade_bands:
                    # Pass the class and subject objects to the generator
                    generated_data = gg.generate_plausible_grades(grade, current_class, subject)
                    results.append(generated_data)

            if not results:
                continue

            # --- OUTPUT Formatting (DataFrame preparation)
            df = pd.DataFrame(results)

            actual_midterm_cols = [f'СОр {j+1}' for j in range(config.num_midterms)]
            midterm_df = pd.DataFrame(df['СОр Scores (Midterms)'].tolist(), columns=actual_midterm_cols, index=df.index)
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

            # --- Sheet creation and data writing ---
            if output_sheet_name in workbook.sheetnames:
                sheet = workbook[output_sheet_name]
                print(f"  -> Found existing sheet: '{output_sheet_name}'. Overwriting its data.")
            else:
                if template_sheet_name not in template_workbook.sheetnames:
                    print(f"  -> ERROR: Template sheet '{template_sheet_name}' not found in template file. Skipping.")
                    continue
                template_sheet = template_workbook[template_sheet_name]
                sheet = workbook.copy_worksheet(template_sheet)
                sheet.title = output_sheet_name
                print(f"  -> Created sheet '{output_sheet_name}' by copying template '{template_sheet_name}'.")

            # Write the subject name to its designated cell
            [subject_name_row, subject_name_col] = config.subject_name_cell
            string_to_enter = f"Наименование предмета: {subject_name}"
            sheet.cell(row=subject_name_row, column=subject_name_col, value=string_to_enter)

            # Write the grade data starting at its designated cell
            rows = dataframe_to_rows(final_df, index=False, header=False)

            for r_idx, row in enumerate(rows, start_row):
                for c_idx, value in enumerate(row, start_col):
                    if pd.isna(value):
                        value = None
                    sheet.cell(row=r_idx, column=c_idx, value=value)

            start_cell_addr = sheet.cell(row=start_row, column=start_col).coordinate
            print(f"  -> Data written to sheet '{output_sheet_name}' starting at cell {start_cell_addr}.")


if __name__ == "__main__":
    main()

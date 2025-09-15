"""
Generates plausible midterm and final exam scores for a given final grade mark
by reverse-engineering the grading process.
### How to Use the Script
0.  **Install libraries:** If you don't have them, open your terminal or command prompt and run:
    `pip install pandas numpy odfpy`
1.  **Run the script:** Execute the file from your terminal while inside the folder with these scripts:
     `python main.py`
2. You can change the input by modifying the `config file` within the main.py.
"""

import pandas as pd
import os
import config
import grade_generator as gg
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows


def main():
    """
    Loads an .xlsx template, iterates through subjects defined in the config,
    generates plausible grades, and populates the template with the data,
    preserving formatting.
    """
    settings = config.settings

    # --- Focus ONLY on the 'subjects' dictionary as requested ---
    subjects_data = settings['subjects']

    # --- File I/O setup ---
    output_dir = settings['output_dir']
    template_path = "reports/template.xlsx"
    output_filename = os.path.splitext(settings['output_filename'])[0] + ".xlsx"
    filepath = os.path.join(output_dir, output_filename)

    os.makedirs(output_dir, exist_ok=True)

    # --- Load the template workbook ---
    try:
        workbook = openpyxl.load_workbook(template_path)
        print(f"Successfully loaded template from '{template_path}'.")
    except FileNotFoundError:
        print(f"Error: Template file not found at '{template_path}'.")
        print("Please ensure you have saved your template as an .xlsx file in the 'reports' folder.")
        return
    except Exception as e:
        print(f"An error occurred while loading the template: {e}")
        return

    # NEW: Assume the first sheet is the master template to be copied ###
    template_sheet = workbook.worksheets[0]
    print(f"Using '{template_sheet.title}' as the master template sheet.")

    # --- Main Processing Loop ---
    for subject_name, grades_string in subjects_data.items():
        print(f"\n--- Processing Subject: {subject_name} ---")
        split_grades = split_string_by_pattern(grades_string)

        for i in range(4):
            quarter_num = i + 1
            sheet_name = f"{subject_name} - Q{quarter_num}"
            quarter_grades = split_grades[i]

            if not any(quarter_grades):
                print(f"  -> Skipping Quarter {quarter_num} (no grades).")
                continue

            print(f"  -> Generating data for Quarter {quarter_num}...")
            results = []
            for grade in quarter_grades:
                if grade == 0:
                    blank_data = {
                        "Input Grade": '', "СОр Scores (Midterms)": [''] * settings['num_midterms'],
                        "СОч Score (Final)": '', "Adjusted СОр %": '', "Actual СОч %": '',
                        "Generated Total %": '',
                    }
                    results.append(blank_data)
                elif grade in settings['grade_bands']:
                    generated_data = gg.generate_plausible_grades(grade, config)
                    results.append(generated_data)

            if not results:
                continue

            # --- OUTPUT Formatting (DataFrame preparation is the same) ---
            df = pd.DataFrame(results)
            num_midterms = settings['num_midterms']
            midterm_cols = [f'СОр {j+1}' for j in range(num_midterms)]
            midterm_df = pd.DataFrame(df['СОр Scores (Midterms)'].tolist(), columns=midterm_cols, index=df.index)
            df = pd.concat([midterm_df, df.drop(columns=['СОр Scores (Midterms)'])], axis=1)

            max_sop_weight = settings['weights']['sop']
            max_so4_weight = settings['weights']['so4']
            final_df = df.rename(columns={
                'СОч Score (Final)': 'Балл СО за четв.',
                'Adjusted СОр %': f'% СОр (макс. {max_sop_weight}%)',
                'Actual СОч %': f'% СОч (макс. {max_so4_weight}%)',
                'Generated Total %': 'Сумма %', 'Input Grade': 'Оценка за четверть'
            })
            column_order = (
                    midterm_cols +
                    ['Балл СО за четв.', f'% СОр (макс. {max_sop_weight}%)',
                     f'% СОч (макс. {max_so4_weight}%)', 'Сумма %', 'Оценка за четверть']
            )
            # This is the DataFrame with just the student data
            final_df = final_df[column_order]

            # --- MODIFIED: Sheet creation and data writing ---

            # NEW: If sheet doesn't exist, copy it from the master template ###
            if sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                print(f"  -> Found existing sheet: '{sheet_name}'.")
            else:
                sheet = workbook.copy_worksheet(template_sheet)
                sheet.title = sheet_name
                print(f"  -> Created sheet '{sheet_name}' by copying template.")

            # NEW: We now write ONLY the student data (from final_df) without headers ###
            rows = dataframe_to_rows(final_df, index=False, header=False)

            # NEW: Define start position as cell AM7 ###
            start_row = 7
            start_col = 39  # Column 'AM'

            for r_idx, row in enumerate(rows, start_row):
                for c_idx, value in enumerate(row, start_col):
                    sheet.cell(row=r_idx, column=c_idx, value=value)

            print(f"  -> Data written to sheet '{sheet_name}' starting at cell AM7.")

    # --- Save the modified workbook to the output file ---
    try:
        # If the template sheet is no longer needed in the final output, you can remove it
        # workbook.remove(template_sheet)
        workbook.save(filepath)
        print(f"\nSuccessfully saved the complete report to '{filepath}'.")
    except Exception as e:
        print(f"\nAn error occurred while saving the file: {e}")


def split_string_by_pattern(data_string: str) -> list[list[int]]:
    """Splits a string of grades into 5 lists for (Q1, Q2, Q3, Q4, Final)."""
    result_lists = [[], [], [], [], []]
    for index, char in enumerate(data_string):
        result_lists[index % 5].append(int(char))
    return result_lists


def split_string_by_pattern_special(data_string: str) -> list[list[int]]:
    """Splits a string of grades into 7 lists for (Q1, Q2, Q3, Q4, Final)."""
    result_lists = [[], [], [], [], [], [], []]
    for index, char in enumerate(data_string):
        result_lists[index % 7].append(int(char))
    return result_lists


if __name__ == "__main__":
    main()

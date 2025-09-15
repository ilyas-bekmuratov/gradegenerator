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


def main():
    """
    Iterates through all subjects and quarters defined in the config,
    generates plausible grades, and saves them to a single spreadsheet file
    with one sheet per subject/quarter combination.
    """
    settings = config.settings
    subjects_data = settings['subjects']
    output_dir = settings['output_dir']
    filename = settings['output_filename']
    filepath = os.path.join(output_dir, filename)
    num_midterms = settings['num_midterms']

    # --- File I/O setup ---
    os.makedirs(output_dir, exist_ok=True)

    all_sheets_data = {}
    if os.path.exists(filepath):
        try:
            existing_file = pd.ExcelFile(filepath, engine='odf')
            all_sheets_data = {sheet: existing_file.parse(sheet) for sheet in existing_file.sheet_names}
            print(f"Loaded {len(all_sheets_data)} existing sheets from '{filepath}'.")
        except Exception as e:
            print(f"Could not read existing file '{filepath}'. A new file will be created. Error: {e}")

    # --- Main Processing Loop ---
    for subject_name, grades_string in subjects_data.items():
        print(f"\n--- Processing Subject: {subject_name} ---")
        split_grades = split_string_by_pattern(grades_string)

        # Inner loop for quarters (Q1 to Q4)
        for i in range(4):
            quarter_num = i + 1
            sheet_name = f"{subject_name} - Q{quarter_num}"
            quarter_grades = split_grades[i]

            # Skip quarters with no grades at all
            if not any(quarter_grades):
                print(f"  -> Skipping Quarter {quarter_num} (no grades).")
                continue

            print(f"  -> Generating data for Quarter {quarter_num}...")
            results = []
            for grade in quarter_grades:
                if grade == 0:
                    blank_data = {
                        "Input Grade": '', "СОр Scores (Midterms)": [''] * num_midterms,
                        "СОч Score (Final)": '', "Adjusted СОр %": '', "Actual СОч %": '',
                        "Generated Total %": '',
                    }
                    results.append(blank_data)
                elif grade in settings['grade_bands']:
                    generated_data = gg.generate_plausible_grades(grade, config)
                    results.append(generated_data)

            if not results:
                continue

            # --- OUTPUT Formatting ---
            df = pd.DataFrame(results)
            midterm_cols = [f'СОр {j+1}' for j in range(num_midterms)]
            midterm_df = pd.DataFrame(df['СОр Scores (Midterms)'].tolist(), columns=midterm_cols, index=df.index)
            df = pd.concat([midterm_df, df], axis=1)

            max_sop_weight = settings['weights']['sop']
            max_so4_weight = settings['weights']['so4']
            final_df = df.rename(columns={
                'СОч Score (Final)': 'Балл СО за четв.',
                'Adjusted СОр %': f'% СОр (макс. {max_sop_weight}%)',
                'Actual СОч %': f'% СОч (макс. {max_so4_weight}%)',
                'Generated Total %': 'Сумма %',
                'Input Grade': 'Оценка за четверть'
            })
            column_order = (
                    midterm_cols +
                    ['Балл СО за четв.', f'% СОр (макс. {max_sop_weight}%)',
                     f'% СОч (макс. {max_so4_weight}%)', 'Сумма %', 'Оценка за четверть']
            )
            final_df = final_df[column_order]

            max_scores_row = {col: '' for col in final_df.columns}
            for j, col in enumerate(midterm_cols):
                max_scores_row[col] = settings['max_scores'][j]
            max_scores_row['Балл СО за четв.'] = settings['max_scores'][-1]
            max_scores_df = pd.DataFrame([max_scores_row])

            output_df = pd.concat([max_scores_df, final_df], ignore_index=True)

            # Add the generated DataFrame to our dictionary of all sheets
            all_sheets_data[sheet_name] = output_df
            print(f"  -> Data prepared for sheet '{sheet_name}'.")

    # --- Save to File ---
    if not all_sheets_data:
        print("\nNo data was generated. The output file will not be created.")
        return

    print(f"\nWriting {len(all_sheets_data)} sheets to '{filepath}'...")
    with pd.ExcelWriter(filepath, engine='odf') as writer:
        for sheet, data in all_sheets_data.items():
            data.to_excel(writer, sheet_name=sheet, index=False)

    print(f"Successfully saved the complete report to '{filepath}'.")


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

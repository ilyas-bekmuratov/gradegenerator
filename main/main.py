import pandas as pd
import os
import config
import grade_generator as gg
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


def main():
    # --- INPUT: Provide your list of final grades here ---
    settings = config.settings
    final_grades_list = settings['final_grades']
    split_grades = split_string_by_pattern(final_grades_list)
    current_quarter = settings['current_quarter']
    # --- Processing ---
    results = []
    for grade in split_grades[current_quarter]:
        if grade in settings['grade_bands']:
            generated_data = gg.generate_plausible_grades(grade, config)
            results.append(generated_data)

    # --- OUTPUT Formatting ---
    df = pd.DataFrame(results)

    num_midterms = settings['num_midterms']
    midterm_cols = [f'СОр {i+1}' for i in range(num_midterms)]
    midterm_df = pd.DataFrame(df['СОр Scores (Midterms)'].tolist(), columns=midterm_cols, index=df.index)
    df = pd.concat([midterm_df, df], axis=1)

    # Recalculate the actual СОр % from the generated scores and new max scores
    total_max_midterm_score = sum(settings['max_scores'][:num_midterms])
    if total_max_midterm_score > 0:
        df['Calculated СОр %'] = (df[midterm_cols].sum(axis=1) / total_max_midterm_score) * settings['weights']['sop']
    else:
        df['Calculated СОр %'] = 0

    df['Calculated СОр %'] = df['Calculated СОр %'].round(1)

    # Prepare the final table for printing
    max_sop_weight = settings['weights']['sop']
    max_so4_weight = settings['weights']['so4']
    final_df = df.rename(columns={
        'СОч Score (Final)': 'Балл СО за четв.',
        'Penalty/Bonus Applied': 'Adjustment %',
        'Calculated СОр %': f'% СОр (макс. {max_sop_weight}%)',
        'Actual СОч %': f'% СОч (макс. {max_so4_weight}%)',
        'Generated Total %': 'Сумма %',
        'Input Grade': 'Оценка за четверть'
    })
    column_order = (
            midterm_cols +
            ['Балл СО за четв.', 'Adjusted СОр %', 'Adjustment %', f'% СОр (макс. {max_sop_weight}%)',
             f'% СОч (макс. {max_so4_weight}%)', 'Сумма %', 'Оценка за четверть']
    )
    final_df = final_df[column_order]

    # Create the "Максимальные баллы" (Maximum Scores) row
    max_scores_row = {col: '' for col in final_df.columns}
    for i, col in enumerate(midterm_cols):
        max_scores_row[col] = settings['max_scores'][i]
    max_scores_row['Балл СО за четв.'] = settings['max_scores'][-1]
    max_scores_df = pd.DataFrame([max_scores_row])

    # Combine the max scores row with the data for printing
    output_df = pd.concat([max_scores_df, final_df], ignore_index=True)
    print(output_df.to_string())

    # --- Save to Excel File ---
    output_dir = settings['output_dir']
    filename = settings['output_filename']
    sheet_name = settings['sheet_name']
    filepath = os.path.join(output_dir, filename)

    os.makedirs(output_dir, exist_ok=True)

    # ODF engine does not support append mode, so we read, modify, and write.
    sheets = {}
    if os.path.exists(filepath):
        # Read all existing sheets into a dictionary
        existing_file = pd.ExcelFile(filepath, engine='odf')
        sheets = {sheet: existing_file.parse(sheet) for sheet in existing_file.sheet_names}

    # Add or replace the new data in the dictionary
    sheets[sheet_name] = output_df

    # Write all sheets back to the file, overwriting it
    with pd.ExcelWriter(filepath, engine='odf') as writer:
        for sheet, data in sheets.items():
            data.to_excel(writer, sheet_name=sheet, index=False)

    print(f"Successfully saved the report to '{filepath}' in sheet '{sheet_name}'.")


def split_string_by_pattern(data_string: str) -> list[list[int]]:
    result_lists = [[], [], [], [], []]
    for index, char in enumerate(data_string):
        result_lists[index % 5].append(int(char))
    return result_lists


if __name__ == "__main__":
    main()

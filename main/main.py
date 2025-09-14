import pandas as pd
import os
import config
import grade_generator as gg
"""
Generates plausible midterm and final exam scores for a given final grade mark
by reverse-engineering the grading process.
### How to Use the Script
0.  **Install libraries:** If you don't have them, open your terminal or command prompt and run:
    `pip install pandas numpy`
1.  **Run the script:** Execute the file from your terminal:
     `python grade_generator.py`
2. You can change the input by modifying the `final_grades_list` within the main.py.
To use a `.csv` file, you can uncomment the lines for reading a CSV and specify your column name.
"""


def main():
    # --- INPUT: Provide your list of final grades here ---
    final_grades_list = config.settings['final_grades']

    # --- Processing ---
    results = []
    for grade in final_grades_list:
        if grade in config.settings['grade_bands']:
            generated_data = gg.generate_plausible_grades(grade, config)
            results.append(generated_data)

    # --- OUTPUT Formatting ---
    df = pd.DataFrame(results)

    num_midterms = config.settings['num_midterms']
    midterm_cols = [f'СОр {i+1}' for i in range(num_midterms)]
    midterm_df = pd.DataFrame(df['СОр Scores (Midterms)'].tolist(), columns=midterm_cols, index=df.index)
    df = pd.concat([midterm_df, df], axis=1)

    # Recalculate the actual СОр % from the generated scores and new max scores
    total_max_midterm_score = sum(config.settings['max_scores'][:num_midterms])
    if total_max_midterm_score > 0:
        df['Calculated СОр %'] = (df[midterm_cols].sum(axis=1) / total_max_midterm_score) * config.settings['weights']['sop']
    else:
        df['Calculated СОр %'] = 0

    df['Calculated СОр %'] = df['Calculated СОр %'].round(1)

    # Prepare the final table for printing
    max_sop_weight = config.settings['weights']['sop']
    max_so4_weight = config.settings['weights']['so4']
    final_df = df.rename(columns={
        'Penalty/Bonus Applied': 'Adjustment %',
        'СОч Score (Final)': 'Балл СО за четв.',
        'Calculated СОр %': f'% СОр (макс. {max_sop_weight}%)',
        'Actual СОч %': f'% СОч (макс. {max_so4_weight}%)',
        'Generated Total %': 'Сумма %',
        'Input Grade': 'Оценка за четверть'
    })

    column_order = (
            midterm_cols +
            ['Балл СО за четв.', 'Adjustment %', f'% СОр (макс. {max_sop_weight}%)',
             f'% СОч (макс. {max_so4_weight}%)', 'Сумма %', 'Оценка за четверть']
    )
    final_df = final_df[column_order]

    # Create the "Максимальные баллы" (Maximum Scores) row
    max_scores_row = {col: '' for col in final_df.columns}
    for i, col in enumerate(midterm_cols):
        max_scores_row[col] = config.settings['max_scores'][i]
    max_scores_row['Балл СО за четв.'] = config.settings['max_scores'][-1]
    max_scores_df = pd.DataFrame([max_scores_row])

    # Combine the max scores row with the data for printing
    output_df = pd.concat([max_scores_df, final_df], ignore_index=True)
    print(output_df.to_string())

    # --- Save to Excel File ---
    output_dir = config.settings['output_dir']
    filename = config.settings['output_filename']
    sheet_name = config.settings['sheet_name']
    filepath = os.path.join(output_dir, filename)

    # Create the directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

    # Handle file creation and appending separately to avoid the error
    if os.path.exists(filepath):
        # If file exists, open in append mode and replace sheet if it exists
        with pd.ExcelWriter(
                filepath, engine='openpyxl', mode='a', if_sheet_exists='replace'
        ) as writer:
            output_df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        # If file does not exist, create it in write mode
        with pd.ExcelWriter(filepath, engine='openpyxl', mode='w') as writer:
            output_df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"Successfully saved the report to '{filepath}' in sheet '{sheet_name}'.")


if __name__ == "__main__":
    main()

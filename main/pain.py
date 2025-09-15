import pandas as pd
import os
import config
import grade_generator as gg
import shutil  # <-- Import the shutil library for file copying

def main():
    """
    Copies a template file and writes generated grade data into the appropriate
    sheets, preserving the template's formatting.
    """
    settings = config.settings
    subjects_data = settings['subjects']  # <-- We will only focus on this part of the config
    output_dir = settings['output_dir']
    template_path = settings['template_path']
    output_filename = settings['output_filename']
    output_filepath = os.path.join(output_dir, output_filename)
    num_midterms = settings['num_midterms']

    # --- 1. File I/O setup: Copy the template to the output path ---
    os.makedirs(output_dir, exist_ok=True)
    try:
        shutil.copy(template_path, output_filepath)
        print(f"Successfully copied template to '{output_filepath}'.")
    except FileNotFoundError:
        print(f"ERROR: Template file not found at '{template_path}'. Please check the path in config.py.")
        return
    except Exception as e:
        print(f"ERROR: Could not copy template file. {e}")
        return

    # --- 2. Main Processing Loop ---
    # Open the copied file in append mode to write data sheet by sheet
    with pd.ExcelWriter(output_filepath, engine='odf', mode='a', if_sheet_exists='overlay') as writer:
        # Loop through only the main 'subjects', as requested
        for subject_name, grades_string in subjects_data.items():
            print(f"\n--- Processing Subject: {subject_name} ---")

            # This logic for splitting grades and handling quarters remains the same
            split_grades = split_string_by_pattern(grades_string)

            for i in range(4):
                quarter_num = i + 1
                # We will create a sheet name that combines subject and quarter
                sheet_name = f"{subject_name} - Q{quarter_num}"
                quarter_grades = split_grades[i]

                if not any(quarter_grades):
                    print(f"  -> Skipping Quarter {quarter_num} for '{subject_name}' (no grades).")
                    continue

                print(f"  -> Generating data for '{sheet_name}'...")
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

                # --- OUTPUT Formatting (This logic is unchanged) ---
                df = pd.DataFrame(results)
                midterm_cols = [f'СОр {j+1}' for j in range(num_midterms)]
                midterm_df = pd.DataFrame(df['СОр Scores (Midterms)'].tolist(), columns=midterm_cols, index=df.index)
                df = pd.concat([midterm_df, df.drop(columns=['СОр Scores (Midterms)'])], axis=1)

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

                # --- 3. Write DataFrame to the specific sheet and location ---
                # We write the DataFrame to the sheet matching the subject name.
                # startrow=1 and startcol=0 means we start writing at cell A2.
                # This leaves the first row (row 1) for your own headers.
                final_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1, startcol=0)

                # You might want to also write the max scores above the headers
                max_scores_row = {col: '' for col in final_df.columns}
                for j, col in enumerate(midterm_cols):
                    max_scores_row[col] = settings['max_scores'][j]
                max_scores_row['Балл СО за четв.'] = settings['max_scores'][-1]
                max_scores_df = pd.DataFrame([max_scores_row])
                max_scores_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=0, startcol=0)


                print(f"  -> Data successfully written to sheet '{sheet_name}'.")

    print(f"\n✅ All processing complete. Report saved to '{output_filepath}'.")


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

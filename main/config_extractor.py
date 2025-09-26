import pandas as pd


def clean_grade(grade):
    """
    Cleans and standardizes a single grade value.
    - Converts numbers like 4.0 to '4'.
    - Converts empty cells or non-numeric text (e.g., 'зачет') to '0'.
    """
    # Check for empty or NaN values first
    if pd.isna(grade) or str(grade).strip() == '':
        return '0'

    try:
        # Convert float (e.g., 4.0) to an integer, then to a string
        return str(int(float(grade)))
    except (ValueError, TypeError):
        # If conversion fails, it's non-numeric text like 'зачет'
        return '0'


def process_sheet(xls, sheet_name):
    print(f"\n# --- Configuration for Class: {sheet_name} ---")

    # Read the sheet. We assume the headers are on the 2nd row (index 1).
    df = pd.read_excel(xls, sheet_name=sheet_name, header=1)

    # --- Validate Sheet Structure ---
    # We expect at least 4 columns: Index, Name, Quarter, Subject1
    if len(df.columns) < 4:
        print(f"# Skipping sheet '{sheet_name}' - it does not have the expected format.")
        return

    student_col_name = df.columns[1]  # Column B: 'ФИО обучающегося'
    quarter_col_name = df.columns[2]  # Column C: 'Четверть'
    subject_col_names = df.columns[3:]  # Columns D onwards

    # --- 1. Extract Student Names ---
    # Forward-fill handles merged cells for student names
    df[student_col_name] = df[student_col_name].ffill()

    # Filter for rows that contain actual grade data
    # This removes headers and empty rows from the data section
    data_df = df.dropna(subset=[student_col_name, quarter_col_name]).copy()

    # Get the unique list of students in their original order
    student_list = data_df[student_col_name].unique().tolist()

    print("\n# List of student names")
    # Printing with a placeholder variable name for easy copy-pasting
    print(f"student_names_{sheet_name.replace(' ', '_')} = {student_list}")

    # --- 2. Extract Subjects and Grades ---
    subjects_grades_dict = {}
    for subject in subject_col_names:
        # Pandas sometimes reads empty columns as 'Unnamed: X', we skip those.
        if 'Unnamed' in str(subject):
            continue

        # Get the entire column of grades for the current subject
        grade_series = data_df[subject]

        # Apply the cleaning function to each grade and join them into a single string
        grade_string = "".join(grade_series.apply(clean_grade))

        subjects_grades_dict[subject] = grade_string

    # --- 3. Print the Subjects Dictionary ---
    print("\n# Dictionary of subjects and their grade strings")
    print(f"subjects_{sheet_name.replace(' ', '_')} = {{")
    # Format the output to be clean and easy to read
    for subject, grades in subjects_grades_dict.items():
        # Escape single quotes in subject names to avoid syntax errors
        # subject_escaped = subject.replace("'", "\\'")
        print(f"    '{subject}':\n        \"{grades}\",")
    print("}")


def extract_config_from_excel(filepath="reports/grades.xlsx"):
    """
    Reads an Excel file with student grades and prints a Python dictionary
    and list that can be used for the grade generator's config.

    Args:
        filepath (str): The path to the grades Excel file.
    """
    try:
        # Load the entire Excel file to access its sheets
        xls = pd.ExcelFile(filepath)
    except FileNotFoundError:
        print(f"Error: The file '{filepath}' was not found.")
        print("Please make sure the Excel file is in the same directory as this script.")
        return

    # Process each sheet in the Excel file
    for sheet_name in xls.sheet_names:
        process_sheet(xls, sheet_name)


if __name__ == "__main__":
    # You can specify a different filename here if needed
    # For example: extract_config_from_excel("my_other_grades.xlsx")
    extract_config_from_excel("reports/grades.xlsx")

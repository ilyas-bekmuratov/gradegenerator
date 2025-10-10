import pandas as pd
import config
from typing import List, Dict
import re
from classes import Class, Subject


def extract_days(filepath=config.days_path) -> Dict[int, List[str]]:
    """
    Opens an Excel file and processes each sheet to extract dates grouped by quarter.
    """
    days = {}

    try:
        xls = pd.ExcelFile(filepath)
    except FileNotFoundError:
        print(f"Error: The file '{filepath}' was not found.")
        return {}
    print(f"\n--- Processing Days ---")
    for sheet_name in xls.sheet_names:
        try:
            # Process each sheet and merge the results
            sheet_days = process_days_sheet(xls, sheet_name)
            for quarter, date_list in sheet_days.items():
                if quarter not in days:
                    days[quarter] = []
                days[quarter].extend(date_list)
        except Exception as e:
            print(f"# ERROR: Could not process sheet '{sheet_name}'. Reason: {e}")
    return days


def process_days_sheet(xls, sheet_name) -> Dict[int, List[str]]:
    """
    Scrapes a single sheet for dates, organizing them by quarter.

    Args:
        xls (pd.ExcelFile): The opened Excel file object.
        sheet_name (str): The name of the sheet to process.

    Returns:
        A dictionary where keys are quarter numbers (int) and values are lists of
        date strings for that quarter.
    """
    # Read the sheet, assuming the first row is the header
    df = pd.read_excel(xls, sheet_name, header=0)

    # Check if the sheet has the expected structure (at least 6 columns: Quarter + 5 weekdays)
    if len(df.columns) < 6:
        print(f"# WARNING: Skipping '{sheet_name}'. It has fewer than 6 columns.")
        return {}

    # Replace any NaN (empty) cells with empty strings for consistent processing
    df.fillna('', inplace=True)

    days_by_quarter = {}
    current_quarter = None

    # Iterate over each row of the DataFrame
    for index, row in df.iterrows():
        # Column A (index 0) contains the quarter number
        quarter_cell_value = str(row.iloc[0])

        # Use regex to find if the cell in Column A specifies a new quarter
        match = re.search(r'\d+', quarter_cell_value)
        if match:
            current_quarter = int(match.group(0))
            # Initialize an empty list for this quarter if it's the first time we've seen it
            if current_quarter not in days_by_quarter:
                days_by_quarter[current_quarter] = []

        # If we have identified a quarter, process the dates in that row
        if current_quarter is not None:
            # Columns B to F (indices 1 to 5) contain the dates for Monday to Friday
            week_dates = row.iloc[1:6].tolist()

            # Format and clean the date values before adding them
            formatted_dates = []
            for date_val in week_dates:
                # Keep empty values as empty strings
                if not date_val:
                    formatted_dates.append('')
                    continue
                try:
                    # Use pandas to_datetime to robustly parse dates (DD.MM.YY)
                    # and format them into a consistent string format (DD.MM.YYYY)
                    dt_obj = pd.to_datetime(date_val, dayfirst=True)
                    formatted_dates.append(dt_obj.strftime('%d.%m.%Y'))
                except (ValueError, TypeError):
                    # If conversion fails, it's not a date; keep the original value as a string
                    formatted_dates.append(str(date_val))

            # Add the week's dates to the list for the current quarter
            days_by_quarter[current_quarter].extend(formatted_dates)

    print(f"Successfully processed sheet '{sheet_name}', found data for quarters: {list(days_by_quarter.keys())}")
    return days_by_quarter


def extract_class_subjects(filepath=config.timetable_path) -> Dict[str, Class]:
    """
    Opens the timetable Excel file and extracts the schedule for all classes from each sheet.
    """
    all_classes_data = {}
    try:
        xls = pd.ExcelFile(filepath)
    except FileNotFoundError:
        print(f"Error: The file '{filepath}' was not found.")
        return {}

    print(f"\n--- Processing Timetable ---")
    for sheet_name in xls.sheet_names:
        try:
            # Read the entire sheet without headers, as the structure is not a simple table
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

            # Process the sheet to get class schedules
            sheet_classes = process_timetable_sheet(df, sheet_name)

            # Add the extracted classes from the current sheet to the main dictionary
            all_classes_data.update(sheet_classes)
        except Exception as e:
            print(f"# ERROR: Could not process sheet '{sheet_name}'. Reason: {e}")

    return all_classes_data


def process_timetable_sheet(df: pd.DataFrame, sheet_name: str) -> Dict[str, Class]:
    """
    Processes a single timetable sheet to extract subjects and their schedules for each class.
    Assumes a specific format where each class has a subject row, a teacher row, and a blank row.
    """
    all_class_subjects = {}
    print(f"Processing sheet: '{sheet_name}'")

    # The actual data starts from the 3rd row (index 2 in pandas)
    # The structure is: subjects row, teachers row, empty row. So we step by 3.
    for i in range(2, len(df), 3):
        # Ensure we don't go past the end of the DataFrame
        if i + 1 >= len(df):
            break

        subject_row = df.iloc[i]
        teacher_row = df.iloc[i + 1]

        class_name = subject_row.iloc[0]
        # If the first cell in the subject row is empty, we assume it's the end of the class list
        if pd.isna(class_name):
            break

        class_name = str(class_name).strip()
        subjects_in_class = {}

        # A week has 5 days, and each day has 7 lesson slots
        for day_index in range(5):  # 0 for Monday, 1 for Tuesday, ...
            # Calculate the column range for the current day
            # Column B is index 1, so we start there
            start_col = 1 + (day_index * 7)
            end_col = start_col + 7

            for col_index in range(start_col, end_col):
                subject_name = subject_row.iloc[col_index]

                # If there's no subject in this slot, skip to the next one
                if pd.isna(subject_name):
                    continue

                teacher_name = teacher_row.iloc[col_index]

                # Clean and normalize the data
                normalized_name = str(subject_name).replace('\n', ' ').strip().lower()
                normalized_teacher = str(teacher_name).strip() if not pd.isna(teacher_name) else "No Teacher Assigned"

                # If we haven't seen this subject for this class yet, create a new Subject object
                if normalized_name not in subjects_in_class:
                    subject = Subject(name=normalized_name, teacher=normalized_teacher)
                    subjects_in_class[normalized_name] = subject

                # Increment the hour count for the current day
                subjects_in_class[normalized_name].hours_in_days[day_index] += 1

        current_class = Class(class_name, subjects_in_class)
        all_class_subjects[class_name] = current_class
        print(f"  -> Processed schedule for class '{class_name}' with {len(subjects_in_class)} unique subjects.")

    return all_class_subjects


def test2():
    all_class_data = extract_class_subjects()
    if not all_class_data:
        print("No class schedules were extracted.")
        return

    # Print schedule for a sample class, for example '1A'
    sample_class_name_1 = '10A'
    if sample_class_name_1 in all_class_data:
        print(f"\n--- Schedule for Class {sample_class_name_1} ---")
        class_subjects = all_class_data[sample_class_name_1].subjects
        for subject_name, subject_obj in class_subjects.items():
            print(f"  Subject: {subject_obj.name}")
            print(f"    Teacher: {subject_obj.teacher}")
            print(f"    Total Weekly Hours: {subject_obj.hours()}")
            print(f"    Hours per Day (Mon-Fri): {subject_obj.hours_in_days}")


def test1():
    all_days_in_quarters = extract_days()
    if not all_days_in_quarters:
        print("No days were extracted.")
        return

    for quarter, days_list in all_days_in_quarters.items():
        print(f"\n---> In quarter {quarter} there are {len(days_list)} date entries:")
        # Join the list of days for cleaner printing
        print(", ".join(filter(None, days_list)))


if __name__ == "__main__":
    test2()

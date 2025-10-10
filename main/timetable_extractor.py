import pandas as pd
import config
from typing import List, Dict, Optional
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


def extract_subjects(filepath=config.timetable_path) -> Dict[str, Dict[str, Subject]]:
    all_class_subjects = {}
    try:
        xls = pd.ExcelFile(filepath)
    except FileNotFoundError:
        print(f"Error: The file '{filepath}' was not found.")
        return all_class_subjects
    print(f"\n--- Processing Class Sheets ---")
    for sheet_name in xls.sheet_names:
        try:
            all_class_subjects = process_timetable_sheet(xls, sheet_name)
        except Exception as e:
            print(f"# ERROR: Could not process sheet '{sheet_name}'. Reason: {e}")
    return all_class_subjects


def process_timetable_sheet(xls, sheet_name) -> Optional[Dict[str, Subject]]:
    df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
    if len(df.columns) < 3:
        print(f"# WARNING: Skipping '{sheet_name}'. It has fewer than 3 columns.")
        return None
    df.dropna(how='all', inplace=True)
    subjects_list: List[str] = df.iloc[:, 0].tolist()
    teachers_list = df.iloc[:, 1].tolist()
    hours_list = df.iloc[:, 2].tolist()
    subjects_in_class = {}

    for subject_name, teacher, hours in zip(subjects_list, teachers_list, hours_list):
        if pd.isna(subject_name):
            continue
        try:
            num_hours = int(hours) if not pd.isna(hours) else 0
        except (ValueError, TypeError):
            num_hours = 0

        normalized_name = str(subject_name).strip().lower()

        subject = Subject(
            name=normalized_name,
            teacher=str(teacher),
            hours=num_hours
        )
        subjects_in_class[subject.name] = subject

    print(f"  -> Successfully created subject list for '{sheet_name}' with {len(subjects_in_class)} subjects.")
    return subjects_in_class


def test():
    all_days_in_quarters = extract_days()
    if not all_days_in_quarters:
        print("No days were extracted.")
        return

    # Correctly iterate over the dictionary using .items() to get both the
    # quarter (key) and the list of days (value).
    for quarter, days_list in all_days_in_quarters.items():
        print(f"\n---> In quarter {quarter} there are {len(days_list)} date entries:")
        # Join the list of days for cleaner printing
        print(", ".join(filter(None, days_list)))


if __name__ == "__main__":
    test()

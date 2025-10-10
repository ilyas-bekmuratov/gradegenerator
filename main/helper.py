import pandas as pd


def split_string_by_pattern(data_string: str, grades_per_student=7) -> list[list[int]]:
    result_lists = [[] for _ in range(grades_per_student)]
    for index, char in enumerate(data_string):
        result_lists[index % grades_per_student].append(int(char))
    return result_lists


def clean_grade(grade):
    """
    Cleans and standardizes a single grade value.
    - Converts pass/fail words to a special integer '1'.
    - Converts numbers like 4.0 to '4'.
    - Converts empty cells or other non-numeric text to '0'.
    """
    if pd.isna(grade) or str(grade).strip() == '':
        return '0'

    grade_str = str(grade).strip().lower()
    if grade_str in ["зачет", "зачёт", "сынақ", "есептелінді"]:
        return '1'  # Use '1' as a special marker for pass/fail

    try:
        return str(int(float(grade)))
    except (ValueError, TypeError):
        return '0'

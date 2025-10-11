import pandas as pd
from classes import Subject, Class
from typing import Dict, List
import config
import timetable_extractor
import writer
from os import path


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


def get_hours_this_quarter(
        subject: Subject, quarter_num: int,
        all_days_in_quarters: Dict[int, List[str]] = config.all_days_in_each_quarter
) -> int:
    return len(get_days_this_quarter(subject, quarter_num, all_days_in_quarters))


def get_days_this_quarter(
        subject: Subject, quarter_num: int,
        all_days_in_quarters: Dict[int, List[str]] = config.all_days_in_each_quarter
) -> List[str]:
    if len(all_days_in_quarters) == 0:
        print("all_days_in_quarters empty")
        return []

    days_this_quarter: List[str] = []
    # print(f"getting days for {subject.name} for quarter {quarter_num}")
    for idx, date in enumerate(all_days_in_quarters[quarter_num]):
        if date == "NaT":
            continue

        day = idx % 5
        hours_that_day = subject.hours_in_days[day]

        if hours_that_day == 0:
            continue

        for i in range(hours_that_day):
            # print(f"for day {get_day_name_by_index(day_idx=day)} added date {date}")
            days_this_quarter.append(date)

    return days_this_quarter


def get_quarter_start_index(
        subject: Subject, quarter_num: int,
        all_days_in_quarters: Dict[int, List[str]] = config.all_days_in_each_quarter
) -> int:
    index = 0
    for i in range(quarter_num):
        index += get_hours_this_quarter(subject, i+1, all_days_in_quarters)
    return index


def get_day_name_by_index(day_idx: int):
    if day_idx == 0:
        return "Monday"
    if day_idx == 1:
        return "Tuesday"
    if day_idx == 2:
        return "Wednesday"
    if day_idx == 3:
        return "Thursday"
    if day_idx == 4:
        return "Friday"
    if day_idx == 5:
        return "Saturday"
    if day_idx == 6:
        return "Sunday"


def full_test():
    all_classes: Dict[str, Class] = timetable_extractor.extract_class_subjects()
    class_obj: Class = all_classes["8F"]
    test_subject = class_obj.subjects['физика']
    quarter_num = 3
    days = get_days_this_quarter(test_subject, quarter_num)
    hours = get_hours_this_quarter(test_subject, quarter_num)
    index = get_quarter_start_index(test_subject, quarter_num)
    print(f"starting at index {index}, subject {test_subject} has {hours} hours {quarter_num} quarter \n {days}")
    output_path = path.join(config.output_dir, "test"+config.output_filename)
    writer.test(config.template_path, output_path, hours)


if __name__ == "__main__":
    full_test()

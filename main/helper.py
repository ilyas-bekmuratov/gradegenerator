import pandas as pd
from classes import Subject, Class
from typing import Dict, List, Any
import config
import openpyxl
import main
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


def get_dod_days(
        subject: Subject,
        all_days_in_quarters: Dict[int, List[str]] = config.all_days_in_each_quarter,
        skip_week=False
) -> List[str]:
    if len(all_days_in_quarters) == 0:
        print("all_days_in_quarters empty")
        return []

    skip = skip_week
    days: List[str] = []
    # print(f"getting days for {subject.name}")
    for q in range(1, 5):
        for idx, date in enumerate(all_days_in_quarters[q]):
            if date == "nan":
                continue

            day = idx % 5
            hours_that_day = subject.hours_in_days[day]

            if hours_that_day == 0:
                continue

            if skip_week and skip:
                skip = not skip
                continue

            for i in range(hours_that_day):
                days.append(date)

            if skip_week:
                skip = not skip

    print(f"     -> subject {subject} has {len(days)} days total, skip_week = {skip_week}")
    return days


def get_days_this_quarter(
        subject: Subject,
        quarter_num: int,
        all_days_in_quarters: Dict[int, List[str]] = config.all_days_in_each_quarter
) -> List[str]:
    if len(all_days_in_quarters) == 0:
        print("all_days_in_quarters empty")
        return []

    valid_q = [1, 2, 3, 4]
    if quarter_num not in valid_q:
        return []
    days_this_quarter: List[str] = []
    # print(f"getting days for {subject.name} for quarter {quarter_num}")
    for idx, date in enumerate(all_days_in_quarters[quarter_num]):
        if date == "nan":
            continue

        day = idx % 5
        hours_that_day = subject.hours_in_days[day]

        if hours_that_day == 0:
            continue

        for i in range(hours_that_day):
            # print(f"for day {get_day_name_by_index(day_idx=day)} added date {date}")
            days_this_quarter.append(date)

    return days_this_quarter


def split_by_proportion(list_to_split: List[Any],
                        original_part_sizes: List[int]) -> List[List[Any]]:
    # 1. Get the total lengths of both the original set and the new list
    total_original_size = sum(original_part_sizes)
    total_new_size = len(list_to_split)

    if total_original_size == 0:
        print("Error: Original part sizes cannot sum to zero.")
        return []

    # 2. Calculate the proportions of the original parts
    proportions = [size / total_original_size for size in original_part_sizes]

    # 3. Calculate the new integer sizes for the new list
    new_sizes = []
    running_total = 0

    # We must round the first N-1 parts (i.e., the first 3)
    # The last part (the 4th) is calculated by subtraction to ensure the
    # total sum is exactly correct and we don't lose items to rounding.

    num_parts = len(original_part_sizes)

    for i in range(num_parts - 1):
        # Calculate the ideal proportional size
        ideal_size = total_new_size * proportions[i]

        # Round to the nearest whole number
        actual_size = round(ideal_size)

        new_sizes.append(actual_size)
        running_total += actual_size

    # 4. Calculate the size of the last part
    # This is the "remainder" and ensures the sum is perfect.
    last_part_size = total_new_size - running_total
    new_sizes.append(last_part_size)

    # print(f"Original sizes: {original_part_sizes}")
    # print(f"Calculated new sizes: {new_sizes}")
    # print(f"Sum of new sizes: {sum(new_sizes)} (should be {total_new_size})\n")

    # 5. Now, use the new_sizes to split the list
    result_lists = []
    current_index = 0

    for size in new_sizes:
        # Get the slice from the list
        part = list_to_split[current_index: current_index + size]
        result_lists.append(part)

        # Move the index forward for the next slice
        current_index += size

    return result_lists


def get_quarter_start_index(
        subject: Subject,
        quarter_num: int,
        all_days_in_quarters: Dict[int, List[str]] = config.all_days_in_each_quarter
) -> int:
    if quarter_num == 5:
        return len(subject.topics)

    total = 0
    for q in range(1, 5):
        total += len(get_days_this_quarter(subject, q, all_days_in_quarters))
    sizes: List[int] = []
    for q in range(1, 5):
        days_num = len(get_days_this_quarter(subject, q, all_days_in_quarters))
        sizes.append(days_num)

    topics_split = split_by_proportion(subject.topics, sizes)
    # print(f"len(topics_split) = {len(topics_split)}")
    index = 0
    for q in range(quarter_num-1):
        # print(q)
        index += len(topics_split[q])
    # print(f"quarter index = {index}")
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


def get_month_from_date(date: str):
    month = date[3:5]
    if month == "9":
        return "Сентябрь"
    if month == "09":
        return "Сентябрь"
    if month == "10":
        return "Октябрь"
    if month == "11":
        return "Ноябрь"
    if month == "12":
        return "Декабрь"
    if month == "01":
        return "Январь"
    if month == "02":
        return "Февраль"
    if month == "03":
        return "Март"
    if month == "04":
        return "Апрель"
    if month == "05":
        return "Май"
    if month == "06":
        return "Июнь"
    return month


def get_repeat_str(subject_name: str, is_kaz: bool) -> str:
    repeat_str = config.kaz_repeat_str if is_kaz else config.rus_repeat_str
    if subject_name in config.eng_exception_subject_name:
        repeat_str = config.eng_repeat_str
    elif subject_name in config.kaz_exception_subject_name:
        repeat_str = config.kaz_repeat_str
    elif subject_name in config.rus_exception_subject_name:
        repeat_str = config.rus_repeat_str
    return repeat_str


def test_subject(current_class: Class,
                 class_number: int,
                 workbook,
                 subject_name: str,
                 quarters_to_test: List[int],
                 is_dod=False):
    current_subject = current_class.subjects[subject_name]

    split = 7 if (class_number >= 5 and current_subject.has_exam) else 5
    split_grades: list[list[int]] = split_string_by_pattern(current_subject.grades, split)
    for q in quarters_to_test:
        main.quarter(workbook, current_class, q, current_subject, split_grades, is_dod=is_dod)
        if is_dod:
            break


def full_test():
    is_dod = False
    classes_to_test = ["8D"]
    subjects_to_test = []
    quarters_to_test = [1, 2, 3, 4]
    output_path = str(path.join(config.output_dir, "test"+config.output_filename))
    changes_made = False

    template_path = config.template_path
    workbook = None
    try:
        workbook = openpyxl.load_workbook(template_path)
    except FileNotFoundError:
        print(f"Error: The template file '{template_path}' was not found.")
        return

    for class_str in classes_to_test:
        all_classes: Dict[str, Class] = main.extract_all_data(class_str, is_dod=is_dod)
        current_class: Class = all_classes[class_str]
        class_number = int(class_str[0])

        print(f"\nFULL-TEST   ->class {current_class.name} subjects: {subjects_to_test}")

        for subject_name, subject in current_class.subjects.items():
            # if subject.hours()>1:
            #     continue
            if not subjects_to_test or (subject_name in subjects_to_test):
                print(f"\n--- Processing Subject: {subject_name} ({subject.hours()}h/w) for class {current_class.name} ---")
                test_subject(current_class, class_number, workbook, subject_name, quarters_to_test, is_dod=is_dod)
                changes_made = True

    if changes_made:
        workbook.remove(workbook[config.template_sheet_name])
        workbook.remove(workbook[config.dod_template_sheet_name])
        workbook.save(output_path)


if __name__ == "__main__":
    # print("сынып сағаты" in config.no_grades)
    full_test()
    # main.extract_all_data(is_dod=True)

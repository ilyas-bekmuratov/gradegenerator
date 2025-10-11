output_dir = "reports"
output_filename = "journals.xlsx"

template_path = "template.xlsx"
template_sheet_name = "temp"

grades_path = "grades.xlsx"
subjects_path = "subjects.xlsx"
days_path = "days.xlsx"
timetable_path = "timetable.xlsx"

kaz_topics_path = "1-11 Kaz - Copy"
rus_topics_path = "1-11 Rus - Copy"

grade_bands = {
        2: (0, 39.99),
        3: (40, 64.99),
        4: (65, 84.99),
        5: (85, 100),
    }
DAILY_GRADE_BANDS = {
    2: (-3.0, -2.0),
    3: (-2.0, 0.0),
    4: (0.0, 4.0),
    5: (4.0, 7.0)
}
daily_grade_density = 0.66
weights = {
    'sop': 50,
    'so4': 50
    }
num_midterms = 3
max_midterms = 4
max_scores = [20, 20, 20, 20]

penalty_bonus_range = (-3.0, 7.0)
total_percent_mean_offset = -2.0
# Shifts the mean. E.g., -2.0 makes grades tend 2% lower in their band.
total_percent_sd = 3.0  # (max_pct - min_pct) / 4
# A smaller number (e.g., 3.0) makes scores more consistent, and larger number - more varied.

split_mean_offset = 0.0
# A positive number makes the final exam (СОч) contribute more; negative for midterms (СОр).
split_sd = 2.5  # default value = (max_so4_contrib - min_so4_contrib) / 4
# A smaller number (e.g., 2.0) makes the split very even, and arger number - very uneven splits.

subject_teacher_cell = [1, 2]  # B1
student_name_cell = [7, 2]
start_row = 7
max_row = 42
daily_grade_col = "C"
quarter_grade_col = "D"
year_grade_col = "M"
exam_grade_col = "N"
final_grade_col = "O"

date_col = "P"
topic_col = "Q"

all_days_in_each_quarter = {
    1: [
        "NaT", "NaT", "NaT", "NaT", "01.09.2023",
        "04.09.2023", "05.09.2023", "06.09.2023", "07.09.2023", "08.09.2023",
        "11.09.2023", "12.09.2023", "13.09.2023", "14.09.2023", "15.09.2023",
        "18.09.2023", "19.09.2023", "20.09.2023", "21.09.2023", "22.09.2023",
        "25.09.2023", "26.09.2023", "27.09.2023", "28.09.2023", "29.09.2023",
        "02.10.2023", "03.10.2023", "04.10.2023", "05.10.2023", "06.10.2023",
        "09.10.2023", "10.10.2023", "11.10.2023", "12.10.2023", "13.10.2023",
        "16.10.2023", "17.10.2023", "18.10.2023", "19.10.2023", "20.10.2023",
        "23.10.2023", "24.10.2023", "25.10.2023", "26.10.2023", "27.10.2023"
        ],
    2: [
        "NaT", "07.11.2023", "08.11.2023", "09.11.2023", "10.11.2023",
        "13.11.2023", "14.11.2023", "15.11.2023", "16.11.2023", "17.11.2023",
        "20.11.2023", "21.11.2023", "22.11.2023", "23.11.2023", "24.11.2023",
        "27.11.2023", "28.11.2023", "29.11.2023", "30.11.2023", "01.12.2023",
        "04.12.2023", "05.12.2023", "06.12.2023", "07.12.2023", "08.12.2023",
        "11.12.2023", "12.12.2023", "13.12.2023", "14.12.2023", "15.12.2023",
        "18.12.2023", "19.12.2023", "20.12.2023", "21.12.2023", "22.12.2023",
        "25.12.2023", "26.12.2023", "27.12.2023", "28.12.2023", "NaT"
        ],
    3: [
        "08.01.2024", "09.01.2024", "10.01.2024", "11.01.2024", "12.01.2024",
        "15.01.2024", "16.01.2024", "17.01.2024", "18.01.2024", "19.01.2024",
        "22.01.2024", "23.01.2024", "24.01.2024", "25.01.2024", "26.01.2024",
        "29.01.2024", "30.01.2024", "31.01.2024", "01.02.2024", "02.02.2024",
        "05.02.2024", "06.02.2024", "07.02.2024", "08.02.2024", "09.02.2024",
        "12.02.2024", "13.02.2024", "14.02.2024", "15.02.2024", "16.02.2024",
        "19.02.2024", "20.02.2024", "21.02.2024", "22.02.2024", "23.02.2024",
        "26.02.2024", "27.02.2024", "28.02.2024", "29.02.2024", "01.03.2024",
        "04.03.2024", "05.03.2024", "06.03.2024", "07.03.2024", "NaT",
        "11.03.2024", "12.03.2024", "13.03.2024", "14.03.2024", "15.03.2024"
        ],
    4: [
        "25.03.2024", "26.03.2024", "27.03.2024", "28.03.2024", "29.03.2024",
        "01.04.2024", "02.04.2024", "03.04.2024", "04.04.2024", "05.04.2024",
        "08.04.2024", "09.04.2024", "10.04.2024", "11.04.2024", "12.04.2024",
        "15.04.2024", "16.04.2024", "17.04.2024", "18.04.2024", "19.04.2024",
        "22.04.2024", "23.04.2024", "24.04.2024", "25.04.2024", "26.04.2024",
        "29.04.2024", "30.04.2024", "NaT", "02.05.2024", "03.05.2024",
        "06.05.2024", "NaT", "NaT", "NaT", "10.05.2024", "13.05.2024",
        "14.05.2024", "15.05.2024", "16.05.2024", "17.05.2024", "20.05.2024",
        "21.05.2024", "22.05.2024", "23.05.2024", "24.05.2024"
        ]
}

def get_daily_grade_distribution(bonus):
    """Determines the primary daily grade and a secondary grade based on the bonus."""
    primary_grade = 4  # Default grade
    for grade, (min_bonus, max_bonus) in DAILY_GRADE_BANDS.items():
        if min_bonus <= bonus < max_bonus:
            primary_grade = grade
            break

    # Create a weighted distribution for more realistic grades
    # The primary grade is highly likely, with a small chance of an adjacent grade
    if primary_grade == 5:
        return {10: 0.75, 9: 0.1, 8: 0.06, 7: 0.04, 6: 0.04, 5: 0.007, 4: 0.003}  # Mostly 5s, some 4s
    if primary_grade == 4:
        return {10: 0.05, 9: 0.1, 8: 0.65, 7: 0.1, 6: 0.05, 5: 0.044, 4: 0.005, 3: 0.001}
    if primary_grade == 3:
        return {10: 0.01, 9: 0.015, 8: 0.025, 7: 0.05, 6: 0.1, 5: 0.65, 4: 0.1, 3: 0.045, 2: 0.005}
    if primary_grade == 2:
        return {10: 0.002, 9: 0.003, 8: 0.005, 7: 0.015, 6: 0.25, 5: 0.05, 4: 0.1, 3: 0.65, 2: 0.1}
    return {10: 0.05, 9: 0.1, 8: 0.65, 7: 0.1, 6: 0.05, 5: 0.044, 4: 0.005, 3: 0.001}  # Default case same as 4

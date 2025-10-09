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
daily_grade_col = "C"
quarter_grade_col = "D"
year_grade_col = "M"
exam_grade_col = "N"
final_grade_col = "O"

date_col = "P"
topic_col = "Q"


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

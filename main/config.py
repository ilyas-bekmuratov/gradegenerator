output_dir = "reports"
output_filename = "journals.xlsx"
template_path = "reports/templates.xlsx"
grades_path = "reports/grades.xlsx"
subjects_path = "reports/subjects.xlsx"
kaz_topics_path = ""
rus_topics_path = ""

subject_name_cell = [1, 2]
grade_bands: {
        2: (0, 39.99),
        3: (40, 64.99),
        4: (65, 84.99),
        5: (85, 100),
    }

weights = {'sop': 50, 'so4': 50}
num_midterms = 3
max_midterms = 4
max_scores = [20, 20, 20, 20]

penalty_bonus_range = (-7.0, 7.0)
total_percent_mean_offset = -2.0
# Shifts the mean. E.g., -2.0 makes grades tend 2% lower in their band.
total_percent_sd = 3.0  # (max_pct - min_pct) / 4
# A smaller number (e.g., 3.0) makes scores more consistent.
# A larger number (e.g., 8.0) makes them more varied.

split_mean_offset = 0.0
# A positive number makes the final exam (СОч) contribute more; negative for midterms (СОр).
split_sd = 2.5  # default value = (max_so4_contrib - min_so4_contrib) / 4
# A smaller number (e.g., 2.0) makes the split very even.
# A larger number (e.g., 6.0) allows for very uneven splits.

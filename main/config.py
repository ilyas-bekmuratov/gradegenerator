output_dir = "reports"
output_filename = "journals.xlsx"
template_path = "reports/templates.xlsx"
grades_path = "reports/grades.xlsx"
subjects_path = "reports/subjects.xlsx"
kaz_topics_path = ""
rus_topics_path = ""

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

subject_name_cell = [1, 2]
# Format: hours: { quarter: (template_sheet_name, start_column_letter) }
TEMPLATE_MAPPINGS = {
    1: {
        1: ("1ч1р", "L"), 2: ("2ч1р", "L"), 3: ("3ч1р", "O"), 4: ("4ч1р", "K"),
    },
    2: {
        1: ("1ч2р", "T"), 2: ("2ч2р", "T"), 3: ("3ч2р", "Z"), 4: ("4ч2р", "R"),
    },
    3: {
        1: ("1ч3р", "AB"), 2: ("2ч3р", "AA"), 3: ("3ч3р", "AH"), 4: ("4ч3р", "Y"),
    },
    4: {
        1: ("1ч4р", "AI"), 2: ("2ч4р", "AH"), 3: ("3ч4р", "AQ"), 4: ("4ч4р", "AE"),
    },
    5: {
        1: ("1ч5р", "AP"), 2: ("2ч5р", "AO"), 3: ("3ч5р", "AZ"), 4: ("4ч5р", "AK"),
    },
    6: {
        1: ("1ч6р", "AW"), 2: ("2ч6р", "AV"), 3: ("3ч6р", "BI"), 4: ("4ч6р", "AQ"),
    },
}
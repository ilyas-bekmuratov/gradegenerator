settings: dict = {
    # --- Output Settings ---
    "output_dir": "reports",  # Folder to save the report in
    "output_filename": "generated_grades.xlsx",  # The name of the Excel file

    # --- Grading Bands ---
    "grade_bands": {
        2: (0, 39.99),
        3: (40, 64.99),
        4: (65, 84.99),
        5: (85, 100),
    },

    # --- Exam Structure and Weighting ---
    "weights": {'sop': 50, 'so4': 50},
    "num_midterms": 3,

    # --- Maximum Scores for Each Exam ---
    # List of max scores: one for each midterm, plus the last one for the final exam.
    # The number of items must be num_midterms + 1.
    "max_scores": [20, 20, 20, 20],

    "final_grades": [5, 5, 4, 3, 5, 4, 2, 5, 4, 3, 5],
    "sheet_name": "some sheet 2",

    # --- Generation Settings ---
    "penalty_bonus_range": (-7.0, 7.0),
}
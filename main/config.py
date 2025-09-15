settings: dict = {
    # --- Output Settings ---
    "output_dir": "reports",  # Folder to save the report in
    "output_filename": "generated_grades.ods",  # The name of the Excel file

    # --- Grading Bands ---
    "grade_bands": {
        2: (0, 39.99),
        3: (40, 64.99),
        4: (65, 84.99),
        5: (85, 100),
    },

    # For the total percentage within a grade band:
    "total_percent_mean_offset": -2.0,  # Shifts the mean. E.g., -2.0 makes grades tend 2% lower in their band.
    "total_percent_sd": 3.0,
    # A smaller number (e.g., 3.0) makes scores more consistent. A larger number (e.g., 8.0) makes them more varied.

    # For the split between midterm/final contribution:
    "split_mean_offset": 0.0,
    # A positive number makes the final exam (СОч) contribute more; negative for midterms (СОр).
    "split_sd": 2.5,
    # A smaller number (e.g., 2.0) makes the split very even. A larger number (e.g., 6.0) allows for very uneven splits.

    # --- Exam Structure and Weighting ---
    "weights": {'sop': 50, 'so4': 50},
    "num_midterms": 3,

    # --- Maximum Scores for Each Exam ---
    # List of max scores: one for each midterm, plus the last one for the final exam.
    # The number of items must be num_midterms + 1.
    "max_scores": [20, 20, 20, 20],

    "final_grades": "55555223333344455445333333344444555554455555544444555554444455555224434444455555443342233333444444444444",
    "sheet_name": "Иностранный язык (Английский) 1 четверть",
    "current_quarter": 0,  #0, 1, 2, 3 stands for first, second, third, and fourth

    # --- Generation Settings ---
    "penalty_bonus_range": (-7.0, 7.0),
}

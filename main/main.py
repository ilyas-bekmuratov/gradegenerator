import pandas as pd
import grade_generator as gg
"""
Generates plausible midterm and final exam scores for a given final grade mark
by reverse-engineering the grading process.
### How to Use the Script
0.  **Install libraries:** If you don't have them, open your terminal or command prompt and run:
    `pip install pandas numpy`
1.  **Run the script:** Execute the file from your terminal:
     `python grade_generator.py`
2. You can change the input by modifying the `final_grades_list` within the main.py.
To use a `.csv` file, you can uncomment the lines for reading a CSV and specify your column name.
"""


def main():
    # --- Main Configuration - Adjust these values as needed ---
    config = {
        "grade_bands": {
            2: (0, 39.99),
            3: (40, 64.99),
            4: (65, 84.99),
            5: (85, 100),
        },
        "weights": {'sop': 50, 'so4': 50},
        "num_midterms": 3,
        "max_score": 20,
        "penalty_bonus_range": (-7.0, 7.0),  # The +- range for penalties/bonuses
    }

    # --- INPUT: Provide your list of final grades here ---
    # This list can be populated from a CSV file
    # For example:
    # df = pd.read_csv('your_grades.csv')
    # final_grades_list = df['final_grade_column_name'].tolist()

    final_grades_list = [4, 3, 5, 2, 4, 4, 3]  # Example list

    # --- Processing ---
    results = []
    for grade in final_grades_list:
        if grade in config['grade_bands']:
            generated_data = gg.generate_plausible_grades(grade, config)
            results.append(generated_data)

    # --- OUTPUT: Display the results in a clean table ---
    output_df = pd.DataFrame(results)
    print("Generated Plausible Student Scores:")
    print(output_df.to_string())

    # Optional: Save the generated data to a new CSV file
    # output_df.to_csv('generated_scores.csv', index=False)


if __name__ == "__main__":
    main()

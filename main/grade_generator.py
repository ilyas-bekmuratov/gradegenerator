import numpy as np
import random


def generate_plausible_grades(final_grade_mark, config):
    # --- Unpack configuration variables for easy access ---
    grade_bands = config['grade_bands']
    weights = config['weights']
    num_midterms = config['num_midterms']
    max_score = config['max_score']
    penalty_bonus_range = config['penalty_bonus_range']

    # --- 1. Generate Total Percentage from the Final Grade's Range ---
    min_pct, max_pct = grade_bands[final_grade_mark]
    mean_pct = (min_pct + max_pct) / 2
    std_dev_pct = (max_pct - min_pct) / 4 if max_pct > min_pct else 1
    total_percent = np.random.normal(loc=mean_pct, scale=std_dev_pct)
    total_percent = np.clip(total_percent, min_pct, max_pct)

    # --- 2. Split Total Percentage between СОр and СОч ---
    min_so4_contrib = max(0, total_percent - weights['sop'])
    max_so4_contrib = min(weights['so4'], total_percent)
    mean_split = (min_so4_contrib + max_so4_contrib) / 2
    std_dev_split = (max_so4_contrib - min_so4_contrib) / 4 if max_so4_contrib > min_so4_contrib else 1
    so4_percent_contribution = np.random.normal(loc=mean_split, scale=std_dev_split)
    so4_percent_contribution = np.clip(so4_percent_contribution, min_so4_contrib, max_so4_contrib)
    sop_percent_contribution = total_percent - so4_percent_contribution

    # --- 3. Calculate СОч Score and Handle Rounding ---
    so4_score_float = (so4_percent_contribution / weights['so4']) * max_score if weights['so4'] > 0 else 0
    so4_score_rounded = int(round(so4_score_float))
    so4_score_rounded = np.clip(so4_score_rounded, 0, max_score)
    actual_so4_contribution = (so4_score_rounded / max_score) * weights['so4'] if max_score > 0 else 0

    # --- 4. Adjust СОр Contribution based on rounding ---
    rounding_diff = so4_percent_contribution - actual_so4_contribution
    adjusted_sop_contribution = sop_percent_contribution + rounding_diff

    # --- 5. Apply Inverse Penalty/Bonus to find the "raw" score ---
    penalty_bonus = np.random.uniform(penalty_bonus_range[0], penalty_bonus_range[1])
    raw_sop_contribution = adjusted_sop_contribution - penalty_bonus
    raw_sop_contribution = np.clip(raw_sop_contribution, 0, weights['sop'])

    # --- 6. Generate Individual СОр Midterm Scores ---
    target_sum = 0
    if weights['sop'] > 0 and max_score > 0:
        target_sum = int(round((raw_sop_contribution / weights['sop']) * max_score * num_midterms))

    midterm_scores = [0] * num_midterms
    if target_sum > 0:
        for _ in range(target_sum):
            available_indices = [i for i, score in enumerate(midterm_scores) if score < max_score]
            if not available_indices:
                break
            chosen_index = random.choice(available_indices)
            midterm_scores[chosen_index] += 1

    return {
        "СОр Scores": midterm_scores,
        "СОч Score": so4_score_rounded,
        "Actual СОр %": round(sop_percent_contribution, 1),
        "Penalty/Bonus": round(penalty_bonus, 1),
        "СОр %": round(adjusted_sop_contribution, 1),
        "СОч %": round(actual_so4_contribution, 1),
        "Сумма %": round(total_percent, 1),
        "Оценка": final_grade_mark
    }

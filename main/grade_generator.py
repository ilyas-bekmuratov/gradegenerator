import numpy as np
import random


def generate_plausible_grades(final_grade_mark, config):
    # --- Unpack configuration variables ---
    settings = config.settings
    grade_bands = settings['grade_bands']
    weights = settings['weights']
    num_midterms = settings['num_midterms']
    penalty_bonus_range = settings['penalty_bonus_range']
    midterm_max_scores = settings['max_scores'][:num_midterms]
    so4_max_score = settings['max_scores'][-1]
    total_max_midterm_score = sum(midterm_max_scores)

    # --- 1. Generate Total Percentage ---
    min_pct, max_pct = grade_bands[final_grade_mark]
    mean_pct = (min_pct + max_pct) / 2
    mean_pct += settings.get('total_percent_mean_offset', 0.0)
    std_dev_pct = settings.get('total_percent_sd', (max_pct - min_pct) / 4)

    total_percent = np.random.normal(loc=mean_pct, scale=std_dev_pct)
    total_percent = np.clip(total_percent, min_pct, max_pct)

    # --- 2. Split Total Percentage ---
    min_so4_contrib = max(0, total_percent - weights['sop'])
    max_so4_contrib = min(weights['so4'], total_percent)

    mean_split = (min_so4_contrib + max_so4_contrib) / 2
    mean_split += settings.get('split_mean_offset', 0.0)
    std_dev_split = settings.get('split_sd', (max_so4_contrib - min_so4_contrib) / 4)

    so4_percent_contribution = np.random.normal(loc=mean_split, scale=std_dev_split)
    so4_percent_contribution = np.clip(so4_percent_contribution, min_so4_contrib, max_so4_contrib)
    sop_percent_contribution = total_percent - so4_percent_contribution

    # --- 3. Calculate Final Exam Score (СОч) ---
    so4_score_float = (so4_percent_contribution / weights['so4']) * so4_max_score if weights['so4'] > 0 else 0
    so4_score_rounded = int(round(so4_score_float))
    so4_score_rounded = np.clip(so4_score_rounded, 0, so4_max_score)
    actual_so4_contribution = (so4_score_rounded / so4_max_score) * weights['so4'] if so4_max_score > 0 else 0

    # --- 4. Adjust Midterm (СОр) Contribution ---
    rounding_diff = so4_percent_contribution - actual_so4_contribution
    adjusted_sop_contribution = (sop_percent_contribution + rounding_diff)
    adjusted_sop_contribution = np.clip(adjusted_sop_contribution, 0, weights['sop'])

    # --- 5. Apply Inverse Penalty/Bonus ---
    penalty_bonus = np.random.uniform(penalty_bonus_range[0], penalty_bonus_range[1])
    raw_sop_contribution = adjusted_sop_contribution - penalty_bonus
    raw_sop_contribution = np.clip(raw_sop_contribution, 0, weights['sop'])

    # --- 6. Generate Midterm (СОр) Scores ---
    target_sum = 0
    if weights['sop'] > 0 and total_max_midterm_score > 0:
        target_sum = int(round((raw_sop_contribution / weights['sop']) * total_max_midterm_score))

    midterm_scores = [0] * num_midterms
    if target_sum > 0:
        for _ in range(target_sum):
            available_indices = [i for i, score in enumerate(midterm_scores) if score < midterm_max_scores[i]]
            if not available_indices:
                break
            chosen_index = random.choice(available_indices)
            midterm_scores[chosen_index] += 1

    return {
        "Input Grade": final_grade_mark,
        "Generated Total %": round(total_percent, 1),
        "СОч Score (Final)": so4_score_rounded,
        "СОр Scores (Midterms)": midterm_scores,
        "Adjusted СОр %": round(adjusted_sop_contribution, 1),
        "Actual СОч %": round(actual_so4_contribution, 1),
        "Penalty/Bonus Applied": round(penalty_bonus, 1),
    }

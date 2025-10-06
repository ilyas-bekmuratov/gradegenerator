import numpy as np
import random
import config
from class_class import Class, Subject


def generate_plausible_grades(final_grade_mark, current_class: Class, subject: Subject):
    midterm_max_scores = config.max_scores[:config.num_midterms]
    so4_max_score = config.max_scores[-1]
    total_max_midterm_score = sum(midterm_max_scores)

    # --- 1. Generate Total Percentage ---
    min_pct, max_pct = config.grade_bands[final_grade_mark]

    mean_pct = (min_pct + max_pct) / 2
    mean_pct += config.total_percent_mean_offset
    std_dev_pct = config.total_percent_sd

    total_percent = np.random.normal(loc=mean_pct, scale=std_dev_pct)
    total_percent = np.clip(total_percent, min_pct, max_pct)

    # Initialize penalty/bonus
    penalty_bonus = np.random.uniform(config.penalty_bonus_range[0], config.penalty_bonus_range[1])

    if config.weights.get('so4', 0) == 0:
        # --- CASE: NO FINAL EXAM (СОч weight is 0) ---
        adjusted_sop_contribution = np.clip(total_percent, 0, config.weights['sop'])
        so4_score_rounded = '-'
        actual_so4_contribution = '-'
        raw_sop_contribution = adjusted_sop_contribution - penalty_bonus
        raw_sop_contribution = np.clip(raw_sop_contribution, 0, config.weights['sop'])

    else:
        # --- CASE: FINAL EXAM EXISTS ---
        min_so4_contrib = max(0, total_percent - config.weights['sop'])
        max_so4_contrib = min(config.weights['so4'], total_percent)

        mean_split = (min_so4_contrib + max_so4_contrib) / 2
        mean_split += config.split_mean_offset
        std_dev_split = config.split_sd

        so4_percent_contribution = np.random.normal(loc=mean_split, scale=std_dev_split)
        so4_percent_contribution = np.clip(so4_percent_contribution, min_so4_contrib, max_so4_contrib)
        sop_percent_contribution = total_percent - so4_percent_contribution

        so4_score_float = (so4_percent_contribution / config.weights['so4']) * so4_max_score if config.weights['so4'] > 0 else 0
        so4_score_rounded = int(round(so4_score_float))
        so4_score_rounded = np.clip(so4_score_rounded, 0, so4_max_score)
        actual_so4_contribution = (so4_score_rounded / so4_max_score) * config.weights['so4'] if so4_max_score > 0 else 0

        rounding_diff = so4_percent_contribution - actual_so4_contribution
        adjusted_sop_contribution = sop_percent_contribution + rounding_diff
        adjusted_sop_contribution = np.clip(adjusted_sop_contribution, 0, config.weights['sop'])

        raw_sop_contribution = adjusted_sop_contribution - penalty_bonus
        raw_sop_contribution = np.clip(raw_sop_contribution, 0, config.weights['sop'])

    # --- Generate Midterm (СОр) Scores ---
    target_sum = 0
    if config.weights['sop'] > 0 and total_max_midterm_score > 0:
        target_sum = int(round((raw_sop_contribution / config.weights['sop']) * total_max_midterm_score))

    midterm_scores = [0] * config.num_midterms
    if target_sum > 0:
        for _ in range(target_sum):
            available_indices = [i for i, score in enumerate(midterm_scores) if score < midterm_max_scores[i]]
            if not available_indices:
                break
            chosen_index = random.choice(available_indices)
            midterm_scores[chosen_index] += 1

    # --- Format numbers for the return dictionary ---
    final_sop_percent = round(adjusted_sop_contribution, 1)
    final_so4_percent = round(actual_so4_contribution, 1) \
        if isinstance(actual_so4_contribution, (int, float)) \
        else actual_so4_contribution

    return {
        "Input Grade": final_grade_mark,
        "Generated Total %": round(total_percent, 1),
        "СОч Score (Final)": so4_score_rounded,
        "СОр Scores (Midterms)": midterm_scores,
        "Adjusted СОр %": final_sop_percent,
        "Actual СОч %": final_so4_percent,
        "Penalty/Bonus Applied": round(penalty_bonus, 1),
    }

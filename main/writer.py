import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
import config
from copy import copy


def extend_day_columns(sheet, num_copies, is_last_quarter=False, has_exam=False):
    daily_grade_styles, daily_grade_width = read_styles_and_width(sheet, config.daily_grade_col)
    quarter_styles, quarter_grade_width = read_styles_and_width(sheet, config.quarter_grade_col)
    date_styles, date_width = read_styles_and_width(sheet, config.date_col)
    topic_hw_styles, topic_hw_width = read_styles_and_width(sheet, config.topic_col)

    daily_grade_col_idx = column_index_from_string(config.daily_grade_col)
    new_merges = []
    for merged_range in list(sheet.merged_cells.ranges):
        if daily_grade_col_idx > merged_range.min_col:
            continue

        sheet.unmerge_cells(str(merged_range))

        new_min_col = merged_range.min_col + num_copies - 1
        new_max_col = merged_range.max_col + num_copies - 1
        new_range_str = f"{get_column_letter(new_min_col)}{merged_range.min_row}:{get_column_letter(new_max_col)}{merged_range.max_row}"
        new_merges.append(new_range_str)

    col_idx = column_index_from_string(config.date_col)

    quarter_to_dates_offset = 12
    if not is_last_quarter:
        quarter_to_dates_offset -= 3
        print("      -> not the last quarter removed 3 columns")
        sheet.delete_cols(col_idx + quarter_to_dates_offset)  # delete the final grade, exam, and summary grade columns
        sheet.delete_cols(col_idx + quarter_to_dates_offset)
        sheet.delete_cols(col_idx + quarter_to_dates_offset)
    elif not has_exam:
        print("      -> has no exam, removed 2 columns")
        quarter_to_dates_offset -= 2  # delete the exam and summary grade columns
        sheet.delete_cols(col_idx + quarter_to_dates_offset)
        sheet.delete_cols(col_idx + quarter_to_dates_offset)

    sheet.insert_cols(daily_grade_col_idx, num_copies - 1)

    final_idx = col_idx + num_copies
    while col_idx < final_idx:
        current_col_letter = get_column_letter(col_idx)
        # print(f"Applying template to new column '{current_col_letter}' applied width = {daily_grade_width}")
        sheet.column_dimensions[current_col_letter].width = daily_grade_width
        for row_idx, style_array in daily_grade_styles.items():
            sheet.cell(row=row_idx, column=col_idx)._style = style_array
        col_idx += 1

    final_idx = final_idx + quarter_to_dates_offset
    while col_idx < final_idx:
        current_col_letter = get_column_letter(col_idx)
        # print(f"quarter template '{current_col_letter}' applied width {quarter_grade_width}")
        sheet.column_dimensions[current_col_letter].width = quarter_grade_width
        # for row_idx, style_array in quarter_styles.items():
        #     sheet.cell(row=row_idx, column=col_idx)._style = style_array
        col_idx += 1

    sheet.column_dimensions[get_column_letter(col_idx)].width = date_width
    for row_idx, style_array in date_styles.items():
        sheet.cell(row=row_idx, column=col_idx)._style = style_array
    col_idx += 1

    sheet.column_dimensions[get_column_letter(col_idx)].width = topic_hw_width
    for row_idx, style_array in topic_hw_styles.items():
        sheet.cell(row=row_idx, column=col_idx)._style = style_array
    col_idx += 1
    sheet.column_dimensions[get_column_letter(col_idx)].width = topic_hw_width
    for row_idx, style_array in topic_hw_styles.items():
        sheet.cell(row=row_idx, column=col_idx)._style = style_array

    for merge_str in new_merges:
        sheet.merge_cells(merge_str)


def read_styles_and_width(sheet, col: str):
    styles = {}
    width = sheet.column_dimensions[col].width

    for row_idx in range(1, config.max_row):
        cell = sheet[f'{col}{row_idx}']
        if cell.has_style:
            styles[row_idx] = copy(cell._style)

    return styles, width


def print_widths(sheet):
    for i in range(1, sheet.max_column):
        letter = get_column_letter(i)
        width = sheet.column_dimensions[letter].width
        print(f"before column at {letter} has width {width}")


def test(source_file, output_file, num_copies):
    workbook = None
    try:
        workbook = openpyxl.load_workbook(source_file)
    except FileNotFoundError:
        print(f"Error: The template file '{source_file}' was not found.")
        return

    sheet = workbook[config.template_sheet_name]

    # print_widths(sheet)

    extend_day_columns(sheet, num_copies)

    # print_widths(sheet)

    workbook.save(output_file)
    print(f"\nSuccessfully created '{output_file}' with {num_copies} formatted columns.")

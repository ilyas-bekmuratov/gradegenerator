import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
import config
from copy import copy


def extend_day_columns(sheet, num_copies, is_last_quarter=False, has_exam=False, is_dod=False):
    daily_grade_styles, daily_grade_width = read_styles_and_width(sheet, config.daily_grade_col)
    col_letter = config.dod_grade_col if is_dod else config.quarter_grade_col
    quarter_styles, quarter_grade_width = read_styles_and_width(sheet, col_letter)
    date_styles, date_width = read_styles_and_width(sheet, config.date_col)
    topic_hw_styles, topic_hw_width = read_styles_and_width(sheet, config.topic_col)

    print_widths(sheet, "\ninitial")
    daily_grade_col_idx = column_index_from_string(config.daily_grade_col)
    col_idx = column_index_from_string(col_letter)

    quarter_to_dates_offset = config.quarter_to_dates_offset
    amount_cols_to_delete = 0
    if not is_last_quarter:
        print("      -> not the last quarter removed 3 columns")
        amount_cols_to_delete = 3  # delete the final grade, exam, and summary grade columns
    elif not has_exam:
        print("      -> has no exam, removed 2 columns")
        amount_cols_to_delete = 2  # delete the exam and summary grade columns

    quarter_to_dates_offset -= amount_cols_to_delete
    if amount_cols_to_delete > 0:
        sheet.delete_cols(col_idx + quarter_to_dates_offset, amount=amount_cols_to_delete)
    print_widths(sheet, f"after deletion of {amount_cols_to_delete} columns at index {col_idx + quarter_to_dates_offset}")

    new_merges = []
    for merged_range in list(sheet.merged_cells.ranges):
        if daily_grade_col_idx > merged_range.min_col:
            continue
        try:
            sheet.unmerge_cells(str(merged_range))

            new_min_col = merged_range.min_col + num_copies - 1
            new_max_col = merged_range.max_col + num_copies - 1
            new_range_str = (f"{get_column_letter(new_min_col)}{merged_range.min_row}" +
                             f":{get_column_letter(new_max_col)}{merged_range.max_row}")
            new_merges.append(new_range_str)
        except Exception as e:
            print(f"warning during merge manipulation: {e}")

    print(f"merges to restore = {new_merges}")

    sheet.insert_cols(daily_grade_col_idx, num_copies - 1)
    print_widths(sheet, f"after insertion of {num_copies - 1} columns")

    col_idx = daily_grade_col_idx
    final_idx = col_idx + num_copies
    while col_idx < final_idx:
        current_col_letter = get_column_letter(col_idx)
        sheet.column_dimensions[current_col_letter].width = daily_grade_width
        for row_idx, style_array in daily_grade_styles.items():
            sheet.cell(row=row_idx, column=col_idx)._style = style_array
        col_idx += 1

    i = 0
    final_idx = final_idx + quarter_to_dates_offset
    while col_idx < final_idx:
        current_col_letter = get_column_letter(col_idx)
        sheet.column_dimensions[current_col_letter].width = quarter_grade_width
        col_idx += 1
        i += 1

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
    print_widths(sheet, "after merging back")
    return sheet


def read_styles_and_width(sheet, col: str):
    styles = {}
    width = sheet.column_dimensions[col].width

    for row_idx in range(1, config.max_row):
        cell = sheet[f'{col}{row_idx}']
        if cell.has_style:
            styles[row_idx] = copy(cell._style)

    return styles, width


def print_widths(sheet, message):
    widths = {}
    for i in range(1, sheet.max_column):
        letter = get_column_letter(i)
        widths[letter] = sheet.column_dimensions[letter].width
    print(f"\n{message}\ncolumn widths = {widths}")


def test(source_file, output_file, num_copies):
    workbook = None
    try:
        workbook = openpyxl.load_workbook(source_file)
    except FileNotFoundError:
        print(f"Error: The template file '{source_file}' was not found.")
        return

    sheetname = "9F - казахский язык и лите - Q4"
    sheet = workbook[sheetname]
    print_widths(sheet, "after saving")
    print(sheet.column_dimensions["AP"].width)
    print(sheet.column_dimensions["AU"].width)

    # extend_day_columns(sheet, num_copies)
    # print(sheet.column_dimensions["AJ"].width)

    # print_widths(sheet)

    # workbook.save(output_file)
    print(f"\nSuccessfully created '{output_file}' with {num_copies} formatted columns.")


if __name__ == "__main__":
    filepath = "reports/testjournals.xlsx"
    test(filepath, "copy", 25)

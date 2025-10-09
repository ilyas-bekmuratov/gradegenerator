import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
import config

OUTPUT_FILE = 'copied.xlsx'
total_columns = 5
max_row = 42


def replicate_formatted_column(source_file, output_file, num_copies):
    workbook = None
    try:
        workbook = openpyxl.load_workbook(source_file)
    except FileNotFoundError:
        print(f"Error: The template file '{source_file}' was not found.")
        return

    sheet = workbook[config.template_sheet_name]

    template_styles = {}
    daily_grade_col_idx = column_index_from_string(config.daily_grade_col)
    for row_idx in range(1, max_row):
        cell = sheet[f'{config.daily_grade_col}{row_idx}']
        if cell.has_style:
            template_styles[row_idx] = cell._style

    for i in range(1, sheet.max_column):
        letter = get_column_letter(i)
        width = sheet.column_dimensions[letter].width
        print(f"before column at {letter} has width {width}")

    daily_grade_width = sheet.column_dimensions[config.daily_grade_col].width
    quarter_grade_width = sheet.column_dimensions[config.quarter_grade_col].width
    date_width = sheet.column_dimensions[config.date_col].width
    topic_hw_width = sheet.column_dimensions[config.topic_col].width

    new_merges = []
    for merged_range in list(sheet.merged_cells.ranges):
        if daily_grade_col_idx > merged_range.min_col:
            continue

        sheet.unmerge_cells(str(merged_range))

        new_min_col = merged_range.min_col + num_copies
        new_max_col = merged_range.max_col + num_copies
        new_range_str = f"{get_column_letter(new_min_col)}{merged_range.min_row}:{get_column_letter(new_max_col)}{merged_range.max_row}"
        new_merges.append(new_range_str)

    sheet.insert_cols(daily_grade_col_idx, num_copies)

    for merge_str in new_merges:
        sheet.merge_cells(merge_str)

    col_idx = daily_grade_col_idx
    final_idx = col_idx + num_copies
    while col_idx <= final_idx:
        current_col_letter = get_column_letter(col_idx)
        print(f"Applying template to new column '{current_col_letter}' applied width = {daily_grade_width}")
        sheet.column_dimensions[current_col_letter].width = daily_grade_width
        for row_idx, style_array in template_styles.items():
            sheet.cell(row=row_idx, column=col_idx)._style = style_array
        col_idx += 1

    final_idx = final_idx + 12
    while col_idx <= final_idx:
        current_col_letter = get_column_letter(col_idx)
        print(f"quarter template '{current_col_letter}' applied width {quarter_grade_width}")
        sheet.column_dimensions[current_col_letter].width = quarter_grade_width
        col_idx += 1

    sheet.column_dimensions[get_column_letter(col_idx)].width = date_width
    sheet.column_dimensions[get_column_letter(col_idx + 1)].width = topic_hw_width
    sheet.column_dimensions[get_column_letter(col_idx + 2)].width = topic_hw_width

    for sheet_name in list(workbook.sheetnames):
        if sheet_name != config.template_sheet_name:
            workbook.remove(workbook[sheet_name])

    for i in range(1, sheet.max_column):
        letter = get_column_letter(i)
        width = sheet.column_dimensions[letter].width
        print(f"after column at {letter} has width {width}")

    workbook.save(output_file)
    print(f"\nSuccessfully created '{output_file}' with {num_copies} formatted columns.")


if __name__ == "__main__":
    replicate_formatted_column(
        source_file=config.template_path,
        output_file=OUTPUT_FILE,
        num_copies=total_columns
    )


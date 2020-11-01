import xlrd
import xlwt


def cat_workbook(*workbooks):
    cat_wb = xlwt.Workbook()

    no_sheets = workbooks[0].nsheets
    write_row_idx_sheet = {}

    for wb in workbooks:
        for sheet_idx in range(no_sheets):
            sheet = wb.sheet_by_index(sheet_idx)

            try:
                cat_sheet = cat_wb.get_sheet(sheet.name)
            except Exception as e:
                cat_sheet = cat_wb.add_sheet(sheet.name)
                write_row_idx_sheet[sheet.name] = 0

            for row_idx, row in enumerate(sheet.get_rows()):
                write_col_idx = 0
                for column in row:
                    cat_sheet.write(write_row_idx_sheet[sheet.name], write_col_idx, column.value)
                    write_col_idx += 1
                write_row_idx_sheet[sheet.name] += 1

    return cat_wb


def cat_excel(filename_a, filename_b, filename_result):
    new_wb = cat_workbook(
        xlrd.open_workbook(filename_a),
        xlrd.open_workbook(filename_b),
        xlrd.open_workbook(filename_a),
    )
    new_wb.save(filename_result)


if __name__ == '__main__':
    cat_excel(
        'excel_files/a.xlsx',
        'excel_files/b.xlsx',
        'excel_files/result.xls'
    )

import xlrd
import xlwt


def cat_workbook(workbook_a, workbook_2):
    new_wb = xlwt.Workbook()

    for sheet_idx in range(workbook_a.nsheets):
        sheet_a = workbook_a.sheet_by_index(sheet_idx)
        sheet_b = workbook_2.sheet_by_index(sheet_idx)

        cat_sheet = new_wb.add_sheet(sheet_a.name)

        for col_idx in range(sheet_a.ncols):
            values_1 = sheet_a.col_values(col_idx, 0, sheet_a.nrows)
            values_2 = sheet_b.col_values(col_idx, 0, sheet_b.nrows)

            cat_col = values_1 + values_2

            for row_idx, value in enumerate(cat_col):
                cat_sheet.write(row_idx, col_idx, value)

    return new_wb


def cat_excel(filename_a, filename_b, filename_result):
    new_wb = cat_workbook(
        xlrd.open_workbook(filename_a),
        xlrd.open_workbook(filename_b)
    )
    new_wb.save(filename_result)


if __name__ == '__main__':
    cat_excel(
        'excel_files/a.xlsx',
        'excel_files/b.xlsx',
        'excel_files/result.xls'
    )

import xlrd
import xlwt

MAX_SHEETS = 1
MAX_COLS = 1
MAX_ROWS = 10

if __name__ == '__main__':
    wb1 = xlrd.open_workbook(filename='excel_files/a.xlsx')
    wb2 = xlrd.open_workbook(filename='excel_files/b.xlsx')

    new_wb = xlwt.Workbook()

    for sheet_idx in range(MAX_SHEETS):
        sheet1 = wb1.sheet_by_index(sheet_idx)
        sheet2 = wb2.sheet_by_index(sheet_idx)

        cat_sheet = new_wb.add_sheet(sheet1.name)

        for col_idx in range(MAX_COLS):
            values_1 = sheet1.col_values(col_idx, 0, MAX_ROWS)
            values_2 = sheet2.col_values(col_idx, 0, MAX_ROWS)

            cat_col = values_1 + values_2

            for row_idx, value in enumerate(cat_col):
                cat_sheet.write(row_idx, col_idx, value)

    new_wb.save('excel_files/result.xls')

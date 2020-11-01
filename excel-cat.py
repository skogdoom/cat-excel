import xlrd
import xlwt
import argparse


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


def cat_excel(filename_result, *filenames):
    new_wb = cat_workbook(*[xlrd.open_workbook(filename) for filename in filenames])
    new_wb.save(filename_result)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Concatenate Excel files.')
    parser.add_argument('files', nargs='+', type=str, help='path and filename of an Excel file to concatenate')
    parser.add_argument('-result', type=str, default='result.xls', help='path and filename of result file to create')
    args = parser.parse_args()

    cat_excel(args.result, *args.files)

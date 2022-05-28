import xlrd
import xlwt
import os

OUTPUT_PATH = 'output'

if __name__ == "__main__":
    files = os.listdir()
    files = [file for file in files if file.endswith('.xls')]

    if not os.path.exists(OUTPUT_PATH):
        os.mkdir(OUTPUT_PATH)

    for file in files:
        print('Processing file: {}'.format(file))
        book = xlrd.open_workbook(file)
        new_book = xlwt.Workbook()
        num_of_sheets = book.nsheets
        for i in range(num_of_sheets):
            sheet = book.sheet_by_index(i)
            print('--> Sheet name: {}'.format(sheet.name))
            new_sheet = new_book.add_sheet(sheet.name)
            n_col = sheet.ncols
            n_row = sheet.nrows
            for row in range(n_row):
                for col in range(n_col):
                    cell_value = sheet.cell_value(rowx=row, colx=col).strip()
                    new_sheet.write(row, col, cell_value)
        
        output_path = os.path.join(OUTPUT_PATH, file)
        new_book.save(output_path)
        print('Saved file: {}'.format(output_path))
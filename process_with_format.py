from multiprocessing.sharedctypes import Value
import os

import xlrd
import xlutils.copy
import xlwt

OUTPUT_PATH = 'output'

def getOutCell(outSheet, colIndex, rowIndex):
        """ HACK: Extract the internal xlwt cell representation. """
        row = outSheet._Worksheet__rows.get(rowIndex)
        if not row: return None

        cell = row._Row__cells.get(colIndex)
        return cell

def setOutCell(outSheet, col, row, value):
    """ Change cell value without changing formatting. """
    previousCell = getOutCell(outSheet, col, row)

    outSheet.write(row, col, value)

    if previousCell:
        newCell = getOutCell(outSheet, col, row)
        if newCell:
            newCell.xf_idx = previousCell.xf_idx

if __name__ == "__main__":
    files = os.listdir()
    files = [file for file in files if file.endswith('.xls')]

    if not os.path.exists(OUTPUT_PATH):
        os.mkdir(OUTPUT_PATH)

    for file in files:
        print('Processing file: {}'.format(file))
        book = xlrd.open_workbook(file, formatting_info=True)
        new_book = xlutils.copy.copy(book)
        num_of_sheets = book.nsheets
        for i in range(num_of_sheets):
            sheet = book.sheet_by_index(i)
            print('--> Sheet name: {}'.format(sheet.name))
            new_sheet = new_book.get_sheet(i)
            n_col = sheet.ncols
            n_row = sheet.nrows
            for row in range(n_row):
                for col in range(n_col):
                    cell_value = sheet.cell_value(rowx=row, colx=col).strip()
                    # new_sheet.write(row, col, cell_value)
                    setOutCell(new_sheet, row, col, cell_value)
        
        output_path = os.path.join(OUTPUT_PATH, file)
        new_book.save(output_path)
        print('Saved file: {}'.format(output_path))



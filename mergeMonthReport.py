import xlwings as xw
import xlrd
app = xw.App(visible=False, add_book=False)
wb = app.books.open(r'H:\temp\all.xls')

bookTemp = xlrd.open_workbook(r'H:\temp\all1.xls')

book = xlrd.open_workbook(r'H:\temp\excels\（发潍坊）附件2：数据统计表（共10个表）.xls')
print("The number of worksheets is {0}".format(book.nsheets))
print("Worksheet name(s): {0}".format(book.sheet_names()))
# sh = book.sheet_by_index(0)

for sheetFrom in book.sheets():
    sheetName = sheetFrom.name
    print(sheetName)
    if sheetName.startswith('表'):
        sheetMid = bookTemp.sheet_by_name(sheetName)
        sheetTo = wb.sheets[sheetName]
        # for row in sheetTo.used_range:
        #     print(row)

        for rx in range(sheetFrom.nrows):
            row = (sheetFrom.row_values(rx))
            if len(row) > 3 and row[2].endswith('农商银行'):
                for i in range(sheetMid.nrows):
                    rowT = (sheetMid.row_values(i))
                    if len(rowT) >= 3 and rowT[2] == row[2]:
                        print(row[0],rx,i)
                        sheetTo.range((i, 3)).value=row[2:len(row)-1]
                        print(row[2:len(row)-1])
                        break

        print(sheetName,'finished')
        # info = sheet.used_range
        # nrows = info.last_cell.row
        # ncols = info.last_cell.column
        #
        # for row in sheet.used_range:
        #     continue
        #     print(row)
        # print(nrows)
        # print(ncols)
        # print(sht.range('A2').value)
wb.save()
wb.close()
app.quit()
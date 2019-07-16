import xlrd
from openpyxl import load_workbook

import json
import os
from shutil import copyfile
from tkinter import *
from tkinter.filedialog import askdirectory
from tkinter.messagebox import *

def initDatas(filePath):
    datas = {}
    bookTemp = xlrd.open_workbook(filePath)
    for sheet in bookTemp.sheets():
        sheetName = sheet.name
        if sheetName.startswith('表'):
            data = {}
            for rx in range(sheet.nrows):
                row = (sheet.row_values(rx))
                if len(row) > 3 and row[2].endswith('农商银行'):
                    data[row[2]] = rx
            datas[sheetName] = data
    # print(datas)
    jsona = json.dumps(datas, ensure_ascii=False)
    print(jsona)
    return datas


def dealExcel(fileFrom):
    datas = {}
    bookTemp = xlrd.open_workbook(fileFrom)
    for sheet in bookTemp.sheets():
        sheetName = sheet.name
        if sheetName.startswith('表'):
            data = []
            for rx in range(sheet.nrows):
                row = (sheet.row_values(rx))
                if len(row) > 3 and row[2].endswith('农商银行'):
                    if sheet.cell_type(rx, 3) == 2 or sheet.cell_type(rx, 5) == 2 or sheet.cell_type(rx, 7) == 2:
                        data.append(row)
            datas[sheetName] = data
    # print(datas)
    # jsona = json.dumps(datas, ensure_ascii=False)
    # print(jsona)
    return datas


def writeExcel(fileTo, datas, datasFromFile):
    wb = load_workbook(fileTo)

    for key in datas.keys():
        value = datas.get(key)
        dataInserts = datasFromFile.get(key)
        sheet = wb.get_sheet_by_name(key)
        for dataInsert in dataInserts:
            org = dataInsert[2]
            rowNumber = value.get(org) + 1
            # print(rowNumber)

            # print(sheet.range((rowNumber, 3)).value)
            for col in range(3, len(dataInsert)):
                theValue = dataInsert[col]
                if isinstance(dataInsert[col], (float, int)):
                    theValue = int(dataInsert[col])
                # theValue = dataInsert[col]
                _ = sheet.cell(column=col + 1, row=rowNumber, value=theValue)
            # sheet.cell(row=rowNumber, column=3).value = 'a'
        # print(dataInserts)

    wb.save(fileTo)
    return datas


def listdir(path):  # 传入存储的list
    list_name = []
    for file in os.listdir(path):
        if (file.endswith('.xls') or file.endswith('.xlsx')) and (not file.startswith('all')):
            file_path = os.path.join(path, file)
            list_name.append(file_path)
    return list_name


def deal(fileFolder):
    datas = initDatas('./template.xlsx')
    fileAll = fileFolder + '\\all1.xlsx'
    copyfile('./template.xlsx', fileAll)
    for file in listdir(fileFolder):
        showlog('deal start'+file+'\n')
        print('deal start', file)
        datasFromFile = dealExcel(file)
        writeExcel(fileAll, datas, datasFromFile)



def doIt():
    if len(thePath) <= 0:
        showinfo('提示', '先选择文件夹')
    else:
        deal(thePath)
        print('deal finished')
        showlog('deal finished')

def selectPath():
    global thePath
    thePath = askdirectory()
    path.set(thePath)

def showlog(log):
    t.insert('end', log)

root = Tk()
global thePath
thePath = ''
path = StringVar()
t = Text(root)
t.grid(row=2,columnspan = 4)
Button(root, text="路径选择", command=selectPath).grid(row=0, column=0)
Entry(root, textvariable=path).grid(row=0, column=1)
Button(root, text="DoIt", command=doIt).grid(row=0, column=3)
root.mainloop()

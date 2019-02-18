#! python3
import os
import re
import shutil
import xlrd
import xlwt
from xlutils.copy import copy

basePath = os.path.join("C:\\", "Users", "zbsz", "Desktop", "库存")
regex = "^.*\.(?:xls)$"
fileList = []
modifyRegex = "^.*[abcdABCD]\.(?:xls)$"
centerStyle = xlwt.XFStyle()
al = xlwt.Alignment()
al.horz = 0x02  # 设置水平居中
al.vert = 0x01  # 设置垂直居中
centerStyle.alignment = al


def move_function():
    for filename in fileList:
        try:
            if not os.path.isdir(os.path.join(basePath, "已处理文件")):
                os.mkdir(os.path.join(basePath, "已处理文件"))
            shutil.move(os.path.abspath(filename), os.path.join(basePath, "已处理文件"))
        except shutil.Error:
            print(filename + " already exists in 已处理文件")


def create():
    for rootName, subFolders, filenames in os.walk(os.path.join(basePath, "下载文件")):
        for filename in filenames:
            if re.match(regex, filename):
                fileList.append(os.path.join(rootName, filename))
                if re.match(modifyRegex, filename):
                    modify = True
                else:
                    modify = False
                if not os.path.isdir(os.path.join(basePath, filename[0:6])):
                    os.mkdir(os.path.join(basePath, filename[0:6]))
                if not os.path.isfile(os.path.join(basePath, filename[0:6], filename[0:6]+".xls")):
                    writeBook = xlwt.Workbook()
                    readBook = xlrd.open_workbook(os.path.join(basePath, "原始文件.xls"))
                    for sheetName in readBook.sheet_names():
                        writeBook.add_sheet(sheetName)
                        if sheetName == "物品编号":
                            for rowIndex in range(readBook.sheet_by_name(sheetName).nrows):
                                for colIndex in range(readBook.sheet_by_name(sheetName).ncols):
                                    writeBook.get_sheet(sheetName).write(rowIndex, colIndex,
                                                                         readBook.sheet_by_name(sheetName).cell(
                                                                             rowIndex, colIndex).value)
                        else:
                            for rowIndex in range(1, readBook.sheet_by_name(sheetName).nrows):
                                for colIndex in range(readBook.sheet_by_name(sheetName).ncols):
                                    writeBook.get_sheet(sheetName).write(rowIndex, colIndex,
                                                                         readBook.sheet_by_name(sheetName).cell(
                                                                             rowIndex, colIndex).value)
                    writeBook.save(os.path.join(basePath, filename[0:6], filename[0:6]+".xls"))
                    copy_function(os.path.join(rootName, filename), os.path.join(basePath, filename[0:6], filename[0:6]
                                                                                 + ".xls"), modify)
                else:
                    copy_function(os.path.join(rootName, filename), os.path.join(basePath, filename[0:6], filename[0:6]
                                                                                 + ".xls"), modify)
                if not os.path.isfile(os.path.join(basePath, filename[0:6], filename[0:8]+".xls")):
                    writeBook = xlwt.Workbook()
                    readBook = xlrd.open_workbook(os.path.join(basePath, "原始文件.xls"))
                    for sheetName in readBook.sheet_names():
                        writeBook.add_sheet(sheetName)
                        if sheetName == "物品编号":
                            for rowIndex in range(readBook.sheet_by_name(sheetName).nrows):
                                for colIndex in range(readBook.sheet_by_name(sheetName).ncols):
                                    writeBook.get_sheet(sheetName).write(rowIndex, colIndex,
                                                                         readBook.sheet_by_name(sheetName).cell(
                                                                             rowIndex, colIndex).value)
                        else:
                            for rowIndex in range(1, readBook.sheet_by_name(sheetName).nrows):
                                for colIndex in range(readBook.sheet_by_name(sheetName).ncols):
                                    writeBook.get_sheet(sheetName).write(rowIndex, colIndex,
                                                                         readBook.sheet_by_name(sheetName).cell(
                                                                             rowIndex, colIndex).value)
                    writeBook.save(os.path.join(basePath, filename[0:6], filename[0:8]+".xls"))
                    copy_function(os.path.join(rootName, filename), os.path.join(basePath, filename[0:6], filename[0:8]
                                                                                 + ".xls"), modify)
                else:
                    copy_function(os.path.join(rootName, filename), os.path.join(basePath, filename[0:6], filename[0:8]
                                                                                 + ".xls"), modify)


def copy_function(pre_filename, suf_filename, modify):
    if modify:
        usingStyle = xlwt.easyxf('font: color-index red')
    else:
        usingStyle = xlwt.easyxf('font: color-index black')
    prebook = xlrd.open_workbook(pre_filename)
    sufbook1 = xlrd.open_workbook(suf_filename)
    sufbook2 = copy(sufbook1)
    for sheetIndex in range(1, len(sufbook1.sheet_names())):
        if sufbook1.sheet_names()[sheetIndex] != "Sheet1":
            sufbook2.get_sheet(sheetIndex).write_merge(0, 0, 0, sufbook1.sheet_by_index(sheetIndex).ncols,
                                                   prebook.sheet_by_index(sheetIndex).cell(0, 0).value, centerStyle )
        for rowIndex in range(2, prebook.sheet_by_index(sheetIndex).nrows):
            sufbook2.get_sheet(sheetIndex).write(sufbook1.sheet_by_index(sheetIndex).nrows+rowIndex-2, 0,
                                                 sufbook1.sheet_by_index(sheetIndex).nrows + rowIndex - 3, usingStyle)
            for colIndex in range(1, prebook.sheet_by_index(sheetIndex).ncols):
                sufbook2.get_sheet(sheetIndex).write(sufbook1.sheet_by_index(sheetIndex).nrows+rowIndex-2, colIndex,
                                                     prebook.sheet_by_index(sheetIndex).cell(
                                                                             rowIndex, colIndex).value, usingStyle)
    sufbook2.save(suf_filename)


try:
    create()
except PermissionError:
    print("文件已被打开，请关闭文件重试")
move_function()




















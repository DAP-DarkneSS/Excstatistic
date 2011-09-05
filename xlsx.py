#!/usr/bin/env python
# -*- coding: utf-8 -*-
#Указывается язык и кодировка.

from openpyxl.reader.excel import *
from openpyxl.workbook import *
from openpyxl.worksheet import *
from cPickle import dumps, loads

#Загружается модуль импорта таблиц Excel *.xlsx, графическая библиотека
#и модуль, содержащий текстовый виджет с полосой прокрутки.

def xlsx(filename, t_sheet, t_column):
    matrix = []
    if len(t_column) == 1:
        column_number = (ord(t_column) - 65)
        if (-1 < column_number < 26):
             sheet_number = int(t_sheet)
             book = load_workbook(filename)
             sheet_names = book.get_sheet_names()
             sheet = book.get_sheet_by_name(name = sheet_names[sheet_number - 1])
             for i in xrange(sheet.get_highest_row()):
                 #cell = sheet.cell(row = i, column = column_number
                 #cell = cell.value
                 #cell = u""
                 cell = sheet.cell(t_column + str(i + 1))
                 #cell = unicode(str(buffer(cell)), "utf-8")
                 #matrix.append(str(loads(dumps(cell.value))))
                 t = dumps(cell.value)
                 t = t.replace("V\\", "\\")
                 t = t.replace("\np1\n.", "")
                 t = unicode(t, "utf-8")
                 matrix.append(t)
        else:
            matrix.append(u"Ошибка: введите букву столбца от A до Z!")
    else:
        matrix.append(u"Ошибка: введите букву столбца от A до Z!")
        


#    matrix.append(u"Ошибка: обработка xlsx не готова!")
    return(matrix)
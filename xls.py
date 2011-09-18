#!/usr/bin/env python
# -*- coding: utf-8 -*-
#Указывается язык и кодировка.

from xlrd import *
from cPickle import dumps, loads
#Загружается модуль импорта таблиц Excel *.xls, графическая библиотека
#и модуль, содержащий текстовый виджет с полосой прокрутки.

def xls(filename, t_sheet, t_column):
    matrix = []
    sheet_number = int(t_sheet)
    if len(t_column) == 1:
        column_number = (ord(t_column) - 65)
        if (-1 < column_number < 26):
            book = open_workbook(filename)
            sheet = book.sheet_by_index(sheet_number - 1)
            for i in xrange(sheet.nrows):
                cell = sheet.cell_value(rowx=i, colx=column_number)
                t = dumps(cell)
                t = t.replace("V\\", "\\")
                t = t.replace("\np1", "")
                t = t.replace("\n.", "")
                t = t.replace("\\", "/\\")
                print(repr(t))
                t = t.encode('utf-8')
                print(repr(t))
                t = t.decode('utf-8')
                print(repr(t))
                if t != "S''":
                    matrix.append(t)
        else:
            matrix.append(u"Ошибка: введите букву столбца от A до Z!")
    else:
        matrix.append(u"Ошибка: введите букву столбца от A до Z!")


        
    return(matrix)

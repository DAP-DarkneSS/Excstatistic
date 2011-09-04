#!/usr/bin/env python
# -*- coding: utf-8 -*-
#Указывается язык и кодировка.

from xlrd import *
#Загружается модуль импорта таблиц Excel *.xls, графическая библиотека
#и модуль, содержащий текстовый виджет с полосой прокрутки.

def xls(filename, t_sheet, t_column):
    matrix = []
    book = open_workbook(filename)
    sheet_nunmber = int(t_sheet)
    if len(t_column) == 1:
        column_nunmber = (ord(t_column) - 65)
        if (-1 < column_nunmber < 26):
            sheet = book.sheet_by_index(sheet_nunmber - 1)

            for i in xrange(sheet.nrows):
                cell = sheet.cell_value(rowx=i, colx=column_nunmber)
                matrix.append(cell)
        else:
            matrix.append(u"Ошибка: введите букву столбца от A до Z!")
    else:
        matrix.append(u"Ошибка: введите букву столбца от A до Z!")


        
    return(matrix)
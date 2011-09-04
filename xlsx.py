#!/usr/bin/env python
# -*- coding: utf-8 -*-
#Указывается язык и кодировка.

from openpyxl.reader.excel import *
from openpyxl.workbook import *
from openpyxl.worksheet import *

#Загружается модуль импорта таблиц Excel *.xlsx, графическая библиотека
#и модуль, содержащий текстовый виджет с полосой прокрутки.

def xlsx(filename, t_sheet, t_column):
    matrix = []
    try:
        sheet_number = int(t_sheet)
        book = load_workbook(filename)
        sheet_names = book.get_sheet_names()
        sheet = get_sheet_by_name(name = sheet_names[sheet_number - 1])
        for i in xrange(sheet.get_highest_row()):
            cell = sheet.cell(row = i, column = t_column)
            matrix.append(cell)
    except KeyError:
        matrix.append(u"Ошибка: попробуйте сохранить файл в Excel 2003 - .xls!")
    


#    matrix.append(u"Ошибка: обработка xlsx не готова!")
    return(matrix)
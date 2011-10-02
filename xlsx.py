#!/usr/bin/env python
# -*- coding: utf-8 -*-
#Указывается язык и кодировка.

from openpyxl.reader.excel import *
from openpyxl.workbook import *
from openpyxl.worksheet import *
#Загружаются модули работы с таблицами Excel *.xlsx.

def xlsx(filename, t_sheet, t_column):
#Функция принимает строковые значения имени файла, номера листа и колонки.

    matrix = [""]
#Создаётся ряд с пустой строкой (первое значение).

    sheet_number = int(t_sheet)
#Номер листа преобразуется в целое число.

    book = load_workbook(filename)
#Открывается файл.

    sheet_names = book.get_sheet_names()
    sheet = book.get_sheet_by_name(name = sheet_names[sheet_number - 1])
#Запрашивается список имён листов и открывается нужный лист.

    for i in xrange(sheet.get_highest_row()):
        cell = sheet.cell(t_column + str(i + 1))
        matrix.append(cell.value)
#Значения считываются из колонки и записываются в ряд.

    return(matrix)
#Функция возвращает ряд.
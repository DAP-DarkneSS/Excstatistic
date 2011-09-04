#!/usr/bin/env python
# -*- coding: utf-8 -*-

#This code is released under
#GNU LESSER GENERAL PUBLIC LICENSE
#http://www.gnu.org/copyleft/lesser.html
#Author: DarkneSS at https://github.com/DAP-DarkneSS

from xlrd import *
#Загружается модуль импорта таблиц Excel *.xls, графическая библиотека
#и модуль, содержащий текстовый виджет с полосой прокрутки.

def excstatistic(filename, t_sheet, t_column):
    text=u""
    book = open_workbook(filename, encoding_override="utf-8")
    sheet_nunmber = int(t_sheet)
    if len(t_column) == 1:
        column_nunmber = (ord(t_column) - 65)
        if not (-1 < column_nunmber < 26):
            text =  u"Ошибка: введите букву столбца от A до Z!"
    else:
        text =  u"Ошибка: введите букву столбца от A до Z!"

    sheet = book.sheet_by_index(sheet_nunmber - 1)

    name = []
    name_count = []

    for i in xrange(sheet.nrows):
        cell = sheet.cell_value(rowx=i, colx=column_nunmber)
        
        if cell != "":
            j = -1
            try:
                j = name.index(cell)
            except ValueError:
                pass
            if (-1 < j < len(name)):
                name_count[j] += 1
            else:
                name.append(cell)
                name_count.append(1)

    for i in xrange(len(name)):
        text += str(name_count[i])
        text += u" @ "
        text += name[i]
        text += u"\n"
        
    return(text)
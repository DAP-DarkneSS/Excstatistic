#!/usr/bin/env python
# -*- coding: utf-8 -*-
#Указывается язык и кодировка.

from xls import xls
from xlsx import xlsx
#Загружается модуль импорта таблиц Excel *.xls, графическая библиотека
#и модуль, содержащий текстовый виджет с полосой прокрутки.

def excstatistic(filename, t_sheet, t_column):
    
    text = u""
    name = []
    name_count = []
    matrix = []
    
    if filename.endswith(".xlsx"):
        matrix = xlsx(filename, t_sheet, t_column)
    elif filename.endswith(".xls"):
        matrix = xls(filename, t_sheet, t_column)
    else:
        text =  u"Ошибка: неправильный файл!"
    try:
        if not (matrix[0]).startswith(u"Ошибка: "):
            for i in xrange(len(matrix)):
                if matrix[i] != "":
                    j = -1
                    try:
                        j = name.index(matrix[i])
                    except ValueError:
                        pass
                    if (-1 < j < len(name)):
                        name_count[j] += 1
                    else:
                        name.append(matrix[i])
                        name_count.append(1)
            for i in xrange(len(name)):
                text += str(name_count[i])
                text += u" @ "
                text += name[i]
                text += u"\n"
        else:
            text = matrix[0]
    except AttributeError:
        pass        
    return(text)
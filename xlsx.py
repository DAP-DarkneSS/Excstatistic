#!/usr/bin/env python
# -*- coding: utf-8 -*-
#Указывается язык и кодировка.

from xlrd import *
#Загружается модуль импорта таблиц Excel *.xls, графическая библиотека
#и модуль, содержащий текстовый виджет с полосой прокрутки.

def xlsx(filename, t_sheet, t_column):
    matrix = []
    matrix.append(u"Ошибка: обработка xlsx не готова!")
    return(matrix)
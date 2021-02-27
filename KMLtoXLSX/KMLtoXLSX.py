#!/usr/bin/env python
# -*- coding: UTF-8 -*-
import openpyxl
from os.path import join, abspath
from bs4 import BeautifulSoup


def print_description_text():
    '''
        =================================================================
        Parsing <name> from .KML to .XLSX ver. 1.0.2

        Программа сохраняет имена точек (тег:name) из файла «Points.kml».
        В файл «Points.xlsx» (нужен пустой файл таким именем)

        tha7n7@gmail.com freeware ©2021
        =================================================================
    '''
    pass


def print_version_text():
    '''
        ============================================
        Parsing <name> from .KML to .XLSX ver. 1.0.2
        tha7n7@gmail.com freeware ©2021
        ============================================
    '''
    pass


def open_xlsx(name):
    xlsx_file = join('.', name)
    xlsx_file = abspath(xlsx_file)
    try:
        return openpyxl.load_workbook(filename=xlsx_file, read_only=False, data_only=True)
    except FileNotFoundError:
        print(print_description_text.__doc__)
        print("{} не найден".format(name))
        print()
        input("Press any key to continue...")
        quit()


def main():
    points_list = []
    kml_file = join('.', 'Points.kml')
    with open(kml_file, 'r', encoding="utf-8") as f:
        soup = BeautifulSoup(f, 'xml')
        for nm in soup.find_all('name'):
            points_list.append(nm.string)

    if len(points_list) > 0:
        wb_points = open_xlsx('Points.xlsx')
        sh = wb_points.active
        sh.delete_cols(1)
        for i in range(0, len(points_list)):
            sh.cell(row=1 + i, column=1).value = points_list[i]
        print("Сохранено (" + str(len(points_list)) + ")")
        wb_points.save('Points.xlsx')
        wb_points.close()
    else:
        print("Имена точек не найдены")
        print()
    print(print_version_text.__doc__)


if __name__ == '__main__':
    print()
    main()
    print()
    input("Press any key to continue...")

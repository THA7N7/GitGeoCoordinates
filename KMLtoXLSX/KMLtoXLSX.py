#!/usr/bin/env python
# -*- coding: UTF-8 -*-
import sys
import openpyxl
from os.path import join, abspath
from bs4 import BeautifulSoup


def description_text():
    description = [
        ' .KML to .XLSX ver. 1.0.1',
        ' Программа сохраняет имена точек (тег:name) из файла «Points.kml».',
        ' В файл «Points.xlsx» (нужен пустой файл таким именем)',
        ' tha7n7@gmail.com freeware ©2021',
    ]
    return description


def open_points(name):
    xls_file = join('.', name)
    xls_file = abspath(xls_file)
    try:
        return openpyxl.load_workbook(filename=xls_file, read_only=False, data_only=True)
    except:
        print("Points.xlsx не найден")
        print()
        input("Press any key to continue...")
        quit()


def close_points(msg):
    print(msg)
    input("Press any key to continue...")
    wb_points.close()


if __name__ == '__main__':
    description = description_text()
    print(description[0])
    print()
    print(description[1])
    print(description[2])
    print()
    print(description[3])
    print()

    points_list = []
    kml_file = join('.', 'Points.kml')
    with open(kml_file, 'r', encoding="utf-8") as f:
        soup = BeautifulSoup(f, 'xml')
        for nm in soup.find_all('name'):
            points_list.append(nm.string)

    if len(points_list) > 0:
        wb_points = open_points('Points.xlsx')
        i = 0
        sh = wb_points.active
        sh.delete_cols(1)
        for i in range(len(points_list)):
            sh.cell(row=1 + i, column=1).value = points_list[i]
        print("Сохранено (" + str(len(points_list)) +")")
    else:
        print("Имена точек не найдены")

    wb_points.save('Points.xlsx')
    close_points('')

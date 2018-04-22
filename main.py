# -*- coding: utf-8 -*-
import xlrd
from math import log2


def read_excel():
    '''读取excel返回列表D，
    D[0]:RID
    D[1]:age
    D[2]:income
    D[3]:student
    D[4]:credit_rating
    D[5]:class_buys_computer'''
    ExcelFile = xlrd.open_workbook('Book1.xlsx')
    sheet = ExcelFile.sheet_by_name('Book1')
    print('{}已读取 {}行 {}列'.format(sheet.name, sheet.nrows, sheet.ncols))
    D = []
    for col in range(0, sheet.ncols):
        D.append(sheet.col_values(col))
    return D


def count_age(list_age):
    '''统计age中youth、middle_age、senior次数，返回a、b、c'''
    a = list_age.count('youth')
    b = list_age.count('middle_age')
    c = list_age.count('senior')
    return a, b, c


def count_computer(list_class_computer):
    k1 = list_class_computer.count('yes')
    k2 = list_class_computer.count('no')
    return k1, k2


def self_count(age, class_computer, keyword):
    if age == keyword:
        if class_computer == 'yes':
            return 1, 0
        else:
            return 0, 1
    return 0, 0


def count_age_computer(list_age, list_class_computer):
    a1 = 0
    a2 = 0
    b1 = 0
    b2 = 0
    c1 = 0
    c2 = 0
    for i in range(len(list_age)):
        int_yes, int_no = self_count(list_age[i], list_class_computer[i], 'youth')
        a1 += int_yes
        a2 += int_no
        int_yes, int_no = self_count(list_age[i], list_class_computer[i], 'middle_age')
        b1 += int_yes
        b2 += int_no
        int_yes, int_no = self_count(list_age[i], list_class_computer[i], 'senior')
        c1 += int_yes
        c2 += int_no
    return a1, a2, b1, b2, c1, c2


def cal_1(a, b, c, kw):
    tmp = a + b + c
    if kw == 'a':
        return a / tmp
    elif kw == 'b':
        return b / tmp
    else:
        return c / tmp


def cal_2(x1, x2):
    if x1==0:
        return - (x2 / (x1 + x2)) * log2(x2 / (x1 + x2))
    elif x2==0:
        return -(x1 / (x1 + x2)) * log2(x1 / (x1 + x2))
    else:
        return -(x1 / (x1 + x2)) * log2(x1 / (x1 + x2)) - (x2 / (x1 + x2)) * log2(x2 / (x1 + x2))


if __name__ == '__main__':
    D = read_excel()
    a, b, c = count_age(D[1])
    k1, k2 = count_computer(D[5])
    a1, a2, b1, b2, c1, c2 = count_age_computer(D[1], D[5])
    c1=3
    c2=2
    info=cal_2(k1,k2)
    infoage = cal_1(a, b, c, 'a') * cal_2(a1, a2) + \
              cal_1(a, b, c, 'b') * cal_2(b1, b2) + \
              cal_1(a, b, c, 'c') * cal_2(c1, c2)
    print('Info(D)={}\nInfoage(D)={}'.format(info,infoage))

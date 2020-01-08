import matplotlib.pyplot as plt
from matplotlib.widgets import Slider, Button, RadioButtons
import xlrd
import math
import statistics as sta
import numpy as np


def num(s):
    try:
        return int(s)
    except ValueError:
        return float(s)


def getMoyenne(col_array):
    innerMoy = 0
    for value in col_array:
        if value == "":
            break
        innerMoy = num(value) + innerMoy
    return round(innerMoy / len(col_array), 1)


def getMin(col_array):
    minimum = math.inf
    for value in col_array:
        if value != '' and minimum > value:
            minimum = value
    return minimum


def getMax(col_array):
    maximum = -math.inf
    for value in col_array:
        if value != '' and maximum < value:
            maximum = value
    return maximum


def print_month_info(col, data):
    month_lst = ['Janvier', 'Fevrier', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet',
                 'Août', 'Septembre', 'Octobre', 'Novembre', 'Decembre']
    return month_lst[col]+" : "+str(data)


def get_ecart_type(moyenne, col_array):
    variance = 0
    for value in col_array:
        if value == "":
            break
        variance = variance + ((num(value) - moyenne)*(num(value) - moyenne))
    return round(math.sqrt(variance/len(col_array)), 1)



def exec_exercice_1():
    File = xlrd.open_workbook('Climat.xlsx')
    sheet = File.sheet_by_name('Feuil2')
    year_min_values = []
    year_max_values = []
    month_value = []
    years_values = []

    for i in range(sheet.ncols):
        moy = getMoyenne(sheet.col_values(i))
        min = getMin(sheet.col_values(i))
        max = getMax(sheet.col_values(i))
        year_min_values.append(min)
        year_max_values.append(max)
        for h in sheet.col_values(i):
            if h == "":
                break
            month_value.append(num(h))
        print("====================")
        print("Moyenne: " + print_month_info(i, moy))
        print("Minimum: " + print_month_info(i, min))
        print("Maximum: " + print_month_info(i, max))
        print("Ecart type: " + str(get_ecart_type(moy, sheet.col_values(i))))
        print("--------------------------------------")
        plt.figure(print_month_info(i, ""))
        plt.ylabel("Temperature °C")
        plt.xlabel("Jour")
        plt.plot(month_value)
        month_value = []
        for j in sheet.col_values(i):
            if j == "":
                break
            years_values.append(num(j))

    print("====================")
    print("Minimum for year : " + str(getMin(year_min_values)))
    print("Maximum for year : " + str(getMax(year_max_values)))

    fig = plt.figure()
    ax = fig.add_subplot(111)
    fig.subplots_adjust(left=0.25, bottom=0.25)
    min0 = 0
    max0 = 10

    plt.plot(years_values)
    # most examples here return something iterable

    plt.xlim([1, 31])  # initial limits

    axmin = fig.add_axes([0.25, 0.1, 0.65, 0.03])

    smin = Slider(axmin, 'Months', 1, 12, valinit=min0, valstep=1)

    def update(val):
        val = num(val)
        if val == 1:
            plt.xlim([1, 31])
        elif val == 2:
            plt.xlim([32, 59])
        elif val == 3:
            plt.xlim([60, 90])
        elif val == 4:
            plt.xlim([91, 120])
        elif val == 5:
            plt.xlim([121, 151])
        elif val == 6:
            plt.xlim([152, 181])
        elif val == 7:
            plt.xlim([182, 212])
        elif val == 8:
            plt.xlim([213, 243])
        elif val == 9:
            plt.xlim([244, 273])
        elif val == 10:
            plt.xlim([274, 304])
        elif val == 11:
            plt.xlim([305, 334])
        elif val == 12:
            plt.xlim([335, 365])

    plt.subplot(111)
    smin.on_changed(update)

    plt.show()


exec_exercice_1()





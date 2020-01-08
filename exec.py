import matplotlib.pyplot as plt
import xlrd
import math
import statistics as sta

File = xlrd.open_workbook('Climat.xlsx')
sheet = File.sheet_by_name('Feuil2')


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
    month_lst = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
                 'August', 'September', 'October', 'November', 'December']
    return month_lst[col]+" : "+str(data)


def get_ecart_type(moyenne, col_array):
    variance = 0
    for value in col_array:
        if value == "":
            break
        variance = variance + ((num(value) - moyenne)*(num(value) - moyenne))

    return round(math.sqrt(variance/len(col_array)), 1)


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
    plt.plot(month_value)
    month_value = []
    for j in sheet.col_values(i):
        if j == "":
            break
        years_values.append(num(j))

print("====================")
print("Minimum for year : " + str(getMin(year_min_values)))
print("Maximum for year : " + str(getMax(year_max_values)))
plt.figure("Years")
plt.plot(years_values)
plt.ylabel("Temperature °C")
plt.xlabel("Jour")
plt.show()








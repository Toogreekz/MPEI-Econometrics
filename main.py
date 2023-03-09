import openpyxl
import math


def column(array):
    array.append(sum(array))
    array.append(sum(array[:len(array) - 1]) / (len(array) - 1))
    return array


wb = openpyxl.Workbook()
ws = wb.active
ws.title = "EconometricsTable"

N = int(input("Введи длину столбца х и у: "))
x = column([float(input()) for i in range(N)])
y = column([float(input()) for i in range(N)])
Xlog = column([math.log10(x[i]) for i in range(N)])
Ylog = column([math.log10(y[i]) for i in range(N)])
xy = column([x[i] * y[i] for i in range(N)])
x_2 = column([x[i] ** 2 for i in range(N)])
x_minus_x_sr = column([(x[i] - x[-1]) ** 2 for i in range(N)])
y_minus_y_sr = column([(y[i] - y[-1]) ** 2 for i in range(N)])
b = (xy[-1] - x[-1] * y[-1]) / (x_2[-1] - x[-1] * x[-1])
a = y[-1] - b * x[-1]
y_r = column([a + b * x[i] for i in range(N)])
y_minus_y_r = column([y[i] - y_r[i] for i in range(N)])
y_minus_y_r_2 = column([(y[i] - y_r[i]) ** 2 for i in range(N)])
A = column([abs(y_minus_y_r[i] / y[i]) * 100 for i in range(N)])

array = [x, y, Xlog, Ylog, xy, x_2, x_minus_x_sr, y_minus_y_sr, y_r, y_minus_y_r, y_minus_y_r_2, A]
row = ["x", "y", "X", "Y", "xy", "x^2", "(x-СРЕД(х))^2", "(y-СРЕД(y))^2", "ŷ", "y-ŷ",
         "(y-ŷ)^2", "A"]
s = "ABCDEFGHIJKL"
for i in s:
    ws[f"{i}1"] = row[s.index(i)]
for i in s:
    for elem in array[s.index(i)]:
        ws[f"{i}{array[s.index(i)].index(elem) + 2}"] = elem

wb.save(filename="Econometrics.xlsx")
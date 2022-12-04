import openpyxl

wb_2018 = openpyxl.load_workbook('ALBRK2018.xlsx')
wb_2019 = openpyxl.load_workbook('ALBRK2019.xlsx')
sheet_2018 = wb_2018.get_sheet_by_name('Worksheet')
sheet_2019 = wb_2019.get_sheet_by_name('Worksheet')


def creatingList(col1, col2):
    c1List = []
    c2List = []

    for cc1, cc2 in zip(col1, col2):
        c1List.append(str(cc1.value))
        c2List.append(str(cc2.value))
    return c1List, c2List


P, H = creatingList(sheet_2018['A'], sheet_2018['C'])
c2ListLength = len(H) - 1


def leftToRight():
    maxValue1 = H[1]
    index = 0
    for i in range(1, c2ListLength):
        if H[i + 1] > maxValue1:
            maxValue1 = H[i+1]
            maxValueIndex = i + 1
            index = i + 1
        minValue1 = H[index]
        for j in range(index, 0, -1):
            if H[j-1] < minValue1:
                minValue1 = H[j-1]
                minValueIndex = j - 1
    maxValue1 = maxValue1.replace(",", ".")
    minValue1 = minValue1.replace(",", ".")
    # Maxiumum Profit for left to right
    dif1 = (float(maxValue1)) - (float(minValue1))
    print(
        f"If you buy on {P[minValueIndex]} and if you sell on {P[maxValueIndex]} those dates, you will get maximum profit rate which is: %.3f" % dif1)
    return dif1


def rightToLeft():
    minValue2 = H[1]
    index = 0
    for i in range(0, c2ListLength):
        if minValue2 > H[i+1]:
            minValue2 = H[i+1]
            index = i + 1
            minValueIndex2 = i + 1
        maxValue2 = H[index]
        for j in range(index, c2ListLength):
            if H[j+1] > maxValue2:
                maxValue2 = H[j+1]
                maxValueIndex2 = j + 1
    maxValue2 = maxValue2.replace(",", ".")
    minValue2 = minValue2.replace(",", ".")
    # Maxiumum Profit for right to left
    dif2 = (float(maxValue2)) - (float(minValue2))
    print(
        f"If you buy on {P[minValueIndex2]} and if you sell on  {P[maxValueIndex2]} those dates, you will get maximum profit rate which is: %.3f" % dif2)
    return dif2


def compareSolution(dif1, dif2):
    if (dif1 > dif2):
        print("Max profit is satisfied with the leftToRight function.")
    elif (dif2 > dif1):
        print("Max profit is satisfied with the rightToLeft function.")
    else:
        print("Max profit can satisfied with both function.")


compareSolution(leftToRight(), rightToLeft())

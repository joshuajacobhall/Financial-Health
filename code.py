"""
NOTE: pandas will not load first row

Method:

Sort by years, before hand

1. Setup array
2. Calc stdev and mean of expenditure per pupil
3. Create sheet of not outlying schools
4. Calc stdev and mean of each metric
5. Create score according to distance from mean
6. Create weights according to avg proportion of expenditure
7. Final score = sum of weight*score
8. Create grade from final score


Split by school region?
"""

from numpy import mean, std, asarray
from pandas import read_excel
import csv

from mpmath import *
mp.dps = 50


def removeZeros(columns):
    global height
    global sheet
    sheet = sheet.tolist()

    print('\nZero Pupil/Expenditure Errors:')
    for column in columns:
        for row in range(height):
            try:
                if sheet[row][column] == 0:
                    print(sheet[row][COL_NAME])
                    del sheet[row]
                    height -= 1
            except:
                print('Error on:', row, column)

    print('\nEmpty Column Errors:')
    for column in range(COL_PUPIL, len(sheet[0])):
        try:
            if sum([sheet[row][column] for row in range(height)]) == 0:
                for row in range(height):
                    del sheet[row][column]
        except:
            print('Error on:', row, column)

    sheet = asarray(sheet)


def outliers(column):
    global height
    global sheet

    outliers = []
    data = [mpf(sheet[row][column]) for row in range(height)]

    upper_bound = mean(data) + 3*std(data, ddof=1)
    lower_bound = mean(data) - 3*std(data, ddof=1)

    for row in range(height):
        value = mpf(sheet[row][column])
        if value > upper_bound or value < lower_bound:
            outliers.append(row)

    sheet = sheet.tolist()
    print('\nOutliers:')
    for row in outliers:
        print(sheet[row][COL_NAME])
        del sheet[row]
        height -= 1
    sheet = asarray(sheet)


def weights(start, stop):
    return [mean([mpf(sheet[row][column]) / mpf(sheet[row][COL_EXPENDITURE])
            for row in range(height)]) for column in range(start, stop)]


def perPupil(start, stop):
    global sheet
    for column in range(start, stop):
        for row in range(height):
                sheet[row][column] = mpf(sheet[row][column]) / mpf(sheet[row][COL_PUPIL])


def metrics(start, stop):
    return [[std([mpf(sheet[row][column]) for row in range(height)], ddof=1),
            mean([mpf(sheet[row][column]) for row in range(height)])]
            for column in range(start, stop)]


def scores(stats, start, stop):
    bands = [[stats[column][1]+stats[column][0],
             stats[column][1],
             stats[column][1]-stats[column][0]]
             for column in range(stop-start)]

    scores = [[4 if mpf(sheet[row][column]) >= bands[column-start][0] else
              3 if bands[column-start][1] <= mpf(sheet[row][column]) < bands[column-start][0] else
              2 if bands[column-start][2] <= mpf(sheet[row][column]) < bands[column-start][1] else
              1 if mpf(sheet[row][column]) < bands[column-start][2] else -1
              for column in range(start, stop)] for row in range(height)]

    return scores


def grade(scores, weights, start, stop):
    final_score = [sum([scores[row][column-start]*weights[column-start]
                   for column in range(start, stop)]) for row in range(height)]

    avg = mean(final_score)
    stdev = std(final_score, ddof=1)
    bands = [avg, avg-stdev, avg-2*stdev]

    grades = ['A' if value >= bands[0] else
              'B' if bands[1] <= value < bands[0] else
              'C' if bands[2] <= value < bands[1] else
              'D' if value <= bands[1] else -1 for value in final_score]

    return final_score
    return grades


def save(grades):
    output = zip([sheet[row][COL_NAME] for row in range(height)], grades)

    with open(OUTPUT_FILE, 'w') as file:
        writer = csv.writer(file)
        writer.writerows(list(output))


INPUT_FILE = 'data/input.xlsx'
OUTPUT_FILE = 'data/output.csv'

COL_NAME = 2
COL_PUPIL = 4
COL_EXPENDITURE = 6

COL_START = 14
COL_STOP = 20

if __name__ == "__main__":
    COL_NAME -= 1
    COL_PUPIL -= 1
    COL_EXPENDITURE -= 1
    COL_START -= 1
    COL_STOP -= 1

    sheet = read_excel(INPUT_FILE).as_matrix()
    height = len(sheet)
    height = 1000

    removeZeros([COL_PUPIL, COL_EXPENDITURE])
    outliers(COL_EXPENDITURE)

    weights = weights(COL_START, COL_STOP+1)

    perPupil(COL_START, COL_STOP+1)
    stats = metrics(COL_START, COL_STOP+1)
    scores = scores(stats, COL_START, COL_STOP+1)

    grades = grade(scores, weights, COL_START, COL_STOP+1)
    save(grades)

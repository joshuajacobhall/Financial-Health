from numpy import mean, std, asarray
from pandas import read_excel
import csv

from mpmath import *  # likely inneficient - will fix!
mp.dps = 50


def setup():
    global COL_NAME, COL_PUPIL, COL_EXPENDITURE, COL_START, COL_STOP

    COL_NAME -= 1  # subtract 1, due to nature of arrays
    COL_PUPIL -= 1
    COL_EXPENDITURE -= 1
    COL_START -= 1
    COL_STOP -= 1


def removeZeros(columns):  # removes "problem" columns (eg zero division error)
    global height, sheet
    sheet = sheet.tolist()  # converts sheet object to list - easier to modify

    print('\nZero Pupil/Expenditure Errors:')
    for column in columns:
        for row in range(height):
            try:
                if sheet[row][column] == 0:  # searches for cells equal to 0
                    print(sheet[row][COL_NAME])
                    del sheet[row]
                    height -= 1  # change max row variable to stop index error
            except:
                print('Error on:', row, column)  # catch unknown errors

    print('\nEmpty Column Errors:')
    for column in range(COL_PUPIL, len(sheet[0])):
        try:
            if sum([sheet[row][column] for row in range(height)]) == 0:
                # checks if a column is full of remove
                # it *could* that this will return true for a non-zero column
                # but unlikely and much more efficient that alternative method
                for row in range(height):
                    del sheet[row][column]  # remove all rows in problem column
        except:
            print('Error on:', row, column)

    sheet = asarray(sheet)  # convert back to data type used elsewhere in code


def outliers(column):  # finds schools with anomalous data in a given column
    global height, sheet

    outliers = []
    data = [mpf(sheet[row][column]) for row in range(height)]
    # creates list of rows in a given column
    # mpf is a datatype from mpmath with high decimal precision

    upper_bound = mean(data) + 3*std(data, ddof=1)
    lower_bound = mean(data) - 3*std(data, ddof=1)
    # we claim that data 3 stdevs away from the mean is anomalous
    # ddof=1 used because we need sample variation, not population

    for row in range(height):
        value = mpf(sheet[row][column])
        if value > upper_bound or value < lower_bound:
            outliers.append(row)  # creates a list of outlier schools

    sheet = sheet.tolist()  # converts to a modifiable datatype
    print('\nOutliers:')
    for row in outliers:
        print(sheet[row][COL_NAME])
        del sheet[row]  # removes outlier schools - possible an issue!
        height -= 1
    sheet = asarray(sheet)  # converts back to a original datatype


def weights(start, stop):  # calculates "importance" of each metric
    return [mean([mpf(sheet[row][column]) / mpf(sheet[row][COL_EXPENDITURE])
            for row in range(height)]) for column in range(start, stop)]
    # returns mean of every school's proportion of expendature in each metric


def perPupil(start, stop):  # divides metric columns by the number of pupils
    global sheet

    for column in range(start, stop):
        for row in range(height):
                value = mpf(sheet[row][column]) / mpf(sheet[row][COL_PUPIL])
                # divide each value by the number of pupils
                sheet[row][column] = value


def metrics(start, stop):  # calculates stdev and mean of every metric
    return [[std([mpf(sheet[row][column]) for row in range(height)], ddof=1),
            mean([mpf(sheet[row][column]) for row in range(height)])]
            for column in range(start, stop)]
    # returns an array [stdev, mean] for every column of metrics


def scores(stats, start, stop):  # calculates a raw score for each school
    bands = [[stats[column][1]+stats[column][0],  # mean + stdev
             stats[column][1],  # mean
             stats[column][1]-stats[column][0]]  # mean - stdev
             for column in range(stop-start)]
    # calculates bands of for every metric

    scores = [[4 if mpf(sheet[row][column]) >= bands[column-start][0] else
              3 if bands[column-start][1] <= mpf(sheet[row][column]) < bands[column-start][0] else
              2 if bands[column-start][2] <= mpf(sheet[row][column]) < bands[column-start][1] else
              1 if mpf(sheet[row][column]) < bands[column-start][2] else -1
              for column in range(start, stop)] for row in range(height)]
    # if value above top band, max score
    # if value if between top and second band, next max score
    # etc

    return scores


def grade(scores, weights, start, stop):  # creates a final grade for schools
    final_score = [sum([scores[row][column-start]*weights[column-start]
                   for column in range(start, stop)]) for row in range(height)]
    # school's score equals the sum of metric value * weight of metric

    avg = mean(final_score)
    stdev = std(final_score, ddof=1)
    bands = [avg, avg-stdev, avg-2*stdev]
    # creates 3 bands for grade scores

    grades = ['A' if value >= bands[0] else
              'B' if bands[1] <= value < bands[0] else
              'C' if bands[2] <= value < bands[1] else
              'D' if value <= bands[1] else -1 for value in final_score]
    # if above mean, grade is A
    # if above mean - stdev, grade is B
    # etc

    return grades


def save(grades):  # outputs a file of schools and corresponding grades
    output = zip([sheet[row][COL_NAME] for row in range(height)], grades)
    # creates array of names and grades

    with open(OUTPUT_FILE, 'w') as file:  # write information to a cvs file
        writer = csv.writer(file)
        writer.writerows(list(output))


#
# change variables below as appropriate
#


INPUT_FILE = 'data/input.xlsx'
OUTPUT_FILE = 'data/output.csv'

COL_NAME = 2
COL_PUPIL = 4
COL_EXPENDITURE = 6

COL_START = 14
COL_STOP = 20


if __name__ == "__main__":
    setup()

    sheet = read_excel(INPUT_FILE).as_matrix()
    height = len(sheet)
    height = 1000  # fix

    removeZeros([COL_PUPIL, COL_EXPENDITURE])
    outliers(COL_EXPENDITURE)

    weights = weights(COL_START, COL_STOP+1)

    perPupil(COL_START, COL_STOP+1)
    stats = metrics(COL_START, COL_STOP+1)
    scores = scores(stats, COL_START, COL_STOP+1)

    grades = grade(scores, weights, COL_START, COL_STOP+1)
    save(grades)

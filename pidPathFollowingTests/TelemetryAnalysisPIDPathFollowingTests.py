"""
This script takes the telemetry data from the various-radius straight-line PID path-following
tests, calculates the error at every time interval, then exports the data to an excel sheet.

FINISHED: DO NOT RUN THIS SCRIPT
"""


import sys
import math
import xlwt
from xlwt import Workbook
import matplotlib.pyplot as plt
sys.path.append('../')
from MotionProfiles import *

tests = [6, 8, 12, 16, 24]

all_data = []
all_data_trad = []
from _6in_t import *
all_data_trad.append(data)
# data_trad = data
from _6in import *
all_data.append(data)

from _8in_t import *
all_data_trad.append(data)
from _8in import *
all_data.append(data)

from _12in_t import *
all_data_trad.append(data)
from _12in import *
all_data.append(data)

from _16in_t import *
all_data_trad.append(data)
from _16in import *
all_data.append(data)

from _24in_t import *
all_data_trad.append(data)
from _24in import *
all_data.append(data)

def calcError(x, y, target):
    """
    Returns a list containing the error compared to the target endpoint at every time interval

    Params:
     - x: list of x values
     - y: list of y values
     - target: the y-distance of the target endpoint
    """
    errors = []
    for i in range(len(x)):
        xDiff = abs(x[i])
        yDiff = abs(target-y[i])
        diff = math.sqrt(xDiff**2 + yDiff**2)
        errors.append(diff)
    return errors


wb = Workbook()
styleBold = xlwt.easyxf('font: bold 1')
style = xlwt.easyxf('font: bold off')

counter = 0
for n in tests:
    sheet = wb.add_sheet('Telemetry ' + str(n) + 'in. Look-Ahead')

    #row, column
    sheet.write(0, 0, 'PI(t)D(t) Time', styleBold)
    sheet.write(0, 1, 'PI(t)D(t) Error', styleBold)
    sheet.write(0, 2, 'PI(t)D(t) y', styleBold)
    sheet.write(0, 3, 'PI(t)D(t) v', styleBold)
    sheet.write(0, 5, 'PID Time', styleBold)
    sheet.write(0, 6, 'PID Error', styleBold)
    sheet.write(0, 7, 'PID y', styleBold)
    sheet.write(0, 8, 'PID v', styleBold)

    #Get rid of the motion profile
    data = all_data[counter].split('*')[1]
    data_trad = all_data_trad[counter].split('*')[1]

    #Delete whitespace and alphabetical letters
    data = data.replace("\n", "").replace("\r", "").replace(" ", "").replace("angV", "").replace("angA", "").replace("t", "").replace("x", "").replace("y", "").replace("h", "").replace("v", "").replace("a", "")
    data_trad = data_trad.replace("\n", "").replace("\r", "").replace(" ", "").replace("angV", "").replace("angA", "").replace("t", "").replace("x", "").replace("y", "").replace("h", "").replace("v", "").replace("a", "")

    #Split data in a list with t, x, y, h, etc. separated
    data = data.split(':')
    data_trad = data_trad.split(':')

    #Split t, x, y, h, etc. into lists of individual doubles
    t = [float(i) for i in data[1].split(",")[:-1]]             #Removing the last item of the list because it is an empty string
    t_trad = [float(i) for i in data_trad[1].split(",")[:-1]]   #Removing the last item of the list because it is an empty string
    x = [float(i) for i in data[2].split(",")[:-1]]             #Removing the last item of the list because it is an empty string
    x_trad = [float(i) for i in data_trad[2].split(",")[:-1]]   #Removing the last item of the list because it is an empty string
    y = [float(i) for i in data[3].split(",")[:-1]]             #Removing the last item of the list because it is an empty string
    y_trad = [float(i) for i in data_trad[3].split(",")[:-1]]   #Removing the last item of the list because it is an empty string
    h = [float(i) for i in data[4].split(",")[:-1]]             #Removing the last item of the list because it is an empty string
    h_trad = [float(i) for i in data_trad[4].split(",")[:-1]]   #Removing the last item of the list because it is an empty string
    v = [float(i) for i in data[5].split(",")[:-1]]             #Removing the last item of the list because it is an empty string
    v_trad = [float(i) for i in data_trad[5].split(",")[:-1]]   #Removing the last item of the list because it is an empty string

    errors = calcError(x, y, 108)
    errors_trad = calcError(x_trad, y_trad, 108)

    print(errors)
    print(errors_trad)

    length = max(len(t), len(t_trad))   #Since the lengths of the PI(t)D(t) and PID lists may be different, the for loops have to loop over enough times to process both
    for i in range(length):
        #Must surround with try catches to avoid out of bounds exceptions, if the lengths of the lists for PID and PI(t)D(t) are different
        try:
            sheet.write(i+1, 0, t[i], style)
            sheet.write(i+1, 1, errors[i], style)
            sheet.write(i+1, 2, y[i], style)
            sheet.write(i+1, 3, v[i], style)
        except:
            pass
        try:
            sheet.write(i+1, 5, t_trad[i]*1000, style)
            sheet.write(i+1, 6, errors_trad[i], style)
            sheet.write(i+1, 7, y_trad[i], style)
            sheet.write(i+1, 8, v_trad[i], style)
        except:
            pass

    counter += 1

wb.save("TelemetryAnalysisPIDPathFollowingTests.xls")

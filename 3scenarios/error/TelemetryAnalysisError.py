import numpy as np
import matplotlib.pyplot as plt
import xlwt
from xlwt import Workbook
import pandas as pd

# from errorTelemetryTest import *

s1 = []
s1_t = []
s2 = []
s2_t = []
s3 = []
s3_t = []

from s1t1 import *
s1.append(data)
from s1t2 import *
s1.append(data)
from s1t3 import *
s1.append(data)
from s1t4 import *
s1.append(data)
from s1t5 import *
s1.append(data)
from s1t6 import *
s1.append(data)
from s1t7 import *
s1.append(data)
from s1t8 import *
s1.append(data)
from s1t9 import *
s1.append(data)
from s1t10 import *
s1.append(data)

from s2t1 import *
s2.append(data)
from s2t2 import *
s2.append(data)
from s2t3 import *
s2.append(data)
from s2t4 import *
s2.append(data)
from s2t5 import *
s2.append(data)
from s2t6 import *
s2.append(data)
from s2t7 import *
s2.append(data)
from s2t8 import *
s2.append(data)
from s2t9 import *
s2.append(data)
from s2t10 import *
s2.append(data)

from s3t1 import *
s3.append(data)
from s3t2 import *
s3.append(data)
from s3t3 import *
s3.append(data)
from s3t4 import *
s3.append(data)
from s3t5 import *
s3.append(data)
from s3t6 import *
s3.append(data)
from s3t7 import *
s3.append(data)
from s3t8 import *
s3.append(data)
from s3t9 import *
s3.append(data)
from s3t10 import *
s3.append(data)

from s1t1_t import *
s1_t.append(data)
from s1t2_t import *
s1_t.append(data)
from s1t3_t import *
s1_t.append(data)
from s1t4_t import *
s1_t.append(data)
from s1t5_t import *
s1_t.append(data)
from s1t6_t import *
s1_t.append(data)
from s1t7_t import *
s1_t.append(data)
from s1t8_t import *
s1_t.append(data)
from s1t9_t import *
s1_t.append(data)
from s1t10_t import *
s1_t.append(data)

from s2t1_t import *
s2_t.append(data)
from s2t2_t import *
s2_t.append(data)
from s2t3_t import *
s2_t.append(data)
from s2t4_t import *
s2_t.append(data)
from s2t5_t import *
s2_t.append(data)
from s2t6_t import *
s2_t.append(data)
from s2t7_t import *
s2_t.append(data)
from s2t8_t import *
s2_t.append(data)
from s2t9_t import *
s2_t.append(data)
from s2t10_t import *
s2_t.append(data)

from s3t1_t import *
s3_t.append(data)
from s3t2_t import *
s3_t.append(data)
from s3t3_t import *
s3_t.append(data)
from s3t4_t import *
s3_t.append(data)
from s3t5_t import *
s3_t.append(data)
from s3t6_t import *
s3_t.append(data)
from s3t7_t import *
s3_t.append(data)
from s3t8_t import *
s3_t.append(data)
from s3t9_t import *
s3_t.append(data)
from s3t10_t import *
s3_t.append(data)

data = []
data_t = []
data.append(s1)
data.append(s2)
data.append(s3)
data_t.append(s1_t)
data_t.append(s2_t)
data_t.append(s3_t)

styleBold = xlwt.easyxf('font: bold 1')
style = xlwt.easyxf('font: bold off')

#For PI(t)D(t) Scenarios 1, 2, and 3
for n in range(3):
    wb = Workbook()

    for i in range(10):
        sheet = wb.add_sheet('T' + str(i+1))

        #row, column
        sheet.write(0, 0, 'PI(t)D(t) Time', styleBold)
        sheet.write(0, 1, 'PI(t)D(t) Error', styleBold)

        trial_data = data[n][i]
        trial_data = trial_data.replace("\n", "").replace("\r", "").replace(" ", "").replace(':','').replace('t','').split('e')
        times = trial_data[0].split(',')[:-1]
        errors = trial_data[1].split(',')[:-1]

        times = list(map(float, times))
        errors = list(map(float, errors))

        # plt.scatter(times, errors, alpha=0.5)
        # plt.show()

        for i in range(len(times)):
            sheet.write(i+1, 0, times[i], style)
            sheet.write(i+1, 1, errors[i], style)


    wb.save("ErrorVsTime_S" + str(n+1) + ".xls")

#For Traditional PID Scenarios 1, 2, and 3
for n in range(3):
    wb = Workbook()

    for i in range(10):
        sheet = wb.add_sheet('T' + str(i+1))

        #row, column
        sheet.write(0, 0, 'PID Time', styleBold)
        sheet.write(0, 1, 'PID Error', styleBold)

        trial_data = data_t[n][i]
        trial_data = trial_data.replace("\n", "").replace("\r", "").replace(" ", "").replace(':','').replace('t','').split('e')
        times = trial_data[0].split(',')[:-1]
        errors = trial_data[1].split(',')[:-1]

        times = list(map(float, times))
        errors = list(map(float, errors))

        # plt.scatter(times, errors, alpha=0.5)
        # plt.show()

        for i in range(len(times)):
            sheet.write(i+1, 0, times[i], style)
            sheet.write(i+1, 1, errors[i], style)


    wb.save("ErrorVsTime_S" + str(n+1) + "Traditional.xls")

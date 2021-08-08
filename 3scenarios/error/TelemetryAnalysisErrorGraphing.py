"""
This script takes the error files in 3scenarios/error/, exports the error and time values
to an excel sheet, then creates error vs time plots for all 10 trials for all 3 scenarios.
"""

import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from bisect import bisect

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



def extrapolate_coordinates(times, errors):
    """
    This function takes the time and error telemetry then interpolates an error value
    at an inverval of 1ms.

    This is used to create an average graph of all trials.

    Note: In the case that one trial takes less time than another, the shortage of
    time will be filled with repeats of the last coordinate.
    """
    times = [round(x) for x in times]   #convert all times to integers

    new_times = []
    new_errors = []
    current_time_index = 0
    for i in range(int(times[len(times)-1]) + 1):
        new_times.append(i)
        if i == times[current_time_index+1]:
            new_errors.append(round(errors[current_time_index+1], 3))
            current_time_index += 1
        else:
            prev_time = times[current_time_index]
            next_time = times[current_time_index+1]
            prev_error = errors[current_time_index]
            next_error = errors[current_time_index+1]

            time_diff = next_time-prev_time
            ratio = (i-prev_time) / time_diff

            new_error = prev_error + (ratio * (next_error-prev_error))
            new_errors.append(round(new_error, 3))
    return new_times, new_errors





#For PI(t)D(t) Scenarios 1, 2, and 3
for n in range(3):
    writer = pd.ExcelWriter("ErrorVsTime_S" + str(n+1) + ".xlsx", engine='xlsxwriter')

    all_extrapolated_times = []
    all_extrapolated_errors = []

    for i in range(10):
        sheet_name = 'T' + str(i+1)

        trial_data = data[n][i]
        trial_data = trial_data.replace("\n", "").replace("\r", "").replace(" ", "").replace(':','').replace('t','').split('e')
        times = trial_data[0].split(',')[:-1]
        errors = trial_data[1].split(',')[:-1]

        times = list(map(float, times))
        errors = list(map(float, errors))

        extrapolated_times, extrapolated_errors = extrapolate_coordinates(times, errors)    #Used to create the average error vs time from all trials
        all_extrapolated_times.append(extrapolated_times)
        all_extrapolated_errors.append(extrapolated_errors)

        #Create Pandas DataFrame
        d = {'PI(t)D(t) Time': times, 'PI(t)D(t) Error': errors}
        df = pd.DataFrame(data=d)
        # print(df)

        df.to_excel(writer, sheet_name=sheet_name)

        # Access the XlsxWriter workbook and worksheet objects from the dataframe.
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Create a chart object.
        chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight'})
        # Configure the series of the chart from the dataframe data.
        max_row = len(df)
        chart.add_series({
            'categories':     '='+sheet_name+'!$B$2:$B$' + str(max_row),
            'values':         '='+sheet_name+'!$C$2:$C$' + str(max_row),
            # 'line':           {'color': '#bf0000', 'width': 1.5, 'dash_type': 'dash'},    #Dashed line
            'line':           {'color': '#bf0000', 'width': 1.5},
        })

        # Configure the chart axes.
        if n==0:
            max_x=3400
            max_y=108
        elif n==1:
            max_x=4400
            max_y=120
        elif n==2:
            max_x=6000
            max_y=200
        chart.set_x_axis({'name': 'Time (ms)', 'min': 0, 'max': max_x, 'name_font':{'name':'Arial','size':12, 'bold': True}, 'num_font':  {'name': 'Arial', 'size': 12, 'bold': True, 'rotation': 0}, 'line': {'color': '#000000', 'width': 1.5}, 'major_tick_mark': 'inside', 'minor_tick_mark': 'inside'})
        chart.set_y_axis({'name': 'Error (in)', 'min': 0, 'max': max_y, 'name_font':{'name':'Arial','size':12, 'bold': True}, 'num_font':  {'name': 'Arial', 'size': 12, 'bold': True}, 'line': {'color': '#000000', 'width': 1.5}, 'major_tick_mark': 'inside', 'minor_tick_mark': 'inside', 'major_gridlines': {'visible': False}})
        chart.set_title({'name': 'Error\nScenario ' + str(n+1) + ', Trial ' + str(i+1), 'name_font':{'name':'Arial','size':12, 'bold': True}})
        chart.set_legend({'none': True})

        chart.set_plotarea({
            'layout': {
                'x':      0.20,
                'y':      0.12,
                'width':  0.73,
                'height': 0.6,
            }
        })
        chart.set_size({'width': 400, 'height': 240})

        worksheet.insert_chart('D2', chart)



    #Creating Consistency Plot (Count vs Time-To-Reach-Destination)
    smallest_time = 9999
    largest_time = 0
    for i in range(10):
        time = all_extrapolated_times[i][len(all_extrapolated_times[i])-1]
        if time < smallest_time:
            smallest_time = int(time)
        elif time > largest_time:
            largest_time = int(time)
    # time_range1 = largest_time - smallest_time
    mid = int((largest_time+smallest_time)/2)
    # time_low_bound = smallest_time - int(time_range1/14)
    # time_high_bound = largest_time + int(time_range1/14)
    time_low_bound = mid - 299
    time_high_bound = mid + 300
    time_range = time_high_bound - time_low_bound    #time_range is always kept at 600! This is constant for all trials to presever scale in the plots.

    bin_dividers = [time_low_bound]
    for i in range(8):
        bin_dividers.append(bin_dividers[i] + time_range/8)
    bin_dividers = [int(x) for x in bin_dividers]
    #Getting the list of time intervals (all the x-coordinates for the plot)
    time_intervals = []
    for i in range(8):
        time_intervals.append(str(bin_dividers[i]) + "-" + str(bin_dividers[i+1]-1))
    # print(time_intervals)
    # print(bin_dividers)
    #Getting the list of count for each time interval (all the y-coordinate for the plot)
    count = [0, 0, 0, 0, 0, 0, 0, 0]
    for i in range(10):
        bin_num = bisect(bin_dividers, all_extrapolated_times[i][len(all_extrapolated_times[i])-1])
        count[bin_num-1] += 1
    # print(count)




    #Creating average plot of all trials
    #Ensure all lists from all 10 trials are the same length
    max_length = 0
    for i in range(10):
        if all_extrapolated_times[i][len(all_extrapolated_times[i])-1] > max_length:
            max_length = all_extrapolated_times[i][len(all_extrapolated_times[i])-1]
    for i in range(10):
        diff = max_length - len(all_extrapolated_times[i])
        for j in range(diff):
            all_extrapolated_times[i].append(all_extrapolated_times[i][len(all_extrapolated_times[i])-1]+1)     #Keep increasing the time 1 ms at a time
            # all_extrapolated_errors[i].append(all_extrapolated_errors[i][len(all_extrapolated_errors[i])-1])
            all_extrapolated_errors[i].append(0.0)  #Add errors of 0.0 to the end

    #Calculate Average Times and Average Errors Lists
    avg_times = all_extrapolated_times[0]
    avg_errors = []
    for i in range(len(avg_times)):
        total = 0
        for j in range(10):
            total += all_extrapolated_errors[j][i]
        avg = round(total/10, 3)
        avg_errors.append(avg)

    #Save the Average Times and Average Errors to Excel and produce a plot
    #Create Pandas DataFrame
    d = {'PI(t)D(t) Avg Time': avg_times, 'PI(t)D(t) Avg Error': avg_errors}
    df = pd.DataFrame(data=d)
    df.to_excel(writer, sheet_name="Average")
    # Access the XlsxWriter workbook and worksheet objects from the dataframe.
    workbook = writer.book
    worksheet = writer.sheets["Average"]
    # Create a chart object.
    chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight'})
    # Configure the series of the chart from the dataframe data.
    chart.add_series({
        'categories':     '='+'Average'+'!$B$2:$B$9999',
        'values':         '='+'Average'+'!$C$2:$C$9999',
        'line':           {'color': '#bf0000', 'width': 1.5},
    })
    # Configure the chart axes.
    if n==0:
        max_x=3400
        max_y=108
    elif n==1:
        max_x=4400
        max_y=120
    elif n==2:
        max_x=6000
        max_y=200
    chart.set_x_axis({'name': 'Time (ms)', 'min': 0, 'max': max_x, 'name_font':{'name':'Arial','size':12, 'bold': True}, 'num_font':  {'name': 'Arial', 'size': 12, 'bold': True, 'rotation': 0}, 'line': {'color': '#000000', 'width': 1.5}, 'major_tick_mark': 'inside', 'minor_tick_mark': 'inside'})
    chart.set_y_axis({'name': 'Error (in)', 'min': 0, 'max': max_y, 'name_font':{'name':'Arial','size':12, 'bold': True}, 'num_font':  {'name': 'Arial', 'size': 12, 'bold': True}, 'line': {'color': '#000000', 'width': 1.5}, 'major_tick_mark': 'inside', 'minor_tick_mark': 'inside', 'major_gridlines': {'visible': False}})
    chart.set_title({'name': 'Error\nScenario ' + str(n+1) + '; Average', 'name_font':{'name':'Arial','size':12, 'bold': True}})
    chart.set_legend({'none': True})
    chart.set_plotarea({
        'layout': {
            'x':      0.20,
            'y':      0.12,
            'width':  0.73,
            'height': 0.6,
        }
    })
    chart.set_size({'width': 400, 'height': 240})
    worksheet.insert_chart('D2', chart)


    #Inserting Sheet and Chart for Consistency Data
    d = {'PI(t)D(t) Time Interval': time_intervals, 'PI(t)D(t) Count': count}
    df = pd.DataFrame(data=d)
    df.to_excel(writer, sheet_name="CountVsTime")
    # Access the XlsxWriter workbook and worksheet objects from the dataframe.
    workbook = writer.book
    worksheet = writer.sheets["CountVsTime"]
    # Create a chart object.
    chart = workbook.add_chart({'type': 'column'})
    # Configure the series of the chart from the dataframe data.
    chart.add_series({
        'categories':     '='+'CountVsTime'+'!$B$2:$B$9',
        'values':         '='+'CountVsTime'+'!$C$2:$C$9',
        'line':           {'color': '#000000', 'width': 1.5},
        'fill':           {'color': '#bf0000'}
    })
    line_chart = workbook.add_chart({'type': 'line'})
    line_chart.add_series({
        'categories':     '='+'CountVsTime'+'!$B$2:$B$9',
        'values':         '='+'CountVsTime'+'!$C$2:$C$9',
        'line':           {'color': '#700000', 'width': 1.5},
        'smooth':         True,
        'marker':         {'type': 'diamond', 'fill': {'color': '#700000'}, 'border': {'color': '#700000'}}
    })
    chart.combine(line_chart)
    chart.set_x_axis({'name': 'Time To Reach Target (ms)', 'name_font':{'name':'Arial','size':12, 'bold': True}, 'num_font':  {'name': 'Arial', 'size': 9, 'bold': True}, 'line': {'color': '#000000', 'width': 1.5}, 'major_tick_mark': 'inside', 'minor_tick_mark': 'inside'})
    chart.set_y_axis({'name': 'Count', 'min': 0, 'max': 8, 'name_font':{'name':'Arial','size':12, 'bold': True}, 'num_font':  {'name': 'Arial', 'size': 12, 'bold': True}, 'line': {'color': '#000000', 'width': 1.5}, 'major_tick_mark': 'inside', 'minor_tick_mark': 'inside', 'major_gridlines': {'visible': False}})
    chart.set_title({'name': 'Count vs Time\nScenario ' + str(n+1), 'name_font':{'name':'Arial','size':12, 'bold': True}})
    chart.set_legend({'none': True})
    chart.set_plotarea({
        'layout': {
            'x':      0.14,
            'y':      0.2,
            'width':  0.82,
            'height': 0.5,
        }
    })
    chart.set_size({'width': 400, 'height': 300})
    worksheet.insert_chart('D2', chart)



    #Save the excel file
    writer.save()





#For Traditional PID Scenarios 1, 2, and 3
for n in range(3):
    writer = pd.ExcelWriter("ErrorVsTime_S" + str(n+1) + "_Traditional.xlsx", engine='xlsxwriter')

    all_extrapolated_times = []
    all_extrapolated_errors = []

    for i in range(10):
        sheet_name = 'T' + str(i+1)

        trial_data = data_t[n][i]
        trial_data = trial_data.replace("\n", "").replace("\r", "").replace(" ", "").replace(':','').replace('t','').split('e')
        times = trial_data[0].split(',')[:-1]
        errors = trial_data[1].split(',')[:-1]

        times = list(map(float, times))
        times = [i * 1000 for i in times]   #Times need to be converted from seconds to milliseconds for the traditional PID data
        errors = list(map(float, errors))

        extrapolated_times, extrapolated_errors = extrapolate_coordinates(times, errors)    #Used to create the average error vs time from all trials
        all_extrapolated_times.append(extrapolated_times)
        all_extrapolated_errors.append(extrapolated_errors)

        d = {'PID Time': times, 'PID Error': errors}
        df = pd.DataFrame(data=d)
        # print(df)

        df.to_excel(writer, sheet_name=sheet_name)

        # Access the XlsxWriter workbook and worksheet objects from the dataframe.
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Create a chart object.
        chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight'})
        # Configure the series of the chart from the dataframe data.
        chart.add_series({
            'categories':     '='+sheet_name+'!$B$2:$B$' + str(max_row),
            'values':         '='+sheet_name+'!$C$2:$C$' + str(max_row),
            # 'line':           {'color': '#bf0000', 'width': 1.5, 'dash_type': 'dash'},    #Dashed line
            'line':           {'width': 1.5},
        })

        # Configure the chart axes.
        if n==0:
            max_x=4000
            max_y=108
        elif n==1:
            max_x=4800
            max_y=120
        elif n==2:
            max_x=7000
            max_y=200
        chart.set_x_axis({'name': 'Time (ms)', 'min': 0, 'max': max_x, 'name_font':{'name':'Arial','size':12, 'bold': True}, 'num_font':  {'name': 'Arial', 'size': 12, 'bold': True, 'rotation': 0}, 'line': {'color': '#000000', 'width': 1.5}, 'major_tick_mark': 'inside', 'minor_tick_mark': 'inside'})
        chart.set_y_axis({'name': 'Error (in)', 'min': 0, 'max': max_y, 'name_font':{'name':'Arial','size':12, 'bold': True}, 'num_font':  {'name': 'Arial', 'size': 12, 'bold': True}, 'line': {'color': '#000000', 'width': 1.5}, 'major_tick_mark': 'inside', 'minor_tick_mark': 'inside', 'major_gridlines': {'visible': False}})
        chart.set_title({'name': 'Error\nScenario ' + str(n+1) + ', Trial ' + str(i+1), 'name_font':{'name':'Arial','size':12, 'bold': True}})
        chart.set_legend({'none': True})

        chart.set_plotarea({
            'layout': {
                'x':      0.20,
                'y':      0.12,
                'width':  0.73,
                'height': 0.6,
            }
        })
        chart.set_size({'width': 400, 'height': 240})

        worksheet.insert_chart('D2', chart)





    #Creating Consistency Plot (Count vs Time-To-Reach-Destination)
    smallest_time = 9999
    largest_time = 0
    for i in range(10):
        time = all_extrapolated_times[i][len(all_extrapolated_times[i])-1]
        if time < smallest_time:
            smallest_time = int(time)
        elif time > largest_time:
            largest_time = int(time)
    # time_range1 = largest_time - smallest_time
    mid = int((largest_time+smallest_time)/2)
    # time_low_bound = smallest_time - int(time_range1/14)
    # time_high_bound = largest_time + int(time_range1/14)
    time_low_bound = mid - 299
    time_high_bound = mid + 300
    time_range = time_high_bound - time_low_bound    #time_range is always kept at 600! This is constant for all trials to presever scale in the plots.

    bin_dividers = [time_low_bound]
    for i in range(8):
        bin_dividers.append(bin_dividers[i] + time_range/8)
    bin_dividers = [int(x) for x in bin_dividers]
    #Getting the list of time intervals (all the x-coordinates for the plot)
    time_intervals = []
    for i in range(8):
        time_intervals.append(str(bin_dividers[i]) + "-" + str(bin_dividers[i+1]-1))
    # print(time_intervals)
    # print(bin_dividers)
    #Getting the list of count for each time interval (all the y-coordinate for the plot)
    count = [0, 0, 0, 0, 0, 0, 0, 0]
    for i in range(10):
        bin_num = bisect(bin_dividers, all_extrapolated_times[i][len(all_extrapolated_times[i])-1])
        count[bin_num-1] += 1
    # print(count)






    #Creating average plot of all trials
    #Ensure all lists from all 10 trials are the same length
    max_length = 0
    for i in range(10):
        if all_extrapolated_times[i][len(all_extrapolated_times[i])-1] > max_length:
            max_length = all_extrapolated_times[i][len(all_extrapolated_times[i])-1]
    for i in range(10):
        diff = max_length - len(all_extrapolated_times[i])
        for j in range(diff):
            all_extrapolated_times[i].append(all_extrapolated_times[i][len(all_extrapolated_times[i])-1]+1)     #Keep increasing the time 1 ms at a time
            # all_extrapolated_errors[i].append(all_extrapolated_errors[i][len(all_extrapolated_errors[i])-1])
            all_extrapolated_errors[i].append(0.0)  #Add errors of 0.0 to the end

    #Calculate Average Times and Average Errors Lists
    avg_times = all_extrapolated_times[0]
    avg_errors = []
    for i in range(len(avg_times)):
        total = 0
        for j in range(10):
            total += all_extrapolated_errors[j][i]
        avg = round(total/10, 3)
        avg_errors.append(avg)

    #Save the Average Times and Average Errors to Excel and produce a plot
    #Create Pandas DataFrame
    d = {'PID Avg Time': avg_times, 'PID Avg Error': avg_errors}
    df = pd.DataFrame(data=d)
    df.to_excel(writer, sheet_name="Average")
    # Access the XlsxWriter workbook and worksheet objects from the dataframe.
    workbook = writer.book
    worksheet = writer.sheets["Average"]
    # Create a chart object.
    chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight'})
    # Configure the series of the chart from the dataframe data.
    chart.add_series({
        'categories':     '='+'Average'+'!$B$2:$B$9999',
        'values':         '='+'Average'+'!$C$2:$C$9999',
        'line':           {'width': 1.5},
    })
    # Configure the chart axes.
    if n==0:
        max_x=4000
        max_y=108
    elif n==1:
        max_x=4800
        max_y=120
    elif n==2:
        max_x=7000
        max_y=200
    chart.set_x_axis({'name': 'Time (ms)', 'min': 0, 'max': max_x, 'name_font':{'name':'Arial','size':12, 'bold': True}, 'num_font':  {'name': 'Arial', 'size': 12, 'bold': True, 'rotation': 0}, 'line': {'color': '#000000', 'width': 1.5}, 'major_tick_mark': 'inside', 'minor_tick_mark': 'inside'})
    chart.set_y_axis({'name': 'Error (in)', 'min': 0, 'max': max_y, 'name_font':{'name':'Arial','size':12, 'bold': True}, 'num_font':  {'name': 'Arial', 'size': 12, 'bold': True}, 'line': {'color': '#000000', 'width': 1.5}, 'major_tick_mark': 'inside', 'minor_tick_mark': 'inside', 'major_gridlines': {'visible': False}})
    chart.set_title({'name': 'Error\nScenario ' + str(n+1) + '; Average', 'name_font':{'name':'Arial','size':12, 'bold': True}})
    chart.set_legend({'none': True})
    chart.set_plotarea({
        'layout': {
            'x':      0.20,
            'y':      0.12,
            'width':  0.73,
            'height': 0.6,
        }
    })
    chart.set_size({'width': 400, 'height': 240})
    worksheet.insert_chart('D2', chart)



    #Inserting Sheet and Chart for Consistency Data
    d = {'PID Time Interval': time_intervals, 'PID Count': count}
    df = pd.DataFrame(data=d)
    df.to_excel(writer, sheet_name="CountVsTime")
    # Access the XlsxWriter workbook and worksheet objects from the dataframe.
    workbook = writer.book
    worksheet = writer.sheets["CountVsTime"]
    # Create a chart object.
    chart = workbook.add_chart({'type': 'column'})
    # Configure the series of the chart from the dataframe data.
    chart.add_series({
        'categories':     '='+'CountVsTime'+'!$B$2:$B$9',
        'values':         '='+'CountVsTime'+'!$C$2:$C$9',
        'line':           {'color': '#000000', 'width': 1.5},
    })
    line_chart = workbook.add_chart({'type': 'line'})
    line_chart.add_series({
        'categories':     '='+'CountVsTime'+'!$B$2:$B$9',
        'values':         '='+'CountVsTime'+'!$C$2:$C$9',
        'line':           {'color': '#00306b', 'width': 1.5},
        'smooth':         True,
        'marker':         {'type': 'diamond', 'fill': {'color': '#00306b'}, 'border': {'color': '#00306b'}}
    })
    chart.combine(line_chart)
    chart.set_x_axis({'name': 'Time To Reach Target (ms)', 'name_font':{'name':'Arial','size':12, 'bold': True}, 'num_font':  {'name': 'Arial', 'size': 9, 'bold': True}, 'line': {'color': '#000000', 'width': 1.5}, 'major_tick_mark': 'inside', 'minor_tick_mark': 'inside'})
    chart.set_y_axis({'name': 'Count', 'min': 0, 'max': 8, 'name_font':{'name':'Arial','size':12, 'bold': True}, 'num_font':  {'name': 'Arial', 'size': 12, 'bold': True}, 'line': {'color': '#000000', 'width': 1.5}, 'major_tick_mark': 'inside', 'minor_tick_mark': 'inside', 'major_gridlines': {'visible': False}})
    chart.set_title({'name': 'Count vs Time\nScenario ' + str(n+1), 'name_font':{'name':'Arial','size':12, 'bold': True}})
    chart.set_legend({'none': True})
    chart.set_plotarea({
        'layout': {
            'x':      0.14,
            'y':      0.2,
            'width':  0.82,
            'height': 0.5,
        }
    })
    chart.set_size({'width': 400, 'height': 300})
    worksheet.insert_chart('D2', chart)


    writer.save()

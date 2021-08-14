"""
COMPLETELY FINISHED. DO NOT EDIT OR RERUN.
"""

import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import time
import math
from scipy import spatial
from bisect import bisect

# def interpolate_coordinates(times_raw, x_raw, y_raw, ms = True):
#     """
#     This function takes the time and position telemetry then interpolates an error value
#     at an inverval of 1ms.
#
#     This is used to calculate RMSE error for each trial.
#     """
#     if not ms:
#         times_raw = [1000*i for i in times_raw]     #Convert from sec to ms
#         times_raw = [round(i) for i in times_raw]   #convert all times to integers
#
#     #Getting rid of data points that have the exact same millisecond time stamp (this causes a bug with the code below)
#     times = []
#     x = []
#     y = []
#     for i in range(len(times_raw)-1):
#         if i == 0:
#             times.append(times_raw[i])
#             x.append(x_raw[i])
#             y.append(y_raw[i])
#         else:
#             if times_raw[i] != times_raw[i+1]:
#                 times.append(times_raw[i+1])
#                 x.append(x_raw[i+1])
#                 y.append(y_raw[i+1])
#     # print(len(times))
#     # print(len(x))
#     # print(len(y))
#     # print(times)
#
#
#     new_times = []
#     new_x = []
#     new_y = []
#     current_time_index = 0
#     time.sleep(3)
#     for i in range(0, int(times[len(times)-1]) + 1, 1):
#         new_times.append(i)
#         # print("current_time_index: " + str(current_time_index))
#         if (i - times[current_time_index+1]) <= 1 and current_time_index < len(times)-2:
#             new_x.append(round(x[current_time_index+1], 3))
#             new_y.append(round(y[current_time_index+1], 3))
#             current_time_index += 1
#         else:
#             prev_time = times[current_time_index]
#             next_time = times[current_time_index+1]
#             prev_x = x[current_time_index]
#             prev_y = y[current_time_index]
#             next_x = x[current_time_index+1]
#             next_y = y[current_time_index+1]
#
#             time_diff = next_time-prev_time
#             if time_diff == 0:
#                 # new_x.append(prev_x)
#                 # new_y.append(prev_y)
#                 del new_times[-1]
#             else:
#                 ratio = (i-prev_time) / time_diff
#
#                 new_x_val = prev_x + (ratio * (next_x - prev_x))
#                 new_y_val = prev_y + (ratio * (next_y - prev_y))
#                 new_x.append(round(new_x_val, 3))
#                 new_y.append(round(new_y_val, 3))
#     return new_times, new_x, new_y

def interpolate_coordinates(times_raw, x_raw, y_raw, ms=True):
    """
    This function takes the time and position telemetry then interpolates an error value
    at an inverval of 1ms.
    This is used to calculate RMSE error for each trial.
    """
    if not ms:
        times_raw = [1000*i for i in times_raw]     #Convert from sec to ms
        times_raw = [round(i) for i in times_raw]   #convert all times to integers

    #Getting rid of data points that have the exact same millisecond time stamp (this causes a bug with the code below)
    times = []
    x = []
    y = []
    for i in range(len(times_raw)-1):
        if i == 0:
            times.append(times_raw[i])
            x.append(x_raw[i])
            y.append(y_raw[i])
        else:
            if times_raw[i] != times_raw[i+1]:
                times.append(times_raw[i+1])
                x.append(x_raw[i+1])
                y.append(y_raw[i+1])
    # print(len(times))
    # print(len(x))
    # print(len(y))
    # print(times)


    new_times = []
    new_x = []
    new_y = []
    current_time_index = 0
    for i in range(int(times[len(times)-1]) + 1):
        new_times.append(i)
        if i == times[current_time_index+1]:
            new_x.append(round(x[current_time_index+1], 3))
            new_y.append(round(y[current_time_index+1], 3))
            current_time_index += 1
        else:
            prev_time = times[current_time_index]
            next_time = times[current_time_index+1]
            prev_x = x[current_time_index]
            prev_y = y[current_time_index]
            next_x = x[current_time_index+1]
            next_y = y[current_time_index+1]

            time_diff = next_time-prev_time
            if time_diff == 0:
                # new_x.append(prev_x)
                # new_y.append(prev_y)
                del new_times[-1]
            else:
                ratio = (i-prev_time) / time_diff

                new_x_val = prev_x + (ratio * (next_x - prev_x))
                new_y_val = prev_y + (ratio * (next_y - prev_y))
                new_x.append(round(new_x_val, 3))
                new_y.append(round(new_y_val, 3))
    return new_times, new_x, new_y



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

#Setting lists to contain deviation from path data
mae = []
mae_t = []
max_deviations = []
# mae_errors_1 = []
# mae_errors_2 = []
# mae_errors_3 = []
# mae_errors_4 = []
# mae_errors_5 = []
# mae_errors_6 = []
# mae_errors_7 = []
# mae_errors_8 = []
# mae_errors_9 = []
# mae_errors_10 = []
# mae_errors_avg = []

#For PI(t)D(t) Scenarios 1, 2, and 3
for n in range(3):
    writer = pd.ExcelWriter("PositionTelemetry_S" + str(n+1) + ".xlsx", engine='xlsxwriter')

    all_speeds = []
    for i in range(10):
        print("PI(t)D(t) Scenario %d Trial %d" % (n+1, i+1))
        sheet_name = 'T' + str(i+1)



        """REAL TELEMETRY"""
        trial_data = data[n][i].split('*')[1]

        #Delete whitespace and alphabetical letters
        trial_data = trial_data.replace("\n", "").replace("\r", "").replace(" ", "").replace("angV", "").replace("angA", "").replace("t", "").replace("x", "").replace("y", "").replace("h", "").replace("v", "").replace("a", "")

        #Split data in a list with t, x, y, h, etc. separated
        trial_data = trial_data.split(':')

        #Split t, x, y, h, etc. into lists of individual doubles
        t = [float(a) for a in trial_data[1].split(",")[:-1]]             #Removing the last item of the list because it is an empty string
        x = [float(a) for a in trial_data[2].split(",")[:-1]]             #Removing the last item of the list because it is an empty string
        y = [float(a) for a in trial_data[3].split(",")[:-1]]             #Removing the last item of the list because it is an empty string
        h = [float(a) for a in trial_data[4].split(",")[:-1]]             #Removing the last item of the list because it is an empty string
        # print(x)
        # print()
        # print(y)
        # print()
        total_distance = 0
        for j in range(len(x)-1):
            total_distance += math.hypot(x[j+1]-x[j], y[j+1]-y[j])
        # print("total_distance: " + str(total_distance))
        avg_speed = total_distance/(t[len(t)-1]/1000)
        # print("avg_speed: " + str(avg_speed))
        all_speeds.append(avg_speed)



        """MOTION PROFILE TELEMETRY"""
        trial_data_mp = data[n][i].split('*')[0].lower()

        #Delete whitespace and alphabetical letters
        trial_data_mp = trial_data_mp.replace("\n", "").replace("=", "").replace("{", "").replace(" ", "").replace("t", "").replace("x", "").replace("y", "").replace("h", "").replace("veloci", "").replace("acceleraion", "").replace("ang", "")

        mp_master_list = trial_data_mp.split('}')
        for j in range(len(mp_master_list)):
            mp_master_list[j] = mp_master_list[j].split(',')

        #Saving all the MP data even though most of it won't be used
        t_mp = []
        x_mp = []
        y_mp = []
        h_mp = []
        v_mp = []
        a_mp = []
        angV_mp = []
        angA_mp = []

        for k in range(len(mp_master_list)):
            # print(mp_master_list[i])
            if len(mp_master_list[k]) > 2:
                t_mp.append(mp_master_list[k][0])
                x_mp.append(mp_master_list[k][1])
                y_mp.append(mp_master_list[k][2])
                h_mp.append(mp_master_list[k][3])
                v_mp.append(mp_master_list[k][4])
                a_mp.append(mp_master_list[k][5])
                angV_mp.append(mp_master_list[k][6])
                angA_mp.append(mp_master_list[k][7])

        t_mp = [float(a) for a in t_mp]
        x_mp = [float(a) for a in x_mp]
        y_mp = [float(a) for a in y_mp]
        h_mp = [float(a) for a in h_mp]
        v_mp = [float(a) for a in v_mp]
        a_mp = [float(a) for a in a_mp]
        angV_mp = [float(a) for a in angV_mp]
        angA_mp = [float(a) for a in angA_mp]

        # print("--------------------------------")
        # print(x_mp)
        # print()
        # print(y_mp)
        # print()



        """CALCULATE ERROR BETWEEN REAL AND MP"""
        interpolated_t, interpolated_x, interpolated_y = interpolate_coordinates(t, x, y)
        interpolated_t_mp, interpolated_x_mp, interpolated_y_mp = interpolate_coordinates(t_mp, x_mp, y_mp, ms=False)

        # plt.scatter(interpolated_x, interpolated_y)
        # plt.scatter(interpolated_x_mp, interpolated_y_mp)
        # plt.show()

        all_errors = []
        for j in range(len(interpolated_t)):    #Looping through every point on the real path
            real_path_pt = [interpolated_x[j], interpolated_y[j]]

            path_mp = np.column_stack((interpolated_x_mp, interpolated_y_mp)) #Converting to np array for vectorization

            distance, index = spatial.KDTree(path_mp).query(real_path_pt)
            # print(distance)
            # print(real_path_pt)
            # print(path_mp[index])
            # print()
            all_errors.append(distance)

            # min_dist = 999.9
            # min_dist_mp_idx = 0
            # for k in range(len(interpolated_t_mp)): #Looping through every point on the MP path to find the nearest point to the point on the real path.
            #     dist = math.hypot(interpolated_x[j]-interpolated_x_mp[k], interpolated_y[j]-interpolated_y_mp[k])
            #     if dist < min_dist:
            #         min_dist = dist
            #         min_dist_mp_idx = k
            #     elif dist > min_dist + 1:   #If the distance starts increasing, means the nearest point has already been passed (adding 1 just for safety)
            #         break
            # all_errors.append(min_dist)

        avg_error = round(sum(all_errors) / len(all_errors), 3)
        print(avg_error)
        mae.append(avg_error)
        #On the last trial of the scenario, calculate the average MAE for the scenario
        if i == 9:
            scenario_total_mae = 0
            for j in range(10):
                scenario_total_mae += mae[len(mae)-1 - j]
            scenario_avg_mae = round(scenario_total_mae/10, 3)
            mae.append(scenario_avg_mae)

        #Calculate and save the max singular error from the 10 trails
        max_deviations.append(round(max(all_errors), 3))
        print("max(all_errors): " + str(max(all_errors)))





        """""PLOTTING IN EXCEL"""""
        #Ensuring all lists are the same length (this is required by Pandas DataFrame)
        l = len(x)
        l_mp = len(x_mp)
        if l > l_mp:
            diff = l - l_mp
            for j in range(diff):
                x_mp.append(x_mp[len(x_mp)-1])  #Repeating the last value
                y_mp.append(y_mp[len(y_mp)-1])  #Repeating the last value
        else:
            diff = l_mp - l
            for j in range(diff):
                x.append(x[len(x)-1])  #Repeating the last value
                y.append(y[len(y)-1])  #Repeating the last value

        #Create Pandas DataFrame
        d = {'X': x, 'Y': y, 'Motion Profile X': x_mp, 'Motion Profile Y': y_mp}
        df = pd.DataFrame(data=d)
        df.to_excel(writer, sheet_name=sheet_name)

        # Access the XlsxWriter workbook and worksheet objects from the dataframe.
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Create a chart object.
        chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight'})
        # Configure the series of the chart from the dataframe data.
        max_row = len(df)
        chart.add_series({  #Real Telemetry
            'categories':     '='+sheet_name+'!$B$2:$B$' + str(max_row),
            'values':         '='+sheet_name+'!$C$2:$C$' + str(max_row),
            # 'line':           {'color': '#bf0000', 'width': 1.5, 'dash_type': 'dash'},    #Dashed line
            'line':           {'color': '#bf0000', 'width': 1.5},
        })
        chart.add_series({  #Desired Path
            'categories':     '='+sheet_name+'!$D$2:$D$' + str(max_row),
            'values':         '='+sheet_name+'!$E$2:$E$' + str(max_row),
            # 'line':           {'color': '#bf0000', 'width': 1.5, 'dash_type': 'dash'},    #Dashed line
            'line':           {'color': '#FF9900', 'width': 2.5, 'dash_type': 'round_dot'},
        })

        # Configure the chart axes.
        if n==0:
            min_x=0
            min_y=0
            max_x=40
            max_y=80
        elif n==1:
            min_x=-72
            min_y=-8
            max_x=8
            max_y=40
        elif n==2:
            min_x=-4
            min_y=0
            max_x=44
            max_y=152

        # chart.set_x_axis({'name': 'X (in)', 'name_font':{'name':'Arial','size':12, 'bold': True}, 'num_font':  {'name': 'Arial', 'size': 12, 'bold': True, 'rotation': 0}, 'line': {'color': '#000000', 'width': 1.5}, 'major_tick_mark': 'inside', 'minor_tick_mark': 'inside'})
        # chart.set_y_axis({'name': 'Y (in)', 'name_font':{'name':'Arial','size':12, 'bold': True}, 'num_font':  {'name': 'Arial', 'size': 12, 'bold': True}, 'line': {'color': '#000000', 'width': 1.5}, 'major_tick_mark': 'inside', 'minor_tick_mark': 'inside', 'major_gridlines': {'visible': False}})
        chart.set_x_axis({'name': 'Time (ms)', 'min': min_x, 'max': max_x, 'name_font':{'name':'Arial','size':12, 'bold': True}, 'num_font':  {'name': 'Arial', 'size': 12, 'bold': True, 'rotation': 0}, 'line': {'color': '#000000', 'width': 1.5}, 'major_tick_mark': 'inside', 'minor_tick_mark': 'inside'})
        chart.set_y_axis({'name': 'Error (in)', 'min': min_y, 'max': max_y, 'name_font':{'name':'Arial','size':12, 'bold': True}, 'num_font':  {'name': 'Arial', 'size': 12, 'bold': True}, 'line': {'color': '#000000', 'width': 1.5}, 'major_tick_mark': 'inside', 'minor_tick_mark': 'inside', 'major_gridlines': {'visible': False}})
        chart.set_title({'name': 'Error\nScenario ' + str(n+1) + ', Trial ' + str(i+1), 'name_font':{'name':'Arial','size':12, 'bold': True}})
        chart.set_legend({'none': True})

        if n==0:
            chart.set_size({'width': 265, 'height': 500})
            chart.set_plotarea({
                'layout': {
                    'x':      0.35,
                    'y':      0.125,
                    'width':  0.75,
                    'height': 0.75,
                }
            })
        elif n==1:
            chart.set_size({'width': 350, 'height': 262})
            chart.set_plotarea({
                'layout': {
                    'x':      0.15,
                    'y':      0.2,
                    'width':  0.8,
                    'height': 0.7,
                }
            })
        elif n==2:
            chart.set_size({'width': 220, 'height': 700})
            chart.set_plotarea({
                'layout': {
                    'x':      0.40,
                    'y':      0.09,
                    'width':  0.73,
                    'height': 0.82,
                }
            })

        worksheet.insert_chart('F2', chart)

    #Writing avg speed
    net_avg_speed = [sum(all_speeds)/len(all_speeds)]
    # print(net_avg_speed)
    #Create Pandas DataFrame
    d = {'Scenario Average Speed (in/s)': net_avg_speed}
    df = pd.DataFrame(data=d)
    df.to_excel(writer, sheet_name="Average Speed")


    writer.save()




#For Tradtional PID Scenarios 1, 2, and 3
for n in range(3):
    writer = pd.ExcelWriter("PositionTelemetry_S" + str(n+1) + "_Traditional.xlsx", engine='xlsxwriter')

    all_speeds = []
    for i in range(10):
        print("Traditional PID Scenario %d Trial %d" % (n+1, i+1))
        sheet_name = 'T' + str(i+1)



        """REAL TELEMETRY"""
        trial_data = data_t[n][i].split('*')[1]

        #Delete whitespace and alphabetical letters
        trial_data = trial_data.replace("\n", "").replace("\r", "").replace(" ", "").replace("angV", "").replace("angA", "").replace("t", "").replace("x", "").replace("y", "").replace("h", "").replace("v", "").replace("a", "")

        #Split data in a list with t, x, y, h, etc. separated
        trial_data = trial_data.split(':')

        #Split t, x, y, h, etc. into lists of individual doubles
        t = [float(a) for a in trial_data[1].split(",")[:-1]]             #Removing the last item of the list because it is an empty string
        x = [float(a) for a in trial_data[2].split(",")[:-1]]             #Removing the last item of the list because it is an empty string
        y = [float(a) for a in trial_data[3].split(",")[:-1]]             #Removing the last item of the list because it is an empty string
        h = [float(a) for a in trial_data[4].split(",")[:-1]]             #Removing the last item of the list because it is an empty string
        # print(x)
        # print()
        # print(y)
        # print()

        total_distance = 0
        for j in range(len(x)-1):
            total_distance += math.hypot(x[j+1]-x[j], y[j+1]-y[j])
        # print("total_distance: " + str(total_distance))
        avg_speed = total_distance/(t[len(t)-1])
        # print("avg_speed: " + str(avg_speed))
        all_speeds.append(avg_speed)



        """MOTION PROFILE TELEMETRY"""
        trial_data_mp = data_t[n][i].split('*')[0].lower()

        #Delete whitespace and alphabetical letters
        trial_data_mp = trial_data_mp.replace("\n", "").replace("=", "").replace("{", "").replace(" ", "").replace("t", "").replace("x", "").replace("y", "").replace("h", "").replace("veloci", "").replace("acceleraion", "").replace("ang", "")

        mp_master_list = trial_data_mp.split('}')
        for j in range(len(mp_master_list)):
            mp_master_list[j] = mp_master_list[j].split(',')

        #Saving all the MP data even though most of it won't be used
        t_mp = []
        x_mp = []
        y_mp = []
        h_mp = []
        v_mp = []
        a_mp = []
        angV_mp = []
        angA_mp = []

        for k in range(len(mp_master_list)):
            # print(mp_master_list[i])
            if len(mp_master_list[k]) > 2:
                t_mp.append(mp_master_list[k][0])
                x_mp.append(mp_master_list[k][1])
                y_mp.append(mp_master_list[k][2])
                h_mp.append(mp_master_list[k][3])
                v_mp.append(mp_master_list[k][4])
                a_mp.append(mp_master_list[k][5])
                angV_mp.append(mp_master_list[k][6])
                angA_mp.append(mp_master_list[k][7])

        t_mp = [float(a) for a in t_mp]
        x_mp = [float(a) for a in x_mp]
        y_mp = [float(a) for a in y_mp]
        h_mp = [float(a) for a in h_mp]
        v_mp = [float(a) for a in v_mp]
        a_mp = [float(a) for a in a_mp]
        angV_mp = [float(a) for a in angV_mp]
        angA_mp = [float(a) for a in angA_mp]

        # print("--------------------------------")
        # print(x_mp)
        # print()
        # print(y_mp)
        # print()



        """CALCULATE ERROR BETWEEN REAL AND MP"""
        interpolated_t, interpolated_x, interpolated_y = interpolate_coordinates(t, x, y, ms=False)
        interpolated_t_mp, interpolated_x_mp, interpolated_y_mp = interpolate_coordinates(t_mp, x_mp, y_mp, ms=False)

        all_errors = []
        for j in range(len(interpolated_t)):    #Looping through every point on the real path
            real_path_pt = [interpolated_x[j], interpolated_y[j]]

            path_mp = np.column_stack((interpolated_x_mp, interpolated_y_mp)) #Converting to np array for vectorization

            distance, index = spatial.KDTree(path_mp).query(real_path_pt)
            # print(distance)
            # print(real_path_pt)
            # print(path_mp[index])
            # print()
            all_errors.append(distance)

            # min_dist = 999.9
            # min_dist_mp_idx = 0
            # for k in range(len(interpolated_t_mp)): #Looping through every point on the MP path to find the nearest point to the point on the real path.
            #     dist = math.hypot(interpolated_x[j]-interpolated_x_mp[k], interpolated_y[j]-interpolated_y_mp[k])
            #     if dist < min_dist:
            #         min_dist = dist
            #         min_dist_mp_idx = k
            #     elif dist > min_dist + 1:   #If the distance starts increasing, means the nearest point has already been passed (adding 1 just for safety)
            #         break
            # all_errors.append(min_dist)

        avg_error = round(sum(all_errors) / len(all_errors), 3)
        print(avg_error)
        mae_t.append(avg_error)
        #On the last trial of the scenario, calculate the average MAE for the scenario
        if i == 9:
            scenario_total_mae = 0
            for j in range(10):
                scenario_total_mae += mae_t[len(mae_t)-1 - j]
            scenario_avg_mae = round(scenario_total_mae/10, 3)
            mae_t.append(scenario_avg_mae)

        #Calculate and save the max singular error from the 10 trails
        max_deviations.append(round(max(all_errors), 3))
        print("max(all_errors): " + str(max(all_errors)))

        # plt.scatter(interpolated_x, interpolated_y)
        # plt.scatter(interpolated_x_mp, interpolated_y_mp)
        # plt.show()





        """""PLOTTING IN EXCEL"""""
        #Ensuring all lists are the same length (this is required by Pandas DataFrame)
        l = len(x)
        l_mp = len(x_mp)
        if l > l_mp:
            diff = l - l_mp
            for j in range(diff):
                x_mp.append(x_mp[len(x_mp)-1])  #Repeating the last value
                y_mp.append(y_mp[len(y_mp)-1])  #Repeating the last value
        else:
            diff = l_mp - l
            for j in range(diff):
                x.append(x[len(x)-1])  #Repeating the last value
                y.append(y[len(y)-1])  #Repeating the last value

        #Create Pandas DataFrame
        d = {'X': x, 'Y': y, 'Motion Profile X': x_mp, 'Motion Profile Y': y_mp}
        df = pd.DataFrame(data=d)
        df.to_excel(writer, sheet_name=sheet_name)

        # Access the XlsxWriter workbook and worksheet objects from the dataframe.
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Create a chart object.
        chart = workbook.add_chart({'type': 'scatter', 'subtype': 'straight'})
        # Configure the series of the chart from the dataframe data.
        max_row = len(df)
        chart.add_series({  #Real Telemetry
            'categories':     '='+sheet_name+'!$B$2:$B$' + str(max_row),
            'values':         '='+sheet_name+'!$C$2:$C$' + str(max_row),
            # 'line':           {'color': '#bf0000', 'width': 1.5, 'dash_type': 'dash'},    #Dashed line
            'line':           {'width': 1.5},
            # 'marker':         {'type': 'diamond', 'fill': {'color': '#00306b'}, 'border': {'color': '#00306b'}}
        })
        chart.add_series({  #Desired Path
            'categories':     '='+sheet_name+'!$D$2:$D$' + str(max_row),
            'values':         '='+sheet_name+'!$E$2:$E$' + str(max_row),
            # 'line':           {'color': '#bf0000', 'width': 1.5, 'dash_type': 'dash'},    #Dashed line
            'line':           {'color': '#FF9900', 'width': 2.5, 'dash_type': 'round_dot'},
        })

        # Configure the chart axes.
        if n==0:
            min_x=0
            min_y=0
            max_x=40
            max_y=80
        elif n==1:
            min_x=-72
            min_y=-8
            max_x=8
            max_y=40
        elif n==2:
            min_x=-4
            min_y=0
            max_x=44
            max_y=152

        # chart.set_x_axis({'name': 'X (in)', 'name_font':{'name':'Arial','size':12, 'bold': True}, 'num_font':  {'name': 'Arial', 'size': 12, 'bold': True, 'rotation': 0}, 'line': {'color': '#000000', 'width': 1.5}, 'major_tick_mark': 'inside', 'minor_tick_mark': 'inside'})
        # chart.set_y_axis({'name': 'Y (in)', 'name_font':{'name':'Arial','size':12, 'bold': True}, 'num_font':  {'name': 'Arial', 'size': 12, 'bold': True}, 'line': {'color': '#000000', 'width': 1.5}, 'major_tick_mark': 'inside', 'minor_tick_mark': 'inside', 'major_gridlines': {'visible': False}})
        chart.set_x_axis({'name': 'Time (ms)', 'min': min_x, 'max': max_x, 'name_font':{'name':'Arial','size':12, 'bold': True}, 'num_font':  {'name': 'Arial', 'size': 12, 'bold': True, 'rotation': 0}, 'line': {'color': '#000000', 'width': 1.5}, 'major_tick_mark': 'inside', 'minor_tick_mark': 'inside'})
        chart.set_y_axis({'name': 'Error (in)', 'min': min_y, 'max': max_y, 'name_font':{'name':'Arial','size':12, 'bold': True}, 'num_font':  {'name': 'Arial', 'size': 12, 'bold': True}, 'line': {'color': '#000000', 'width': 1.5}, 'major_tick_mark': 'inside', 'minor_tick_mark': 'inside', 'major_gridlines': {'visible': False}})
        chart.set_title({'name': 'Error\nScenario ' + str(n+1) + ', Trial ' + str(i+1), 'name_font':{'name':'Arial','size':12, 'bold': True}})
        chart.set_legend({'none': True})

        if n==0:
            chart.set_size({'width': 265, 'height': 500})
            chart.set_plotarea({
                'layout': {
                    'x':      0.35,
                    'y':      0.125,
                    'width':  0.75,
                    'height': 0.75,
                }
            })
        elif n==1:
            chart.set_size({'width': 350, 'height': 262})
            chart.set_plotarea({
                'layout': {
                    'x':      0.15,
                    'y':      0.2,
                    'width':  0.8,
                    'height': 0.7,
                }
            })
        elif n==2:
            chart.set_size({'width': 220, 'height': 700})
            chart.set_plotarea({
                'layout': {
                    'x':      0.40,
                    'y':      0.09,
                    'width':  0.73,
                    'height': 0.82,
                }
            })

        worksheet.insert_chart('F2', chart)

    #Writing avg speed
    net_avg_speed = [sum(all_speeds)/len(all_speeds)]
    # print(net_avg_speed)
    #Create Pandas DataFrame
    d = {'Scenario Average Speed (in/s)': net_avg_speed}
    df = pd.DataFrame(data=d)
    df.to_excel(writer, sheet_name="Average Speed")


    writer.save()



#Saving the MAE deviations to excel
"""
Labeling (must be done manually):
 - Top row is merged into 1 that says "Mean Average Deviation From Target Path (in)"
 - Left-most rows are merged in groups of 3. They are titled "PI(t)D(t)" and "Traditional PID"
 - 2nd left-most rows are titled, from top to bottom: Trial, Scenario 1, Scenario 2, Scenario 3, Scenario 1, Scenario 2, Scenario 3
"""
s1_max = max(max_deviations[0:10])
s2_max = max(max_deviations[10:20])
s3_max = max(max_deviations[20:30])
s1_max_t = max(max_deviations[30:40])
s2_max_t = max(max_deviations[40:50])
s3_max_t = max(max_deviations[50:60])
writer = pd.ExcelWriter("DeviationsTable_Traditional.xlsx", engine='xlsxwriter')
#Get average MAE of for all trials in each scenario
d = {'Trial 1': [mae[0], mae[11], mae[22], mae_t[0], mae_t[11], mae_t[22]], \
    'Trial 2': [mae[1], mae[12], mae[23], mae_t[1], mae_t[12], mae_t[23]], \
    'Trial 3': [mae[2], mae[13], mae[24], mae_t[2], mae_t[13], mae_t[24]], \
    'Trial 4': [mae[3], mae[14], mae[25], mae_t[3], mae_t[14], mae_t[25]], \
    'Trial 5': [mae[4], mae[15], mae[26], mae_t[4], mae_t[15], mae_t[26]], \
    'Trial 6': [mae[5], mae[16], mae[27], mae_t[5], mae_t[16], mae_t[27]], \
    'Trial 7': [mae[6], mae[17], mae[28], mae_t[6], mae_t[17], mae_t[28]], \
    'Trial 8': [mae[7], mae[18], mae[29], mae_t[7], mae_t[18], mae_t[29]], \
    'Trial 9': [mae[8], mae[19], mae[30], mae_t[8], mae_t[19], mae_t[30]], \
    'Trial 10': [mae[9], mae[20], mae[31], mae_t[9], mae_t[20], mae_t[31]], \
    'Scenario Average': [mae[10], mae[21], mae[32], mae_t[10], mae_t[21], mae_t[32]], \
    'Scenario Max': [s1_max, s2_max, s3_max, s1_max_t, s2_max_t, s3_max_t]}
df = pd.DataFrame(data=d)
df.to_excel(writer, sheet_name="Deviations")



#Plotting Count vs Deviation
num_bins = 16
#Setting up bins
smallest_dev = 99
largest_dev = 0
for i in range(60):
    if i >= 30: #Use the deviations from traditional PID instead
        dev = mae_t[i-30]
    else:
        dev = mae[i]
    if dev < smallest_dev:
        smallest_dev = round(dev, 3)
    elif dev > largest_dev:
        largest_dev = round(dev, 3)
mid = round((largest_dev+smallest_dev)/2, 3)
dev_low_bound = mid - 1.25
dev_high_bound = mid + 1.25
dev_range = dev_high_bound - dev_low_bound    #dev_range is always kept at 2.5! This is constant for all trials to presever scale in the plots.

bin_dividers = [dev_low_bound]
for i in range(num_bins):
    bin_dividers.append(bin_dividers[i] + dev_range/num_bins)
bin_dividers = [round(x, 2) for x in bin_dividers]
#Getting the list of dev intervals (all the x-coordinates for the plot)
dev_intervals = []
for i in range(num_bins):
    dev_intervals.append(str(bin_dividers[i]) + "-" + str(bin_dividers[i+1]))
print(dev_intervals)
print(bin_dividers)
print()



#Calculating Count data for PI(t)D(t)
#Getting the list of count for each dev interval (all the y-coordinate for the plot)
count = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
for i in range(30):
    bin_num = bisect(bin_dividers, mae[i])
    count[bin_num-1] += 1
print(count)

d = {"deviation intervals": dev_intervals, "count": count}

df = pd.DataFrame(data=d)
df.to_excel(writer, sheet_name="CountVsDeviation")

# Access the XlsxWriter workbook and worksheet objects from the dataframe.
workbook = writer.book
worksheet = writer.sheets["CountVsDeviation"]
# Create a chart object.
chart = workbook.add_chart({'type': 'column'})

# Configure the series of the chart from the dataframe data.
chart.add_series({
    'categories':     '='+'CountVsDeviation'+'!$B$2:$B$17',
    'values':         '='+'CountVsDeviation'+'!$C$2:$C$17',
    'line':           {'color': '#000000', 'width': 1.5},
    'fill':           {'color': '#bf0000'}
})
line_chart = workbook.add_chart({'type': 'line'})
line_chart.add_series({
    'categories':     '='+'CountVsDeviation'+'!$B$2:$B$17',
    'values':         '='+'CountVsDeviation'+'!$C$2:$C$17',
    'line':           {'color': '#700000', 'width': 1.5},
    'smooth':         True,
    'marker':         {'type': 'diamond', 'fill': {'color': '#700000'}, 'border': {'color': '#700000'}}
})
chart.combine(line_chart)
chart.set_x_axis({'name': 'Trial Mean Deviation from Target Path (in)', 'name_font':{'name':'Arial','size':12, 'bold': True}, 'num_font':  {'name': 'Arial', 'size': 9, 'bold': True}, 'line': {'color': '#000000', 'width': 1.5}, 'major_tick_mark': 'inside'})
chart.set_y_axis({'name': 'Count', 'min': 0, 'max': 16, 'name_font':{'name':'Arial','size':12, 'bold': True}, 'num_font':  {'name': 'Arial', 'size': 12, 'bold': True}, 'line': {'color': '#000000', 'width': 1.5}, 'major_tick_mark': 'inside', 'minor_tick_mark': 'inside', 'major_gridlines': {'visible': False}})
chart.set_title({'name': 'Count vs Mean Deviation\nPI(t)D(t) All Scenarios', 'name_font':{'name':'Arial','size':12, 'bold': True}})
chart.set_legend({'none': True})
chart.set_plotarea({
    'layout': {
        'x':      0.09,
        'y':      0.2,
        'width':  0.87,
        'height': 0.5,
    }
})
chart.set_size({'width': 800, 'height': 300})
worksheet.insert_chart('D2', chart)

print("---------------------------------------------------------")

#Calculating Count data for Traditional PID
#Getting the list of count for each dev interval (all the y-coordinate for the plot)
count = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
for i in range(30):
    bin_num = bisect(bin_dividers, mae_t[i])
    count[bin_num-1] += 1
print(count)

d = {"deviation intervals": dev_intervals, "count": count}

df = pd.DataFrame(data=d)
df.to_excel(writer, sheet_name="CountVsDeviation_t")

# Access the XlsxWriter workbook and worksheet objects from the dataframe.
workbook = writer.book
worksheet = writer.sheets["CountVsDeviation_t"]
# Create a chart object.
chart = workbook.add_chart({'type': 'column'})

# Configure the series of the chart from the dataframe data.
chart.add_series({
    'categories':     '='+'CountVsDeviation_t'+'!$B$2:$B$17',
    'values':         '='+'CountVsDeviation_t'+'!$C$2:$C$17',
    'line':           {'color': '#000000', 'width': 1.5},
    # 'fill':           {'color': '#bf0000'}
})
line_chart = workbook.add_chart({'type': 'line'})
line_chart.add_series({
    'categories':     '='+'CountVsDeviation_t'+'!$B$2:$B$17',
    'values':         '='+'CountVsDeviation_t'+'!$C$2:$C$17',
    'line':           {'color': '#00306b', 'width': 1.5},
    'smooth':         True,
    'marker':         {'type': 'diamond', 'fill': {'color': '#00306b'}, 'border': {'color': '#00306b'}}
})
chart.combine(line_chart)
chart.set_x_axis({'name': 'Trial Mean Deviation from Target Path (in)', 'name_font':{'name':'Arial','size':12, 'bold': True}, 'num_font':  {'name': 'Arial', 'size': 9, 'bold': True}, 'line': {'color': '#000000', 'width': 1.5}, 'major_tick_mark': 'inside'})
chart.set_y_axis({'name': 'Count', 'min': 0, 'max': 16, 'name_font':{'name':'Arial','size':12, 'bold': True}, 'num_font':  {'name': 'Arial', 'size': 12, 'bold': True}, 'line': {'color': '#000000', 'width': 1.5}, 'major_tick_mark': 'inside', 'minor_tick_mark': 'inside', 'major_gridlines': {'visible': False}})
chart.set_title({'name': 'Count vs Mean Deviation\nPID All Scenarios', 'name_font':{'name':'Arial','size':12, 'bold': True}})
chart.set_legend({'none': True})
chart.set_plotarea({
    'layout': {
        'x':      0.09,
        'y':      0.2,
        'width':  0.87,
        'height': 0.5,
    }
})
chart.set_size({'width': 800, 'height': 300})
worksheet.insert_chart('D2', chart)



writer.save()

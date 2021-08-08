import numpy as np
import matplotlib.pyplot as plt
import pandas as pd

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

#For PI(t)D(t) Scenarios 1, 2, and 3
for n in range(3):
    writer = pd.ExcelWriter("PositionTelemetry_S" + str(n+1) + ".xlsx", engine='xlsxwriter')

    for i in range(10):
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

    writer.save()




#For PI(t)D(t) Scenarios 1, 2, and 3
for n in range(3):
    writer = pd.ExcelWriter("PositionTelemetry_S" + str(n+1) + "_Traditional.xlsx", engine='xlsxwriter')

    for i in range(10):
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

    writer.save()

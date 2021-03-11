import xlwt
from xlwt import Workbook
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
from data_official_v_turn_5 import *
from MotionProfiles import *

def main():
    #These variables used for 3D plot
    global xRange
    global yRange
    xRange = [-3, 3]
    yRange = [0, 30]

    global data
    global MPData

    MPData, data = data1.split('*')
    # print(MPData)
    # print("*")
    # print(data)

    MPData = MPData.lower()
    MPData = MPData.replace("\n", "").replace("=", "").replace("{", "").replace(" ", "").replace("t", "").replace("x", "").replace("y", "").replace("h", "").replace("veloci", "").replace("acceleraion", "").replace("ang", "")
    parseMPData()

    data = data.replace("\n", "").replace("\r", "").replace(" ", "").replace("x", "").replace("y", "").replace("h", "").replace("v", "").replace("a", "").replace("ngV", "").replace("ngA", "")
    parseData()

    # tListMP = strToList(tStr.replace(" ", ""))
    # xListMP = strToList(vStr.replace(" ", ""))
    # yListMP = strToList(xStr.replace(" ", ""))
    # vListMP = strToList(yStr.replace(" ", ""))
    # aListMP = strToList(aStr.replace(" ", ""))
    # hListMP = strToList(hStr.replace(" ", ""))
    # angVListMP = strToList(angVStr.replace(" ", ""))
    # angAListMP = strToList(angAStr.replace(" ", ""))

    wb = Workbook()
    sheet = wb.add_sheet('Telemetry Data')
    styleBold = xlwt.easyxf('font: bold 1')

    #row, column
    sheet.write(0, 0, 't', styleBold)
    sheet.write(0, 1, 'x', styleBold)
    sheet.write(0, 2, 'y', styleBold)
    sheet.write(0, 3, 'h', styleBold)
    sheet.write(0, 4, 'v', styleBold)
    sheet.write(0, 5, 'a', styleBold)
    sheet.write(0, 6, 'angular v', styleBold)
    sheet.write(0, 7, 'angular a', styleBold)

    sheet.write(0, 9, 't MP', styleBold)
    sheet.write(0, 10, 'x MP', styleBold)
    sheet.write(0, 11, 'y MP', styleBold)
    sheet.write(0, 12, 'h MP', styleBold)
    sheet.write(0, 13, 'v MP', styleBold)
    sheet.write(0, 14, 'a MP', styleBold)
    sheet.write(0, 15, 'angular v MP', styleBold)
    sheet.write(0, 16, 'angular a MP', styleBold)

    style = xlwt.easyxf('font: bold off')
    for i in range(1, len(tList)):
        sheet.write(i, 0, tList[i], style)
        sheet.write(i, 1, xList[i], style)
        sheet.write(i, 2, yList[i], style)
        sheet.write(i, 3, hList[i], style)
        sheet.write(i, 4, vList[i], style)
        sheet.write(i, 5, aList[i], style)
        sheet.write(i, 6, angVList[i], style)
        sheet.write(i, 7, angAList[i], style)

    for i in range(1, len(MPtList)):
        sheet.write(i, 9, MPtList[i], style)
        sheet.write(i, 10, MPxList[i], style)
        sheet.write(i, 11, MPyList[i], style)
        sheet.write(i, 12, MPhList[i], style)
        sheet.write(i, 13, MPvList[i], style)
        sheet.write(i, 14, MPaList[i], style)
        sheet.write(i, 15, MPangVList[i], style)
        sheet.write(i, 16, MPangAList[i], style)



    wb.save("Telemetry Data.xls")

    plot3D()

def plot3D():
    fig = plt.figure()
    ax = fig.add_subplot(111, projection='3d')

    ax.scatter(tList, xList, yList, c='r', marker='o')
    ax.set_xlabel('time')
    ax.set_ylabel('x')
    ax.set_zlabel('y')

    # ax.set_ylim(xRange)
    # ax.set_zlim(yRange)

    plt.show()

def parseMPData():
    MPMasterList = MPData.split('}')
    for i in range(len(MPMasterList)):
        MPMasterList[i] = MPMasterList[i].split(',')

    global MPtList
    global MPxList
    global MPyList
    global MPhList
    global MPvList
    global MPaList
    global MPangVList
    global MPangAList
    MPtList = []
    MPxList = []
    MPyList = []
    MPhList = []
    MPvList = []
    MPaList = []
    MPangVList = []
    MPangAList = []

    for i in range(len(MPMasterList)):
        print(MPMasterList[i])
        if len(MPMasterList[i]) > 2:
            MPtList.append(MPMasterList[i][0])
            MPxList.append(MPMasterList[i][1])
            MPyList.append(MPMasterList[i][2])
            MPhList.append(MPMasterList[i][3])
            MPvList.append(MPMasterList[i][4])
            MPaList.append(MPMasterList[i][5])
            MPangVList.append(MPMasterList[i][6])
            MPangAList.append(MPMasterList[i][7])

    MPtList = [float(i) for i in MPtList]
    MPxList = [float(i) for i in MPxList]
    MPyList = [float(i) for i in MPyList]
    MPhList = [float(i) for i in MPhList]
    MPvList = [float(i) for i in MPvList]
    MPaList = [float(i) for i in MPaList]
    MPangVList = [float(i) for i in MPangVList]
    MPangAList = [float(i) for i in MPangAList]

    print(MPtList)
    print(MPxList)
    print(MPyList)
    print(MPhList)
    print(MPvList)
    print(MPaList)
    print(MPangVList)
    print(MPangAList)

def parseData():
    masterList = data.split(':')[1:]

    global tList
    global xList
    global yList
    global hList
    global vList
    global aList
    global angVList
    global angAList
    tList = masterList[0].split(',')
    while("" in tList):
        tList.remove("")
    xList = masterList[1].split(',')
    while("" in xList):
        xList.remove("")
    yList = masterList[2].split(',')
    while("" in yList):
        yList.remove("")
    hList = masterList[3].split(',')
    while("" in hList):
        hList.remove("")
    vList = masterList[4].split(',')
    while("" in vList):
        vList.remove("")
    aList = masterList[5].split(',')
    while("" in aList):
        aList.remove("")
    angVList = masterList[6].split(',')
    while("" in angVList):
        angVList.remove("")
    angAList = masterList[7].split(',')
    while("" in angAList):
        angAList.remove("")

    #Turning all strings in lists to ints
    tList = [float(i) for i in tList]
    xList = [float(i) for i in xList]
    yList = [float(i) for i in yList]
    hList = [float(i) for i in hList]
    vList = [float(i) for i in vList]
    aList = [float(i) for i in aList]
    angVList = [float(i) for i in angVList]
    angAList = [float(i) for i in angAList]

    # print(tList)
    # print(xList)
    # print(yList)
    # print(hList)
    # print(vList)
    # print(aList)
    # print(angVList)
    # print(angAList)

def strToList(str):
    list = str.split(',')
    try:
        list.remove('')
        list.remove('NaN')
    except:
        pass
    newList = [float(i) for i in list]
    return newList

if __name__ == "__main__":  #Run the main function
    main()

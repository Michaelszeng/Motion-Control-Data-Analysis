import matplotlib.pyplot as plt
from CurrentData2Battery import *

def main():
    data.replace("\n", "").replace("\r", "").replace(" ", "")
    power, current = data.split('*')
    currentAvg, currentSingle = current.split("&")
    power = power[:-2]
    currentAvg = currentAvg[:-2]
    currentSingle = currentSingle[:-2]

    print(power)
    print()
    print(currentAvg)
    print()
    print(currentSingle)
    print()

    powerListStrings = power.split(',')
    currentAvgListStrings = currentAvg.split(',')
    currentSingleListStrings = currentSingle.split(',')

    powerList = strToList(power)
    currentAvgList = strToList(currentAvg)
    currentSingleList = strToList(currentSingle)

    for power in powerList:
        power = round(power, 3)
    for currentAvg in currentAvgList:
        currentAvg = round(currentAvg, 3)
    for current in currentSingleList:
        current = round(current, 3)

    print(powerList)
    print()
    print(currentAvgList)
    print()
    print(currentSingleList)
    print()

    plt.plot(powerList, currentAvgList, marker='o', markersize=2, label='Average Current')
    plt.plot(powerList, currentSingleList, marker='o', markersize=2, label='Current')
    plt.grid()
    plt.xlabel('Input Voltage')
    plt.ylim(0.0, 1.0)
    plt.legend()
    plt.show()

def strToList(str):
    list = str.split(',')
    try:
        list.remove('')
        list.remove('NaN')
    except:
        pass

    newList = [float(i) for i in list]
    # print(list)
    return newList

if __name__ == "__main__":  #Run the main function
    main()

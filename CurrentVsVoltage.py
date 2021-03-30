import matplotlib.pyplot as plt
from CurrentData import *

data.replace("\n", "").replace("\r", "")
power, current = data.split('*')
currentAvg, currentSingle = current.split("&")

powerList = power.split(',')
currentAvgList = currentAvg.split(',')
currentSingleList = currentSingle.split(',')

print(powerList)
print()
print(currentAvgList)
print()
print(currentSingleList)
print()

import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as plticker
from bs4 import BeautifulSoup


with open('scenario1_rounded.xml', 'r') as f:
    data = f.read()

data = BeautifulSoup(data, "xml")
x_xml = data.find_all('x')
y_xml = data.find_all('y')

x1 = []
y1 = []
for val in x_xml:
    x1.append(float(val.get_text()))
for val in y_xml:
    y1.append(float(val.get_text()))


with open('scenario2_rounded.xml', 'r') as f:
    data = f.read()

data = BeautifulSoup(data, "xml")
x_xml = data.find_all('x')
y_xml = data.find_all('y')

x2 = []
y2 = []
for val in x_xml:
    x2.append(float(val.get_text()))
for val in y_xml:
    y2.append(float(val.get_text()))


with open('scenario3_rounded.xml', 'r') as f:
    data = f.read()

data = BeautifulSoup(data, "xml")
x_xml = data.find_all('x')
y_xml = data.find_all('y')

x3 = []
y3 = []
for val in x_xml:
    x3.append(float(val.get_text()))
for val in y_xml:
    y3.append(float(val.get_text()))





fig, (ax1, ax2, ax3) = plt.subplots(1, 3, figsize=(12,8), sharey=True)
ax1.scatter(x1, y1, alpha=0.5)
ax2.scatter(x2, y2, alpha=0.5)
ax3.scatter(x3, y3, alpha=0.5)

ax1.set_ylim([-2, 150])
ax1.set_xlim([-2, 60])
ax2.set_xlim([-62, 2])
ax3.set_xlim([-2, 60])

ax1.set_aspect('equal', adjustable='box')
ax2.set_aspect('equal', adjustable='box')
ax3.set_aspect('equal', adjustable='box')

# plt.scatter(x, y, alpha=0.5)
# plt.gca().set_aspect('equal', adjustable='box')
# plt.xticks(np.arange(min(x), max(x), 10))
# plt.yticks(np.arange(min(y), max(y), 10))

ax1.xaxis.set_ticks(np.arange(0, 60.001, 10))
ax1.yaxis.set_ticks(np.arange(0, 150.001, 10))
ax2.xaxis.set_ticks(np.arange(min(x2), 0.001, 10))
ax2.yaxis.set_ticks(np.arange(0, 150.001, 10))
ax3.xaxis.set_ticks(np.arange(0, 60.001, 10))
ax3.yaxis.set_ticks(np.arange(0, 150.001, 10))

ax1.tick_params(axis="y",direction="in")
ax1.tick_params(axis="x",direction="in")
ax2.tick_params(axis="y",direction="in")
ax2.tick_params(axis="x",direction="in")
ax3.tick_params(axis="y",direction="in")
ax3.tick_params(axis="x",direction="in")


# ax2.get_yaxis().set_ticks([])
# ax3.get_yaxis().set_ticks([])


ax1.spines['top'].set_visible(False)
ax1.spines['right'].set_visible(False)
# ax1.spines['bottom'].set_visible(False)
# ax1.spines['left'].set_visible(False)
ax2.spines['top'].set_visible(False)
ax2.spines['right'].set_visible(False)
# ax2.spines['bottom'].set_visible(False)
# ax2.spines['left'].set_visible(False)
ax3.spines['top'].set_visible(False)
ax3.spines['right'].set_visible(False)
# ax3.spines['bottom'].set_visible(False)
# ax3.spines['left'].set_visible(False)

plt.show()
fig.savefig('matplotlib_3scenarios.png')

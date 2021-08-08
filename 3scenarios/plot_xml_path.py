import numpy as np
import matplotlib.pyplot as plt
from bs4 import BeautifulSoup


with open('scenario3_rounded.xml', 'r') as f:
    data = f.read()

data = BeautifulSoup(data, "xml")
x_xml = data.find_all('x')
y_xml = data.find_all('y')

x = []
y = []
for val in x_xml:
    x.append(float(val.get_text()))
for val in y_xml:
    y.append(float(val.get_text()))

plt.scatter(x, y, alpha=0.5)
plt.gca().set_aspect('equal', adjustable='box')
plt.show()

# print(len(x))
# print(len(y))

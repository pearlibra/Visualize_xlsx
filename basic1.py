import pandas as pd
from matplotlib import pyplot as plt
import openpyxl
import numpy as np
import japanize_matplotlib

df1 = pd.read_excel('ポイント集計表まとめ.xlsx', sheet_name='20201021（第03週まで）', skiprows=4, header=None)
df2 = pd.read_excel('ポイント集計表まとめ.xlsx', sheet_name='20201112（第06週まで）', skiprows=[0,1,2,3,51,68], header=None)
df3 = pd.read_excel('ポイント集計表まとめ.xlsx', sheet_name='20201210（第10週まで）', skiprows=4, header=None)

df1_sub = df1.dropna(subset=[7])
df2_sub = df2.dropna(subset=[10])
df3_sub = df3.dropna(subset=[13])

basic1 = df1_sub[7].map(int)
basic2 = df2_sub[10].map(int)
basic3 = df3_sub[13].map(int)

l1 = [0,0,0,0,0,0]
l2 = [0,0,0,0,0,0]
l3 = [0,0,0,0,0,0]

for i in basic1:
    l1[i-1] += 1
for i in basic2:
    l2[i-1] += 1
for i in basic3:
    l3[i-1] += 1

l1_n = np.array([75-len(basic1)]+l1)
l2_n = np.array([73-len(basic2)]+l2)
l3_n = np.array([75-len(basic3)]+l3)

labels = ['0（未提出）', '1', '2', '3', '4', '5', '6']
x = np.arange(len(labels))
width = 0.2

fig, ax = plt.subplots()
rect1 = ax.bar(x - width, l1_n, width, label='basic1')
rect2 = ax.bar(x, l2_n, width, label='basic2')
rect3 = ax.bar(x + width, l3_n, width, label='basic3')

ax.set_ylabel('人数（人）')
ax.set_xlabel('得点（点）')
ax.set_title('各回のベーシック課題の得点')
ax.set_xticks(x)
ax.set_xticklabels(labels)
ax.legend()

def autolabel(rects):
    """Attach a text label above each bar in *rects*, displaying its height."""
    for rect in rects:
        height = rect.get_height()
        ax.annotate('{}'.format(height),
                    xy=(rect.get_x() + rect.get_width() / 2, height),
                    xytext=(0, 3),  # 3 points vertical offset
                    textcoords="offset points",
                    ha='center', va='bottom')

autolabel(rect1)
autolabel(rect2)
autolabel(rect3)

fig.tight_layout()

plt.savefig('basic1.png')

import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
import japanize_matplotlib

df11 = pd.read_excel('ポイント集計表まとめ.xlsx', sheet_name='20201216（第11週まで）', header=None, usecols=[5], skiprows=4)
df3 = pd.read_excel('ポイント集計表まとめ.xlsx', sheet_name='20201021（第03週まで）', header=None, usecols=[5], skiprows=4)

s11 = df11.iloc[:, 0]
s3 = df3.iloc[:, 0]

cor = s11.corr(s3)
text = 'cor = ' + str(round(cor, 3))

fig = plt.figure()
ax_11 = fig.add_subplot(111, title='第１１週の時点での得点分布', xlabel='合計得点(点）', ylabel='人数（人）', )
ax_11.hist(s11, bins=20)
plt.show()

fig = plt.figure()
ax_sca = fig.add_subplot(111, title='第３週の得点と第11週の得点の相関', xlabel='第３週での合計得点（点）', ylabel='第１１州での合計得点（点）')
ax_sca.scatter(s3, s11)
ax_sca.text(11, 75, text, ha='center', fontsize=18)
plt.show()


import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
import japanize_matplotlib

df10 = pd.read_excel('ポイント集計表まとめ.xlsx', sheet_name='20201210（第10週まで）',header=None, usecols=[14,15], skiprows=4)
df11 = pd.read_excel('ポイント集計表まとめ.xlsx', sheet_name='20201216（第11週まで）',header=None, usecols=[14,15], skiprows=4)

t3_1_e = df10.loc[:, 14].dropna()
t3_2_e = df10.loc[:, 15].dropna()
t3_1_l = df11.loc[:, 14].dropna()
t3_2_l = df11.loc[:, 15].dropna()

t3_e = [len(t3_1_e), len(t3_2_e)]
t3_l = [len(t3_1_l), len(t3_2_l)]

df_t3_special = pd.DataFrame({'week1':t3_e, 'week2':t3_l}, index=['スペシャル課題１', 'スペシャル課題２'])

fig = plt.figure()
ax1 = fig.add_subplot(111, title='スペシャル課題１')
ax1.pie(df_t3_special.loc['スペシャル課題１'], labeldistance=None, autopct='%1.1f%%', startangle=90, counterclock=False)
ax1.legend(['week1', 'week2'], loc='lower right')
plt.savefig('special1or2_1.png')

fig = plt.figure()
ax2 = fig.add_subplot(111, title='スペシャル課題2')
ax2.pie(df_t3_special.loc['スペシャル課題２'], labeldistance=None, autopct='%1.1f%%', startangle=90, counterclock=False)
ax2.legend(['week1', 'week2'], loc='lower right')
plt.savefig('special1or2_2.png')

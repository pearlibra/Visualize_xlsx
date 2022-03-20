import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
import japanize_matplotlib

#ノート点と合計得点のデータのみを取得する
df11 = pd.read_excel('ポイント集計表まとめ.xlsx', sheet_name='20201216（第11週まで）', header=None, usecols=[3, 5], skiprows=4)

#それぞれSeriesを作成
note = df11.loc[:, 3]
su =  df11.loc[:, 5]

#ノート点と合計得点の相関係数を計算
cor = note.corr(su)
cor_text = 'cor = ' + str(round(cor, 3))

#散布図の描画
fig = plt.figure()
ax = fig.add_subplot(111, title='ノート点と合計得点', xlabel='ノート点（点）', ylabel='合計得点（点）')
ax.scatter(note, su)
ax.text(11, 74, cor_text, ha='center', size=18)
plt.savefig('note_points.png')
#!/usr/bin/env python
# coding: utf-8

# In[5]:


get_ipython().system('pip install pywin32')


# In[19]:


#前処理

#Step1)
#Pythonを用いてネットワーク上のファイル（主にCSV）を取得し、エクセルファイルとしてローカルに保存します。
 
#例：（東京都オープンデータカタログサイトより）
#https://www.opendata.metro.tokyo.lg.jp/fukushihoken/130001_shinryoukensa20220105.csv
from pandas._libs import index
import pandas as pd
myData = pd.read_csv("https://www.opendata.metro.tokyo.lg.jp/fukushihoken/130001_shinryoukensa20220105.csv")
df = pd.DataFrame(data=myData)
    
#Step2)
#その際、検索キーが1列目になるようにPythonで加工します。
#例：1列目を削除するなど
df.to_excel("C:\\BBT\\export.xlsx", index=False)

#Step3)
#保存したエクセルにデータ接続するエクセルを作成します。
#「データ」タブの「データの取得」より接続します。
#（保存したエクセルで操作してもOKとします）

#Step4)
#別シートに単票を作成し、検索キーをVLOOKUPなどを利用し検索キーを入れると単票が出来上がるように作成します。
#例：名前と住所を表示するなど
 
#Step5)
#Pythonを用いて指定した検索キーの単票をPDFで保存します。

#検索キー入力
input_key = input("検索キーを入力: ""サンプル値「お茶の水耳鼻咽喉科・アレルギー科神田駿河台2-10-6VORT御茶ノ水2階」")

#Excelの起動
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
#ブックを開く
book = excel.Workbooks.Open("C:\\BBT\\データ接続_単票作成用.xlsx")

#シートを選択
sheet = book.WorkSheets("検索条件")
sheet.Select()

#値を代入する
sheet.Range("B1").value = input_key

#PDF出力
sheet.ExportAsFixedFormat(Type=0, Filename="C:\\BBT\\output.pdf")


# In[ ]:





import webbrowser
import random
import win32com.client
import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from sklearn.preprocessing import MinMaxScaler
from sklearn.model_selection import train_test_split


TEST_SIZE = 200

f = open('stock_list.txt', 'r')

lines = f.readlines()
stock_count = len(lines)
code = lines[round(random.random() * stock_count)]

f.close()

print('주식 목록 : {}'.format(stock_count))

print('무작위 주식 하나 투척 : {}'.format(code))

# 주식 차트 열기
# webbrowser.open('https://finance.daum.net/chart/' + code)


# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()
 
# 차트 객체 구하기
objStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
 
objStockChart.SetInputValue(0, 'A005930')   #종목 코드 - 삼성전자
objStockChart.SetInputValue(1, ord('2')) # 개수로 조회
objStockChart.SetInputValue(4, 2856) # 최근 100일 치 , 2855개 까지가 max
objStockChart.SetInputValue(5, [0,1, 2,3,4,5, 8]) #날짜,시간, 시가,고가,저가,종가,거래량
objStockChart.SetInputValue(6, ord('D')) # '차트 주가 - 일간 차트 요청
objStockChart.SetInputValue(9, ord('1')) # 수정주가 사용
objStockChart.BlockRequest()
 
len = objStockChart.GetHeaderValue(3)

#for i in range(8):
#    print(objStockChart.GetHeaderValue(i))

column = ["날짜", "시간", "시가", "고가", "저가", "종가", "거래량"]
print("날짜", "시간", "시가", "고가", "저가", "종가", "거래량")
print("-------------------------------------------")
df = pd.DataFrame()
list = []
for i in range(len):
    list.append([])
    
    day = objStockChart.GetDataValue(0, i)
    time = objStockChart.GetDataValue(1, i)
    open = objStockChart.GetDataValue(2, i)
    high = objStockChart.GetDataValue(3, i)
    low = objStockChart.GetDataValue(4, i)
    close = objStockChart.GetDataValue(5, i)
    vol = objStockChart.GetDataValue(6, i)

    
    list[i].append(day)
    list[i].append(time)
    list[i].append(open)
    list[i].append(high)
    list[i].append(low)
    list[i].append(close)
    list[i].append(vol)

    
    #print (day[i], time[i], open[i], high[i], low[i], close[i], vol[i])
    # print (day, time, open, high, low, close, vol)


df = pd.DataFrame(list, columns = column)
print(i)
print(df.head(3))

pd.to_datetime(df['날짜'], format="%Y%m%d")

df['날짜'] = pd.to_datetime(df['날짜'], format="%Y%m%d")
df['연도'] = df['날짜'].dt.year
df['월'] = df['날짜'].dt.month
df['일'] = df['날짜'].dt.day

df.to_csv('samsung_day.csv')
#df.to_csv('samsung_hour.csv')

#df = df_price.loc[df['월'] >= ]

plt.figure(figsize=(16,9))
sns.lineplot(y=df['종가'], x=df['날짜'])
plt.xlabel('time')
plt.ylabel('price')

# plt.show()
# print(df.head(3))

scaler = MinMaxScaler()
scale_cols = ['시가', '고가', '저가', '종가', '거래량']
df_scaled = scaler.fit_transform(df[scale_cols])

df_scaled = pd.DataFrame(df_scaled)
df_scaled.columns = scale_cols

print(df_scaled.head(3))

df_scaled.to_csv('scaled.csv')

train = df_scaled[:-TEST_SIZE]
test = df_scaled[-TEST_SIZE:]



print(train.shape, test.shape)
# (2656, 5) (200, 5)

def make_dataset(data, label, window_size=20):
    feature_list = []
    label_list = []
    for i in range(data.shape[0] - window_size):
        feature_list.append(np.array(data.iloc[i:i+window_size]))
        label_list.append(np.array(label.iloc[i+window_size]))
    return np.array(feature_list), np.array(label_list)

feature_cols = ['시가', '고가', '저가', '거래량']
label_cols = ['종가']

train_feature = train[feature_cols]
train_label = train[label_cols]

test_feature = test[feature_cols]
test_label = test[label_cols]

print(train_feature.shape, train_label.shape)
# (2656, 4) (2656, 1)

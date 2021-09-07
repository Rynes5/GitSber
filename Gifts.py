#
# x = ["Ручка", "Болокнот", "Полотенце"]
# y = ["Ручка"]
# print(re.findall(y[0], x[0]))
#
# i = 0
# y = [45,76,80,34,65,21,54,87, 0]
# z = [48,65,40,99,51,30,61]
# y = list(map(int, y))
# z = list(map(int, z))
# n = []
# while True:
#     if y[i] == int(0):
#         break
#     n.append(min(z, key=lambda x: abs(x - y[i])))
#     i = i + 1
# print(n)
#
#
# v0 = v["наименование"]
# v = v[v0 == "РУЧКА"]
# print(v)
#
#
# str = ['собачки', 'бегут собак', 'по полю', 'много собак', 'это собаки']
# ar = []
# subs = 'собак'
# for i in str:
#     if i.find(subs) == -1:
#         ar.append(0)
#     else:
#         ar.append(1)
# print (ar)
#
#
# v0 = v["наименование"].str.lower()
# ar = []
# subs = 'ручка'
# for i in v0:
#     if i.find(subs) == -1:
#         ar.append(0)
#     else:
#         ar.append(1)
# print (ar)
#
#
#
#
#
# ds = defaultdict(list)
# lastEl = 'space'
# nw = []
#
# for item in v:
#     for item2 in iz:
#         if item2[1].lower() in item[1].lower():
#             if item2[1].lower() not in ds:
#                 ds[item2[1].lower()] = [0, 0]
#
#             if ds[item2[1].lower()][0] == ds[item2[1].lower()][1]:
#                 ds[item2[1].lower()][0] += 1
#                 item[3] = item2[2]
#                 nw.append(item)
#                 break
#
#             ds[item2[1].lower()][1] += 1
#
# df = pd.DataFrame(nw)
# print(df)
# print(ds)
# print(iz)
#
# v.loc[v.наименование.str.lower().str.contains('блокнот')]
#
# data_tmc = pd.read_excel(r'C:\Users\Mironov1-AM\PycharmProjects\Дом\в.xlsx')
# iz = pd.read_excel(r'C:\Users\Mironov1-AM\PycharmProjects\Дом\из.xlsx')
#
#
# word = 'ручка'
# i = 0
# j = 1
# v1 = v.loc[v.наименование.str.lower().str.contains(word)]
# iz1 = iz.loc[iz.Наименование_содержание.str.lower().str.contains(word)]
# y = v1[i:j]['поставленная_стоимость']
# z = iz1['цена']
# y = list(map(int, y))
# z = list(map(int, z))
# n = []
# while True:
#     n.append(min(z, key=lambda x: abs(x - y[0])))
#     if y[0] == int(y[0]):
#         break
# iz2 = iz1[iz1["цена"] == n[0]]
# v1 = v1.reset_index(drop=True)
# iz2 = iz2.reset_index(drop=True)
# v1.loc[i, 'Номенклатура'] = iz2.loc[0, 'Номенклатура']
# v1.loc[i, 'цена'] = n[0]
# print(v1)

# import pandas as pd
# import numpy as np
# import warnings
#
# # Строит отдельный ДатаФрайм, в котором суммирует значения 2-х слов по совпадениям из файла
# warnings.filterwarnings('ignore')
#
# data_tmc = pd.read_excel(r'C:\Users\Mironov1-AM\PycharmProjects\Офис\Gifts\\iz.xlsx')
# data_crm = pd.read_excel(r'C:\Users\Mironov1-AM\PycharmProjects\Офис\Gifts\\v.xlsx')
#
# word1 = "ручка"
# data_tmc1 = data_tmc[data_tmc["Наименование_содержание"].str.lower() == word1]
# data_tmc1['pysto'] = 1
# data_tmc1 = data_tmc1[['Наименование_содержание', 'цена', 'pysto']]
#
# word2 = "блокнот"
# data_tmc2 = data_tmc[data_tmc["Наименование_содержание"].str.lower() == word2]
# data_tmc2['pysto'] = 1
# data_tmc2 = data_tmc2[['Наименование_содержание', 'цена', 'pysto']]
#
# res = data_tmc1.merge(data_tmc2, on=["pysto"])
# res['sum'] = res['цена_x'] + res['цена_y']
# del(res['pysto'])
# res['Наименование_содержание'] = res['Наименование_содержание_x'].astype(str) + " + " + res['Наименование_содержание_y']
# res = res[['Наименование_содержание', 'sum']]
#
#
# print(res)

import pandas as pd
import numpy as np

data_tmc = pd.read_excel(r'C:\Users\user\Desktop\Работа\Gifts\iz.xlsx')
text = pd.read_excel(r'C:\Users\user\Desktop\Работа\Gifts\text.xlsx')

i = 0
result = []
word = text['find_text'][0]
try:
    while True:
        data_tmc_3 = data_tmc['Наименование_содержание'][i]
        data_tmc_3 = str(data_tmc_3).lower()
        res = list(filter(lambda x: x.find(data_tmc_3)>=0, word.lower().split(' ')))
        print(res)
        if res != []:
            result.append(res)
        if data_tmc_3 == None:
            break
        i += 1
except KeyError:
    print
result = sum(result, [])
result = set(result)
result = list(result)
print(result)


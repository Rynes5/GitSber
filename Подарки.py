import pandas as pd
import numpy as np
import warnings

warnings.filterwarnings('ignore')


def get_correct_tmc(x: dict, data_tmc: pd.DataFrame) -> list:
    result = ''
# Проверка на пусто в строках

    if type(x['поставленная стоимость']) == str:
        return result

    if pd.isnull(float(x['поставленная стоимость'])):
        return result

    if pd.isnull(str(x['наименование'])):
        return result

    #     if pd.isnull(str(x['Номенклатура (для отчета)'])):
    #          return result

    #     if pd.isnull(str(data_tmc['Поисковое_значение_содержит'])):
    #          return result

    # x['наименование'] = x['наименование'].replace('+', ' ')

    word = str(x['наименование']).lower()
    i = 0
    splitting = []
    try:
        while True:
            data_tmc_3 = data_tmc['Наименование_содержание'][i]
            data_tmc_3 = str(data_tmc_3).lower()
            res = list(filter(lambda x: x.find(data_tmc_3) >= 0, word.split(' ')))
            if res != []:
                splitting.append(res)
            if data_tmc_3 == "":
                break
            i += 1
    except KeyError:
        print
    # Все засовывает в один массив
    splitting = sum(splitting, [])
    # Вытаскивает уникальные элементы
    splitting = set(splitting)
    # Делает список
    splitting = list(splitting)
    # splitting = word.split('+')
    # подумать как можно сделать это все короче... (while?)
    if len(splitting) > 1:
        data_tmc_plus_1 = data_tmc[data_tmc["Наименование_содержание"].str.lower() == splitting[0]]
        data_tmc_plus_1['pysto'] = 1
        data_tmc_plus_1 = data_tmc_plus_1[['Наименование_содержание', 'цена', 'pysto']]

        data_tmc_plus_2 = data_tmc[data_tmc["Наименование_содержание"].str.lower() == splitting[1]]
        data_tmc_plus_2['pysto'] = 1
        data_tmc_plus_2 = data_tmc_plus_2[['Наименование_содержание', 'цена', 'pysto']]

        data_tmc = data_tmc_plus_1.merge(data_tmc_plus_2, on=["pysto"])
        data_tmc['sum'] = data_tmc['цена_x'] + data_tmc['цена_y']
        data_tmc['Наименование_содержание'] = data_tmc[
            'Наименование_содержание_x'].astype(str) + " + " + data_tmc[
            'Наименование_содержание_y']
        data_tmc = data_tmc[['Наименование_содержание', 'sum']]

        data_tmc1 = data_tmc.loc[
          [True if name in word
           else False for name in data_tmc['Наименование_содержание'].str.lower()]
            ]

        data_tmc1['different_price'] = abs(data_tmc['sum'] - x['поставленная стоимость'])
        result = data_tmc1.sort_values('different_price').reset_index(drop=True)

        if len(result) > 0:
            result = result.loc[0, 'Наименование_содержание']
        else:
            return ''
    else:
        data_tmc1 = data_tmc.loc[
            [True if name.replace(' ', '') in word
             else False for name in data_tmc['Наименование_содержание'].str.lower()]
                ]

    # Проверка на "Не содержит"
    # data_crm1 = x.loc[x.наименование.str.lower().str.contains(data_tmc['Не_содержит'])]
    # if data_crm1 == True:
    #     return result

    # ind = np.argmin([abs(i - x['поставленная стоимость']) for i in data_tmc1['цена'].values])

        data_tmc1['different_price'] = abs(data_tmc1['цена'] - x['поставленная стоимость'])
        result = data_tmc1.sort_values('different_price').reset_index(drop=True)

        if len(result) > 0:
            result = result.loc[0, 'Номенклатура']
        else:
            return ''

    return result


def main():
    data_tmc = pd.read_excel(r'C:\Users\Mironov1-AM\PycharmProjects\Офис\Gifts\iz.xlsx')
    data_crm = pd.read_excel(r'C:\Users\Mironov1-AM\PycharmProjects\Офис\Gifts\v.xlsx')
    #data_crm['цена'] = 0
    #Проверка на "Не содержит"
    # data_tmc['Не_содержит'] = data_tmc['Не_содержит'].apply(lambda x: x.lower())

    # Почему от пробела зависит сработает функция или нет???
    data_crm.loc[:, 'Номенклатура'] = data_crm.apply(lambda x: get_correct_tmc(x, data_tmc), axis=1)

    #data_crm.to_excel('result.xlsx', index=False)
    print(data_crm)


if __name__ == '__main__':
    main()




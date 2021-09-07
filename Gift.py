import pandas as pd
import numpy as np
import warnings

warnings.filterwarnings('ignore')


def get_correct_tmc(x: dict, data_tmc: pd.DataFrame) -> list:
    result = ''


# Проверка на нули в строках
    if pd.isnull(x['GIFT_NAME1']):
        return result

    if type(x['GIFT_SUM1']) == str:
        return result

    if pd.isnull(x['GIFT_SUM1']):
        return result

    word = str(x['GIFT_NAME1']).lower()
    i = 0
    splitting = []
    try:
        while True:
            data_tmc_3 = data_tmc['find_text'][i]
            data_tmc_3 = str(data_tmc_3).lower()
            res = list(filter(lambda x: x.find(data_tmc_3) >= 0, word.lower().split(' ')))
            if not res:
                splitting.append(res)
            if data_tmc_3 == '':
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
    # splitting = word.split(' + ')
    # подумать как можно сделать это все короче... (while?)
    if len(splitting) > 1:
        data_tmc_plus_1 = data_tmc[data_tmc["find_text"].str.lower() == splitting[0]]
        data_tmc_plus_1['pysto'] = 1
        data_tmc1 = data_tmc_plus_1[['find_text', 'price', 'pysto']]

        data_tmc_plus_2 = data_tmc[data_tmc["find_text"].str.lower() == splitting[1]]
        data_tmc_plus_2['pysto'] = 1
        data_tmc2 = data_tmc_plus_2[['find_text', 'price', 'pysto']]

        data_tmc = data_tmc1.merge(data_tmc2, on=["pysto"])
        data_tmc['sum'] = data_tmc['price_x'] + data_tmc['price_y']
        data_tmc['find_text'] = data_tmc['find_text_x'].astype(str) + " + " + data_tmc[
            'find_text_y']
        data_tmc = data_tmc[['find_text', 'sum']]

        data_tmc1 = data_tmc.loc[
            [True if name in word
             else False for name in data_tmc['find_text'].str.lower()]
        ]

        data_tmc1['different_price'] = abs(data_tmc['sum'] - x['GIFT_SUM1'])
        result = data_tmc1.sort_values('different_price').reset_index(drop=True)

        if len(result) > 0:
            result = result.loc[0, 'find_text']
        else:
            return ''
    else:
        data_tmc1 = data_tmc.loc[
            [True if name.replace(' ', '') in word
             else False for name in data_tmc['find_text'].str.lower()]
        ]

    # добавить проверку по не содержит

    # ind = np.argmin([abs(i - x['поставленная стоимость']) for i in data_tmc1['цена'].values])

        data_tmc1['different_price'] = abs(data_tmc1['price'] - x['GIFT_SUM1'])
        result = data_tmc1.sort_values('different_price').reset_index(drop=True)

        if len(result) > 0:
            result = result.loc[0, 'short_name']
        else:
            return ''
    print(result)
    return result



def main():

    data_crm = pd.read_excel(r'C:\Users\Mironov1-AM\PycharmProjects\Офис\Gifts\CRM_cut.xlsx')
    data_tmc = pd.read_excel(r'C:\Users\Mironov1-AM\PycharmProjects\Офис\Gifts\TMC.xlsx')
    # data_crm['цена'] = 0

    data_crm.loc[:, 'Номенклатура (для отчета)'] = data_crm.apply(lambda x: get_correct_tmc(x, data_tmc), axis=1)
    # data_crm.to_excel('result.xlsx', index=False)
    print(data_crm)


if __name__ == '__main__':
    main()
# запись в дф
# res = res.append(dist, ignore_index=True)
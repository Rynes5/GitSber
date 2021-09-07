import pandas as pd
import numpy as np
import time
from pathlib import Path

# crm = pd.read_csv(r'C:\Users\Mironov1-AM\PycharmProjects\Офис\Gifts\Отчет по извинительным подаркам клиентам 25.03.2021-24.06.2021.csv', sep=';', encoding='cp1251')
# crm = pd.read_excel(r'C:\Users\Mironov1-AM\PycharmProjects\Офис\Gifts\Отчет по извинительным подаркам клиентам 25.03.2021-24.06.2021.xlsx')
crm = pd.read_excel(r'C:\Users\Mironov1-AM\PycharmProjects\Офис\Gifts\CRM_cut.xlsx')
tmc = pd.read_excel(r'C:\Users\Mironov1-AM\PycharmProjects\Офис\Gifts\TMC.xlsx')
# crm_path = Path(r'\\Braga101\Vol2\SUDR_PCP_BR\Бирюза Корп Карты\Расчеты_2021-06') / 'Отчет по извинительным подаркам клиентам 01.04.2021-30.04.2021.csv'
# crm = pd.read_csv(crm_path, sep=';', encoding='cp1251')

crm = crm.head(5000)

crm = crm[['GIFT_NAME1', 'GIFT_SUM1']]


def reduce_mem_usage(df):
    """ Перебираем все столбцы DataFrame и изменем тип, чтобы уменьшить использование памяти."""

    # start_mem = df.memory_usage().sum() / 1024 ** 2
    # print('Использование памяти DataFrame {:.2f} Mb'.format(start_mem))

    for col in df.columns:
        col_type = df[col].dtype

        if col_type != object:
            c_min = df[col].min()
            c_max = df[col].max()
            if str(col_type)[:3] == 'int':
                if c_min > np.iinfo(np.int8).min and c_max < np.iinfo(np.int8).max:
                    df[col] = df[col].astype(np.int8)
                elif c_min > np.iinfo(np.int16).min and c_max < np.iinfo(np.int16).max:
                    df[col] = df[col].astype(np.int16)
                elif c_min > np.iinfo(np.int32).min and c_max < np.iinfo(np.int32).max:
                    df[col] = df[col].astype(np.int32)
                elif c_min > np.iinfo(np.int64).min and c_max < np.iinfo(np.int64).max:
                    df[col] = df[col].astype(np.int64)
            else:
                if c_min > np.finfo(np.float32).min and c_max < np.finfo(np.float32).max:
                    df[col] = df[col].astype(np.float32)
                else:
                    df[col] = df[col].astype(np.float64)
        else:
            df[col] = df[col].astype('category')

    # end_mem = df.memory_usage().sum() / 1024 ** 2
    # print('Использовании памяти после оптимизации: {:.2f} Mb'.format(end_mem))
    # print('Уменьшился на {:.1f}%'.format(100 * (start_mem - end_mem) / start_mem))

    return df


crm = reduce_mem_usage(crm)

crm['GIFT_NAME1'] = crm['GIFT_NAME1'].str.lower()
crm['GIFT_NAME1'] = crm['GIFT_NAME1'].str.strip()

tmc['find_text'] = tmc['find_text'].str.lower()
tmc['find_text'] = tmc['find_text'].str.strip()
tmc['exclude_text'] = tmc['exclude_text'].str.lower()
tmc['exclude_text'] = tmc['exclude_text'].str.strip()

crm['GIFT_NAME1'] = crm['GIFT_NAME1'].map(lambda x: str(x).replace("ё", "е")).astype(str)
tmc['find_text'] = tmc['find_text'].map(lambda x: str(x).replace("ё", "е")).astype(str)
tmc['exclude_text'] = tmc['exclude_text'].map(lambda x: str(x).replace("ё", "е")).astype(str)

MONTHS = ['янв', 'фев', "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]

stop_words = [',', '.']

for word in stop_words:
    crm['GIFT_NAME1'] = crm['GIFT_NAME1'].str.replace(word, '', regex=False)
    tmc['find_text'] = tmc['find_text'].str.replace(word, '',  regex=False)


def levenshtein(seq1, seq2):
    size_x = len(seq1) + 1
    size_y = len(seq2) + 1
    matrix = np.zeros((size_x, size_y))
    for x in range(size_x):
        matrix[x, 0] = x
    for y in range(size_y):
        matrix[0, y] = y

    for x in range(1, size_x):
        for y in range(1, size_y):
            if seq1[x - 1] == seq2[y - 1]:
                matrix[x, y] = min(
                    matrix[x - 1, y] + 1,
                    matrix[x - 1, y - 1],
                    matrix[x, y - 1] + 1
                )
            else:
                matrix[x, y] = min(
                    matrix[x - 1, y] + 1,
                    matrix[x - 1, y - 1] + 1,
                    matrix[x, y - 1] + 1
                )
    return matrix[size_x - 1, size_y - 1]


def main_series(y,  is_levenshtein=False):
    crm_words = y.split()
    for i in range(tmc.shape[0]):
        is_find = True
        find_text = tmc.loc[i, 'find_text'].split()
        for text in find_text:
            if text not in y:
                if not is_levenshtein:
                    is_find = False
                    break
                for crm_word in crm_words:
                    leven = levenshtein(crm_word, text)
                    if leven <= 1:
                        is_find = True
                        break
                    else:
                        is_find = False
        if not pd.isnull(tmc.loc[i, 'exclude_text']):
            exclude_text = tmc.loc[i, 'exclude_text'].split()
            for text in exclude_text:
                if text in y:
                    is_find = False
                    break

        if not is_find:
            continue

    return tmp


def change_sum_value(x):
    if type(x) != str:
        return x
    if x[-3:] in MONTHS:
        x = x[:x.find('.')] + ',' + str(MONTHS.index(x[-3:]) + 1)

    result = x.replace(",", ".")
    if result.replace('.', '').isdigit():
        result = float(result)

    return result


crm['GIFT_SUM1'] = crm['GIFT_SUM1'].apply(change_sum_value)

ii = 0
number_loading_ii = len(crm['GIFT_NAME1'])


def func_find_tmc(x: dict):
    global ii
    global number_loading_ii

    crm_price = x['GIFT_SUM1']
    crm_text = x['GIFT_NAME1']

    if type(x['GIFT_SUM1']) == str or pd.isnull(x['GIFT_SUM1']) or pd.isnull(x['GIFT_NAME1']):
        ii += 1
        return [-1, '']

    tmp = tmc[tmc['price'] == -1][['price', 'short_name']]

    # tmp = main_series(crm_text, tmc, tmp, is_levenshtein=True)
    tmc_check = tmc.apply(lambda y: func_find_tmc(y, crm_text), axis=1)
    print(tmc_check)
    tmp = tmp.append(tmc.loc['price', 'short_name'], ignore_index=True)
    tmp['ones'] = 1
    tmp = tmp.reset_index()

    merge_tmc = list(tmp['short_name'].unique())
    if len(merge_tmc) == 0:
        ii += 1
        return [-1, '']

    merged_df = tmp[tmp['short_name'] == merge_tmc[0]].reset_index()
    for j, word in enumerate(merge_tmc[1:]):
        try:
            merged_df = merged_df.merge(
                tmp[tmp['short_name'] == word],
                on='ones',
                how='outer',
                suffixes=['_' + str(j), '_' + str(j + 1)]
            )
        except MemoryError:
            print(ii)
            ii += 1
            return [-1, '']

    merged_df['total_cost'] = 0
    merged_df['all_tmc'] = ''
    for column in merged_df.columns:
        if 'price' in column:
            merged_df['total_cost'] += merged_df[column]
        if 'short_name' in column:
            merged_df['all_tmc'] += '|' + merged_df[column]

    merged_df['different_price'] = abs(merged_df['total_cost'] - crm_price)
    result = merged_df.sort_values('different_price').reset_index(drop=True)

    del merged_df
    del tmp

    if len(result) == 0:
        ii += 1
        return [-1, '']

    result = result.loc[0, ['total_cost', 'all_tmc']]
    ii += 1
    loading_ii = ii / number_loading_ii * 100
    loading_ii = int(loading_ii)
    print(f'\rfunc_find_tmc - [{"█" * loading_ii}{" " * (100 - loading_ii)}] - {loading_ii * 1}%', end="")

    return pd.Series(result)


st_t = time.time()
crm[['total_cost', 'all_tmc']] = crm.apply(lambda x: func_find_tmc(x), axis=1)
crm.to_excel('result22.xlsx', index=False)
print("\n", crm)
print(time.time() - st_t)

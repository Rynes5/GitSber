import pandas as pd

crm = pd.read_excel(r'C:\Users\Mironov1-AM\PycharmProjects\Офис\Gifts\Отчет по извинительным подаркам клиентам 01.04.2021-30.04.2021.xlsx')
tmc = pd.read_excel(r'C:\Users\Mironov1-AM\PycharmProjects\Офис\Gifts\TMC.xlsx')

crm = crm.head(15)
crm['GIFT_NAME1'] = crm['GIFT_NAME1'].str.lower()
crm['GIFT_NAME1'] = crm['GIFT_NAME1'].str.strip()

tmc['find_text'] = tmc['find_text'].str.lower()
tmc['find_text'] = tmc['find_text'].str.strip()
tmc['exclude_text'] = tmc['exclude_text'].str.lower()
tmc['exclude_text'] = tmc['exclude_text'].str.strip()

stop_words = [',', '.']

for word in stop_words:
    crm['GIFT_NAME1'] = crm['GIFT_NAME1'].str.replace(word, '', regex=False)
    tmc['find_text'] = tmc['find_text'].str.replace(word, '',  regex=False)


def func_find_tmc(x: dict, is_lev=True):
    crm_text = x['GIFT_NAME1']
    print(crm_text)
    crm_price = x['GIFT_SUM1']

    if type(x['GIFT_SUM1']) == str or pd.isnull(x['GIFT_SUM1']) or pd.isnull(x['GIFT_NAME1']):
        return [-1, '']

    tmp = tmc[tmc['price'] == -1][['price', 'short_name']]

    for i in range(tmc.shape[0]):
        is_find = True
        find_text = tmc.loc[i, 'find_text'].split()
        for text in find_text:
            if text not in crm_text:
                is_find = False
                break
        if not pd.isnull(tmc.loc[i, 'exclude_text']):
            exclude_text = tmc.loc[i, 'exclude_text'].split()
            for text in exclude_text:
                if text in crm_text:
                    is_find = False
                    break

        if not is_find:
            continue

        tmp = tmp.append(tmc.loc[i, ['price', 'short_name']], ignore_index=True)

    tmp['ones'] = 1
    tmp = tmp.reset_index()

    merge_tmc = list(tmp['short_name'].unique())
    if len(merge_tmc) == 0:
        return [-1, '']

    merged_df = tmp[tmp['short_name'] == merge_tmc[0]].reset_index()

    for j, word in enumerate(merge_tmc[1:]):
        merged_df = merged_df.merge(
            tmp[tmp['short_name'] == word],
            on='ones',
            how='outer',
            suffixes=['_' + str(j), '_' + str(j + 1)]
        )

    merged_df['total_cost'] = 0
    merged_df['all_tmc'] = ''
    for column in merged_df.columns:
        if 'price' in column:
            merged_df['total_cost'] += merged_df[column]
        if 'short_name' in column:
            merged_df['all_tmc'] += ' ' + merged_df[column]

    merged_df['different_price'] = abs(merged_df['total_cost'] - crm_price)
    result = merged_df.sort_values('different_price').reset_index(drop=True)

    if len(result) == 0:
        return [-1, '']
    return pd.Series(result.loc[0, ['total_cost', 'all_tmc']])


crm[['total_cost', 'all_tmc']] = crm.apply(lambda x: func_find_tmc(x, is_lev=True), axis=1)

crm.to_excel('result.xlsx', index=False)
print(crm)




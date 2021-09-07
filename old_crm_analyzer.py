import pandas as pd


crm = pd.read_excel(r'C:\Users\Mironov1-AM\PycharmProjects\Офис\Gifts\CRM_cut.xlsx')
tmc = pd.read_excel(r'C:\Users\Mironov1-AM\PycharmProjects\Офис\Gifts\TMC.xlsx')

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


def func_find_tmc(x: dict):
    id_tmc = -1

    if type(x['GIFT_SUM1']) == str or pd.isnull(x['GIFT_SUM1']) or pd.isnull(x['GIFT_NAME1']):
        return id_tmc

    crm_text = x['GIFT_NAME1']
    crm_price = x['GIFT_SUM1']

    ids_tmc = []
    arr = []
    for i in range(tmc.shape[0]):
        is_find = True
        find_text = tmc.loc[i, 'find_text'].split()
        for text in find_text:
            if text not in crm_text:
                is_find = False
            else:
                arr.append(text)
                arr = list(set(arr))
                print(arr)
                break
        if not pd.isnull(tmc.loc[i, 'exclude_text']):
            exclude_text = tmc.loc[i, 'exclude_text'].split()
            for text in exclude_text:
                if text in crm_text:
                    is_find = False
                    break

        if not is_find:
            continue

        ids_tmc.append(i)

    min_delta = 99999999999
    for i in ids_tmc:
        price = tmc.loc[i, 'price']
        delta = abs(price - crm_price)
        if delta < min_delta:
            min_delta = delta
            id_tmc = i

    return id_tmc


crm['tmc'] = crm.apply(func_find_tmc, axis=1)

crm = pd.merge(crm, tmc[['short_name', 'price']], how='left',
               on=None, left_on='tmc', right_on=None,
               left_index=False, right_index=True, sort=False)

crm = crm.drop(['tmc'], axis=1)

crm.to_excel('result.xlsx', index=False)
print(crm)

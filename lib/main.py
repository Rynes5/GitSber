import pandas as pd

from win32com.client import Dispatch


def read_csv(path_to_file):
    with open(path_to_file, "r",
              encoding='utf-8') as file:
        data = file.readlines()

    columns = data[0].split(';')
    columns[0] = columns[0].replace('\ufeff', '')

    df = pd.DataFrame(data[1:])
    df = df.iloc[:, 0].str.replace('"', "").str.split(';', expand=True)
    df.columns = columns
    return df


def get_eso_file(path_to_file):
    return read_csv(path_to_file)


def get_config(path_to_config_file):
    directory_products = pd.read_excel(path_to_config_file, 'продукты')
    directory_code_gosb = pd.read_excel(path_to_config_file, 'Коды ГОСБ для заказа карт',
                                        usecols=['Код ГОСБ 4 знака', 'Код ГОСБ 2 знака', 'Центр эмиссии'])
    return directory_products, directory_code_gosb


def card_order(export: pd.DataFrame, directory_products: pd.DataFrame,
               directory_cod_gosb: pd.DataFrame) -> pd.DataFrame:

    export = export[export['Краткое описание'] == 'Заказ карт Momentum']
    replaced_patterns = ['Валюта:', 'Потребность (шт.):', 'Остатки карт на начало ОД:']
    stop_symbol = '█'
    for i in replaced_patterns:
        export['Описание'] = export['Описание'].map(lambda x: str(x).replace(i, f'{stop_symbol}{i}'))

    patterns = [
        f'Вид карты: (?P<card_name>[^{stop_symbol}]+) {stop_symbol}Валюта: (?P<currency>[^{stop_symbol}]+) {stop_symbol}Потребность \(шт.\): (?P<need>\d+) {stop_symbol}Остатки карт на начало ОД: (?P<balance>\d+)',
        'Причина подкрепления: (?P<cause>.+) ТБ-ГОСБ-ВСП: (?P<tb_gosb_vsp>.+)',
        '-(?P<gosb>\d+)-', '-\d+-(?P<vsp>\d+)'
    ]

    for pattern0 in patterns:
        match_0 = export['Описание'].str.extractall(pattern0)
        match_0.reset_index(inplace=True)
        export = export.merge(match_0, right_on='level_0', left_index=True, how='left')
        export = export.drop(['level_0'], axis=1)
        export = export.drop(['match'], axis=1)

    export['Наименование в ЕСО'] = export['card_name'] + ' ' + export['currency']
    direct = export.merge(directory_products, on='Наименование в ЕСО', how='left')
    direct = direct.astype({'gosb': float})
    direct = direct.merge(directory_cod_gosb, left_on='gosb', right_on='Код ГОСБ 4 знака', how='left')
    direct['дата отправки в ЦЭ'] = ''

    direct = direct[[
        'Код ГОСБ 2 знака', 'gosb', 'vsp', 'tb_gosb_vsp', 'card_name', 'Наименование в СНУИЛ', 'need', 'Центр эмиссии',
        'дата отправки в ЦЭ', 'Код', 'balance', 'cause', 'Организация', 'Время регистрации',
        'Контрольный срок', 'Внутренний Клиент',
    ]]
    direct.fillna('', inplace=True)
    direct = direct.rename(columns={'gosb': 'Код ГОСБ 4 знака',
                                    'tb_gosb_vsp': 'Номер ВСП из запроса',
                                    'vsp': 'Номер ВСП',
                                    'card_name': 'Наименование в запросе',
                                    'Наименование в СНУИЛ': 'Персопродукт',
                                    'need': 'Количество',
                                    'Код': 'Код запроса',
                                    'balance': 'Остаток в ВСП',
                                    'cause': 'Причина подкрепления'
                                    })

    return direct


def save_df_to_excel(df, file_pattern, new_name=None, sheet='ВСП', first_row=1, first_column=1, **kwargs):
    xl = Dispatch("Excel.Application")
    try:
        wb = xl.Workbooks.Open(file_pattern)
        xl.Application.Calculation = -4135
        xl.Application.ScreenUpdating = False
        xl.Application.DisplayAlerts = False
        ws = wb.Sheets(sheet)

        if first_row == -1:
            used_range = ws.UsedRange
            first_row = used_range.Row + used_range.Rows.Count

        rng = ws.Range(
            ws.Cells(first_row, first_column),
            ws.Cells(first_row + len(df) - 1, first_column + len(df.columns) - 1)
        )

        rng.Value = tuple(
            [tuple([str(y) for y in x]) for x in df.values]
        )

        xl.Application.Calculation = -4105
        xl.Application.ScreenUpdating = True

        if not new_name:
            new_name = file_pattern

        wb.SaveAs(new_name)

        xl.Application.DisplayAlerts = True
        wb.Close()
    finally:
        xl.Quit()

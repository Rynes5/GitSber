import pandas as pd
import warnings

from pathlib import Path
from win32com.client import Dispatch

warnings.filterwarnings('ignore')

PATH_ROOT = Path(__file__).parent.resolve()
PATH_ESO_FILE = PATH_ROOT / 'src/исходники/export (139).csv'
PATH_CONFIG_FILE = PATH_ROOT / 'src/Справочник продуктов кодов ГОСБ.xlsx'
PATH_HOME_FILE = PATH_ROOT / 'src/result_directory.xlsx'


def card_order(export: pd.DataFrame, directory_products: pd.DataFrame,
               directory_cod_gosb: pd.DataFrame) -> pd.DataFrame:
    export = export[export['Краткое описание'] == 'Заказ карт Momentum']
    naming = ['Валюта:', 'Потребность (шт.):', 'Остатки карт на начало ОД:']

    for i in naming:
        export['Описание'] = export['Описание'].map(lambda x: str(x).replace(i, f'█{i}'))

    arr = [
        'Вид карты: (?P<card_name>[^█]+) █Валюта: (?P<currency>[^█]+) █Потребность \(шт.\): (?P<need>\d+) █Остатки карт на начало ОД: (?P<balance>\d+)',
        'Причина подкрепления: (?P<cause>.+) ТБ-ГОСБ-ВСП: (?P<tb_gosb_vsp>.+)',
        '-(?P<gosb>\d+)-', '-\d+-(?P<vsp>\d+)'
    ]

    for pattern0 in arr:
        match_0 = export['Описание'].str.extractall(pattern0)
        match_0.reset_index(inplace=True)
        export = export.merge(match_0, right_on='level_0', left_index=True, how='left')
        export = export.drop(['level_0'], axis=1)
        export = export.drop(['match'], axis=1)

    export['Наименование в ЕСО'] = export['card_name'] + ' ' + export['currency']
    direct = export.merge(directory_products, on='Наименование в ЕСО', how='left')
    direct = direct.astype({'gosb': float})
    direct = direct.merge(directory_cod_gosb, left_on='gosb', right_on='Код ГОСБ 4 знака', how='left')

    direct = direct[[
        'Код ГОСБ 2 знака', 'gosb', 'vsp', 'Наименование в СНУИЛ', 'need', 'Центр эмиссии',
        'Время регистрации', 'Код', 'balance', 'cause', 'Организация', 'Время регистрации',
        'Контрольный срок', 'Внутренний Клиент'
    ]]
    direct = direct.rename(columns={'gosb': 'Код ГОСБ 4 знака', 'vsp': 'Номер ВСП',
                                    'Наименование в СНУИЛ': 'Персопродукт', 'need': 'Количество',
                                    'Время регистрации': 'дата отправки в ЦЭ', 'Код': 'Код запроса',
                                    'balance': 'Остаток в ВСП', 'cause': 'Причина подкрепления'
                                    })

    return direct


def save_df_to_excel(df, file_pattern, new_name=None, sheet='ВСП', first_row=1, first_column=1, **kwargs):
    xl = Dispatch("Excel.Application")
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
    xl.Quit()


def main():
    export = pd.read_csv(PATH_ESO_FILE, sep=';')
    directory_products = pd.read_excel(PATH_CONFIG_FILE, 'продукты')
    directory_cod_gosb = pd.read_excel(PATH_CONFIG_FILE, 'Коды ГОСБ для заказа карт',
                                       usecols=['Код ГОСБ 4 знака', 'Код ГОСБ 2 знака', 'Центр эмиссии'])
    direct = card_order(export, directory_products, directory_cod_gosb)
    # direct.to_excel('result_directory.xlsx', index=False)
    save_df_to_excel(direct, str(PATH_HOME_FILE), sheet='Sheet1', first_row=-1)


if __name__ == '__main__':
    main()
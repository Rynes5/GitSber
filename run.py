import warnings

from pathlib import Path
from lib.main import card_order, save_df_to_excel, get_eso_file, get_config

warnings.filterwarnings('ignore')

file = input("Введите название файла, необходимого для обработки: ")

PATH_ROOT = Path(__file__).parent.resolve()
PATH_ESO_FILE = PATH_ROOT / f'src/sources/{file}.csv'
PATH_CONFIG_FILE = PATH_ROOT / 'src/Справочник продуктов кодов ГОСБ.xlsx'
PATH_HOME_FILE = PATH_ROOT / 'заявки.xlsx'


def run():
    print('Загружаю новый файл ЕСО', end=' ')
    export = get_eso_file(PATH_ESO_FILE)
    print(f'...загружено {export.shape[0]} строк')

    directory_products, directory_code_gosb = get_config(PATH_CONFIG_FILE)

    print('Обрабатываю выгрузку')
    direct = card_order(export, directory_products, directory_code_gosb)

    save_df_to_excel(direct, str(PATH_HOME_FILE), sheet='Sheet1', first_row=-1)
    print(f'Данные добавлены в файл {PATH_HOME_FILE.name}')


if __name__ == '__main__':
    print('=' * 20)
    try:
        run()
    except Exception as err:
        print('\nУпс!.. Какая-то ошибка')
        print(err)
    print('=' * 20)
    input('Для выхода нажмите любую кнопку')

import numpy as np
import pandas as pd
from PyQt5.QtWidgets import QMessageBox
from loguru import logger


def read_file(self, file_base, file_tsd, checkbox):
    file_stock = file_base
    file_check = file_tsd
    excess_colum_list = ['Склад', 'Краткое наименование', 'Reason code', 'ТГ', 'Номер документа', 'Счет клиента',
                         'Клиент', 'Неделя', 'Способ \nдоставки']
    try:
        df_base = pd.read_excel(file_stock, skiprows=14, na_values=0).fillna(0)
        df_base = df_base[(df_base['Физические \nзапасы'] != 0)]
        for colum in excess_colum_list:
            try:
                df_base.drop(colum, axis=1, inplace=True)
            except:
                ...
    except Exception as ex:
        logger.error('Ошибка при чтении файла {}\n{}'.format(file_stock, ex))
        QMessageBox.critical(self, 'Ошибка', f'Ошибка открытия файла 6,1\n{ex}')
        self.restart()
    df_tsd = pd.DataFrame()
    for file in file_check:
        try:
            df_temp = pd.read_excel(file, usecols=['Код номенклатуры', 'Местоположение', 'Количество факт'])
            df_tsd = pd.concat([df_tsd, df_temp], ignore_index=True)
        except Exception as ex:
            logger.error('Ошибка при чтении файла {}\n{}'.format(file, ex))
            QMessageBox.critical(self, 'Ошибка', f'Ошибка открытия файла просчета\n{ex}')
            self.restart()
    try:
        df_tsd.rename(columns={'Местоположение': 'Местоположение тсд'}, inplace=True)
    except Exception as ex:
        logger.error('Ошибка при переименовании столбцов файла {}'.format(ex))
        QMessageBox.critical(self, 'Ошибка', f'Ошибка открытия файла просчета\n{ex}')
        self.restart()

    try:
        result_df = pd.merge(df_base, df_tsd, left_on=['Код \nноменклатуры', 'Местоположение'],
                             right_on=['Код номенклатуры', 'Местоположение тсд'])
    except Exception as ex:
        print(ex)
        QMessageBox.critical(self, 'Ошибка', f'Ошибка объединения файлов\n{ex}')
        self.restart()
    try:
        result_df = pd.merge(df_tsd, df_base,
                             left_on=['Код номенклатуры', 'Местоположение тсд'],
                             right_on=['Код \nноменклатуры', 'Местоположение'], how='outer')
        result_df = result_df[[
            'Код номенклатуры', 'Местоположение тсд', 'Местоположение', 'Код \nноменклатуры', 'Описание товара',
            'Физические \nзапасы', 'Передано на доставку', 'Продано', 'Зарезерви\nровано', 'Доступно',
            'Количество факт']]
        result_df['Местоположение'] = np.where(result_df['Местоположение'].isnull(),
                                               result_df['Местоположение тсд'], result_df['Местоположение'])
        result_df['Описание товара'] = np.where(result_df['Описание товара'].isnull(),
                                                'Лишний артикул на ячейке', result_df['Описание товара'])
        result_df['Код \nноменклатуры'] = np.where(result_df['Код \nноменклатуры'].isnull(),
                                                   result_df['Код номенклатуры'],
                                                   result_df['Код \nноменклатуры'])
        result_df['Физические \nзапасы'] = np.where(result_df['Физические \nзапасы'].isnull(), 0,
                                                    result_df['Физические \nзапасы'])
        result_df['Передано на доставку'] = np.where(result_df['Передано на доставку'].isnull(), 0,
                                                     result_df['Передано на доставку'])
        result_df['Продано'] = np.where(result_df['Продано'].isnull(), 0, result_df['Продано'])
        result_df['Зарезерви\nровано'] = np.where(result_df['Зарезерви\nровано'].isnull(), 0,
                                                  result_df['Зарезерви\nровано'])
        result_df['Доступно'] = np.where(result_df['Доступно'].isnull(), 0, result_df['Доступно'])
        result_df['Количество факт'] = np.where(result_df['Количество факт'].isnull(), 0, result_df['Количество факт'])
        result_df.drop(['Код номенклатуры', 'Местоположение тсд'], axis=1, inplace=True)
        result_df = result_df.assign(Разница=(result_df['Количество факт'] - result_df['Физические \nзапасы']))
        if checkbox == 2:
            result_df = result_df[result_df['Разница'] != 0]
    except Exception as ex:
        print(ex)
        QMessageBox.critical(self, 'Ошибка', f'Ошибка сверки в общем файле\n{ex}')
        self.restart()

    write_exsel(self, 'Результат', result_df)


def write_exsel(self, name, df):
    try:
        with pd.ExcelWriter(f'{self.current_dir}/{name}.xlsx', engine='xlsxwriter') as writer:
            df_sort = df.sort_values(by='Местоположение', ascending=True)
            df_colors = df_sort.style.set_properties(
                **{
                    "text-align": "left",
                    "font-weight": "bold",
                    "font-size": "14px",
                    "border": "1px solid black"
                })
            df_colors.to_excel(writer, sheet_name='сверка', index=False, na_rep='')
            worksheet = writer.sheets['сверка']
            set_column(df_sort, worksheet)
    except Exception as ex:
        logger.error('Ошибка записи результата {}'.format(ex))
        QMessageBox.critical(self, 'Ошибка', 'Ошибка записи результата {}')
        self.restart()


def set_column(df, worksheet):
    (max_row, max_col) = df.shape
    worksheet.autofilter(0, 0, max_row, max_col - 1)
    worksheet.set_column('A:B', 20)
    worksheet.set_column('C:C', 60)
    worksheet.set_column('D:J', 20)

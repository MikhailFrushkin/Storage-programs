import csv

import pandas as pd
from PyQt5.QtWidgets import QMessageBox
from loguru import logger


def read_file_cargo(self, file_dvl, file_scan, type_doc):
    res = []
    try:
        df = pd.read_excel(file_dvl, skiprows=1)

        df_ds_result = df[(df["Груз"].notnull()) &
                          (df["Груз"].str.startswith("P")) |
                          (df["Груз"].str.startswith("B"))
                          ]
        list_bu = list(df_ds_result['Маркер'].unique())
        df_ds_result = df_ds_result.drop(
            columns=['Товар', 'Кол-во товара', '№ отгрузки СУС', '№ УП', '№ ПО', 'Дата закрытия ПО'])
    except Exception as ex:
        logger.error('Ошибка при чтении файла c DVL {}\n{}'.format(file_dvl, ex))
        QMessageBox.critical(self, 'Ошибка', 'Ошибка при чтении файла c DLVA  {}\n{}'.format(file_dvl, ex))
        self.restart()
    for file in file_scan:
        try:
            if type_doc == 2:
                with open('{}'.format(file), newline='', encoding='utf-8') as csvfile:
                    reader = csv.DictReader(csvfile)
                    for row in reader:
                        res.append(row["text"])
            else:
                with open('{}'.format(file), newline='') as csvfile:
                    reader = csv.reader(csvfile)
                    for row in reader:
                        res.append(*row)
        except Exception as ex:
            logger.error('Ошибка при чтении файла сканирования {}\n{}'.format(file, ex))
            QMessageBox.critical(self, 'Ошибка', 'Ошибка при чтении файла сканирования {}\n{}'.format(file, ex))
    df_scan = pd.DataFrame(res, columns=['barcode'])

    merged_df = pd.merge(df_scan, df_ds_result, left_on=['barcode'], right_on=['Груз'], how='outer')

    df_none = merged_df[merged_df['barcode'].isnull()]
    df_none = df_none.drop(columns=['barcode'])
    df_over = merged_df[merged_df['Код товара'].isnull()]

    try:
        first_cell = ''
        if len(list_bu) > 1:
            first_cell = 'В файле DVLA несколько БЮ, выберите свой в расхождениях'
        dict_nums = {
            f'{first_cell}': ['БЮ в DLVA', 'Количество в DVLA', 'Количество отсканированных',
                              'Количество не отсканированных',
                              'Количество лишних'],
            ' ': [", ".join(map(str, list_bu)), len(df_ds_result), len(df_scan), len(df_none), len(df_over)]
        }
        df_nums = pd.DataFrame(dict_nums)
    except Exception as ex:
        logger.error(ex)

    with pd.ExcelWriter('Результат сверки R, B.xlsx', engine='xlsxwriter') as writer:
        sheets_list = [(df_nums, 'Сводка'), (df_none, 'Не отсканированные')]
        expand_columns(writer, sheets_list=sheets_list)

        mask = merged_df['Код товара'].isna()
        mask_rolling = mask.rolling(window=5, min_periods=1, center=True).sum() > 0
        result = merged_df.loc[mask | mask_rolling]
        result = result[~(result['barcode'].isna())]
        try:
            df = result.reset_index()
            df = df.drop(columns=['index'])
            check_values = df[df['Код товара'].isna()]
            check_values = list(check_values['barcode'].unique())

            df.to_excel(writer, sheet_name='Лишние R', index=False, startrow=2)
            workbook = writer.book
            worksheet = writer.sheets['Лишние R']
            red_format = writer.book.add_format({'bg_color': '#FFC7CE'})
            worksheet.merge_range('A1:J1',
                                  'Красной заливкой выделены отсканированные Эрки которых нет в документе, '
                                  'для удобства поиска показаны 2 места отсканированных перед ней и после, '
                                  'если таковые были')
            for idx, row in df.iterrows():
                if row['barcode'] in check_values:
                    worksheet.set_row(idx + 3, cell_format=red_format)
            column_max_lengths = []
            for i, column in enumerate(df.columns):
                column_length = max(df[column].astype(str).map(len).max(), len(column)) + 2
                column_max_lengths.append(column_length)
                worksheet.set_column(i, i, column_length)
        except Exception as ex:
            logger.error(ex)


def expand_columns(writer, sheets_list):
    """Расширяет ширину столбцов в DataFrame по содержимому"""
    try:
        for i in sheets_list:
            df, name = i[0], i[1]
            df.to_excel(writer, sheet_name=name, index=False)

            workbook = writer.book
            worksheet = writer.sheets[name]

            column_max_lengths = []
            for i, column in enumerate(df.columns):
                column_length = max(df[column].astype(str).map(len).max(), len(column)) + 2
                column_max_lengths.append(column_length)
                worksheet.set_column(i, i, column_length)
    except Exception as ex:
        logger.error(ex)

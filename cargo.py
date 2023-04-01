import csv

import pandas as pd
from PyQt5.QtWidgets import QMessageBox
from loguru import logger


def read_file_cargo(self, file_dvl, file_scan, type_doc):
    """Чтение файлов и выявление расхождений через пересечение множеств"""
    res = []
    try:
        df = pd.read_excel(file_dvl, skiprows=1)

        df_filter = df[(df["Груз"].notnull()) &
                       (df["Груз"].str.startswith("P")) |
                       (df["Груз"].str.startswith("B"))
                       ]
        list_r = df_filter["Груз"].tolist()
        list_bu = list(df_filter['Маркер'].unique())
    except Exception as ex:
        logger.error('Ошибка при чтении файла c DVL {}\n{}'.format(file_dvl, ex))
        QMessageBox.critical(self, 'Ошибка', 'Ошибка при чтении файла c DLVA  {}\n{}'.format(file_dvl, ex))
        self.restart()
    for file in file_scan:
        if type_doc == 2:
            try:
                with open('{}'.format(file), newline='', encoding='utf-8') as csvfile:
                    reader = csv.DictReader(csvfile)
                    for row in reader:
                        res.append(row["text"])
            except Exception as ex:
                logger.error('Ошибка при чтении файла сканирования {}\n{}'.format(file, ex))
                QMessageBox.critical(self, 'Ошибка', 'Ошибка при чтении файла сканирования {}\n{}'.format(file, ex))
        else:
            try:
                with open('{}'.format(file), newline='') as csvfile:
                    reader = csv.reader(csvfile)
                    for row in reader:
                        res.append(*row)
            except Exception as ex:
                logger.error('Ошибка при чтении файла сканирования {}\n{}'.format(file, ex))
                QMessageBox.critical(self, 'Ошибка', 'Ошибка при чтении файла сканирования {}\n{}'.format(file, ex))
    list_none = sorted(list(set(list_r).difference(set(res))))
    list_over = sorted(list(set(res).difference(set(list_r))))

    none_dict = dict()
    for i in list_none:
        try:
            df_new = df[df["Груз"] == i]
            list_new = df_new.values.tolist()[0]
            list_new[0] = int(list_new[0])
            none_dict[i] = list_new
        except Exception as ex:
            logger.error('\n{}'.format(ex))
            QMessageBox.critical(self, 'Ошибка', '\n{}'.format(ex))

    sorted_dict = dict(sorted(none_dict.items(), key=lambda item: item[1]))

    try:
        len_dlv = len(df_filter)
        len_scan = len(res)
        write_exsel(len_dlv, len_scan, sorted_dict, list_over, list_bu)
    except Exception as ex:
        logger.error('Ошибка записи в exsel\n{}'.format(ex))
        QMessageBox.critical(self, 'Ошибка', 'Ошибка записи в exsel\n{}'.format(ex))


def write_exsel(len_dlv, len_scan, dict_none, list_over, list_bu):
    """Запись расхождений в Exsel с форматированием таблицы"""
    data = {'№ п/п': [],
            'Код товара': [],
            'Товар': [],
            'Кол-во товара': [],
            'Маркер': [],
            '№ отгрузки СУС': [],
            'Номер док. Клиента': [],
            '№ УП': [],
            '№ ПО': [],
            'Груз': [],
            'Поставщик': [],
            'Дата закрытия ПО': [],
            'Упаковка': [],
            'Контейнер': [],
            'Состав груз. места': []}
    for key, value in dict_none.items():
        data['№ п/п'].append(value[0]),
        data['Код товара'].append(value[1]),
        data['Товар'].append(value[2]),
        data['Кол-во товара'].append(value[3]),
        data['Маркер'].append(value[4]),
        data['№ отгрузки СУС'].append(value[5]),
        data['Номер док. Клиента'].append(value[6]),
        data['№ УП'].append(value[7]),
        data['№ ПО'].append(value[8]),
        data['Груз'].append(value[9]),
        data['Поставщик'].append(value[10]),
        data['Дата закрытия ПО'].append(value[11]),
        data['Упаковка'].append(value[12]),
        data['Контейнер'].append(value[13]),
        data['Состав груз. места'].append(value[14])
    try:
        df = pd.DataFrame(data)
        writer = pd.ExcelWriter('Результат сверки R, B.xlsx')
        workbook = writer.book

        try:
            first_cell = ''
            if len(list_bu) > 1:
                first_cell = 'В файле DVL несколько БЮ, выберите свой в расхождениях'
            dict_nums = {
                f'{first_cell}': ['БЮ в DLVA', 'Количество в DVLA', 'Количество отсканированных',
                             'Количество не отсканированных',
                             'Количество лишних'],
                ' ': [", ".join(map(str, list_bu)), len_dlv, len_scan, len(dict_none), len(list_over)]
            }
            df_nums = pd.DataFrame(dict_nums)
        except Exception as ex:
            print(ex)

        df_nums.to_excel(writer, sheet_name='Сводка', index=False, na_rep='NaN')
        cell_format2 = workbook.add_format()
        cell_format2.set_align('center')
        cell_format2.set_bold()
        cell_format2 = workbook.add_format({'align': 'left',
                                            'valign': 'vcenter'})
        worksheet = writer.sheets['Сводка']
        worksheet.set_column('A:A', 55, cell_format2)
        worksheet.set_column('B:B', 18, cell_format2)

        df.to_excel(writer, sheet_name='Не остканированные', index=False, na_rep='NaN')
        worksheet = writer.sheets['Не остканированные']

        cell_format = workbook.add_format()
        cell_format.set_align('center')
        cell_format.set_bold()
        cell_format.set_border(1)
        cell_format = workbook.add_format({'align': 'left',
                                           'valign': 'vcenter',
                                           'border': 1})
        worksheet.set_column('A:A', 5, cell_format)
        worksheet.set_column('B:B', 15, cell_format)
        worksheet.set_column('C:C', 30, cell_format)
        worksheet.set_column('D:F', 15, cell_format)
        worksheet.set_column('G:G', 20, cell_format)
        worksheet.set_column('H:J', 10, cell_format)
        worksheet.set_column('K:K', 30, cell_format)
        worksheet.set_column('L:N', 18, cell_format)
        worksheet.set_column('O:O', 25, cell_format)

        data_over = {
            'Груз': list_over
        }
        df_marks_all = pd.DataFrame(data_over)
        df_marks_all.to_excel(writer, sheet_name='Лишнее', index=False, na_rep='NaN')
        worksheet2 = writer.sheets['Лишнее']
        worksheet2.set_column('A:A', 30, cell_format)
        writer.close()
    except Exception as ex:
        logger.debug(ex)

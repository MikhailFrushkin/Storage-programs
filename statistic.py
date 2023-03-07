import random

import pandas as pd
from PyQt5.QtWidgets import QMessageBox
from loguru import logger
from vincent.colors import brews

users = []
operations_in_doc = ['Отгрузка', 'Подбор', 'Доставка', 'Инвентаризация', 'Приемка',
                     'Внутрискладское перемещение', 'Самовывоз']
date_start = None
date_finish = None
date_list_day = []


def user_oper(self, df):
    result = dict()
    for user in users:
        sum_num_pst = 0
        result[user] = dict()
        result[user]['Пст в зал'] = 0
        result[user]['Перенос'] = 0
        result[user]['Пст в шт.'] = 0
        df = df[~(df['Тип документа'].isnull())]
        new_df = df[(df['Исполнитель'] == user) & (df['Тип документа'].
                                                   isin(operations_in_doc))]['Тип документа'].value_counts()
        if new_df.count() != 0:
            for item in operations_in_doc:
                try:
                    if item == 'Внутрискладское перемещение':
                        pst_df = df[(df['Исполнитель'] == user) & (df['Название документа'].str.contains('->')) &
                                    (df['Название документа'].str.contains('шт'))]
                        count_pst = len(pst_df)
                        result[user]['Пст в зал'] = count_pst
                        result[user]['Перенос'] = new_df[item] - count_pst

                        for i in pst_df['Название документа']:
                            try:
                                num = str(i).split('-')[2].split(',')[0].replace(' ', '').strip()
                                if num.isdigit():
                                    sum_num_pst += int(num)
                            except Exception as ex:
                                print(ex)
                                QMessageBox.critical(self, 'Ошибка',
                                                     'Ошибка колонки "Название документа"')
                                self.restart()
                        result[user]['Пст в шт.'] = sum_num_pst
                    else:
                        result[user][item] = new_df[item]
                except Exception as ex:
                    result[user][item] = 0
        else:
            result.pop(user, None)

    try:
        result_df = pd.DataFrame({
            'Табель': list(result.keys()),
            'Отгрузка': [values['Отгрузка'] for values in result.values()],
            'Подбор': [values['Подбор'] for values in result.values()],
            'Доставка': [values['Доставка'] for values in result.values()],
            'Инвентаризация': [values['Инвентаризация'] for values in result.values()],
            'Приемка': [values['Приемка'] for values in result.values()],
            'Перенос': [values['Перенос'] for values in result.values()],
            'Самовывоз': [values['Самовывоз'] for values in result.values()],
            'Пст в зал': [values['Пст в зал'] for values in result.values()],
            'Пст в шт.': [values['Пст в шт.'] for values in result.values()],
        })
        col_list = list(result_df)
        col_list.pop(0)
        col_list.pop(8)
        result_df = result_df.assign(Всего=result_df[col_list].sum(axis=1))
    except Exception as ex:
        logger.error(ex)
        QMessageBox.critical(self, 'Ошибка',
                             'Ошибка сумирования операций')
        self.restart()

    return result_df


def read_file_tsd(self, file_statistic):
    print('Чтение файла с операциями...')
    global users
    global date_start
    global date_finish
    global date_list_day
    try:
        df = pd.read_excel(file_statistic)
        users = sorted(list(df['Исполнитель'].unique()))
        df = df[df['Статус'] == 'Завершено']
        df['Время сборки'] = pd.to_timedelta(df["Время завершения"] - df["Время создания"]).astype(str)

        date_list = sorted(df['Время завершения'].to_list())
        date_start = date_list[0].date()
        date_finish = date_list[-1].date()
        date_list_day = sorted(set([i.date() for i in date_list]))

        df_dost = df[df['Название документа'].str.contains('Подбор Доставка', na=False, regex=True)]
        df_dost_s = df[df['Название документа'].str.contains('Подбор Самовывоз', na=False, regex=True)]
    except Exception as ex:
        logger.error('Ошибка при чтении файла c 6.1 {}\n{}'.format(file_statistic, ex))
        QMessageBox.critical(self, 'Ошибка',
                             'Ошибка при чтении файла c 6.1 {}\n{}'.format(file_statistic, ex))
        self.restart()

    result_df = user_oper(self, df)
    write_exsel(self, df=result_df, df_dost=df_dost, df_dost_s=df_dost_s, df_input=df)


def total_df_create(self, df, date_s, date_f):
    try:
        total_df_dict = dict()
        total_df_dict['с {} по {}'.format(date_s, date_f)] = ['Итого']
        temp_dict = dict([(i, [df[i].sum()]) for i in ['Отгрузка',
                                                       'Подбор',
                                                       'Доставка',
                                                       'Инвентаризация',
                                                       'Приемка',
                                                       'Перенос',
                                                       'Самовывоз',
                                                       'Пст в зал',
                                                       'Пст в шт.',
                                                       'Всего'
                                                       ]])
        total_df_dict.update(temp_dict)
        df_total = pd.DataFrame(total_df_dict)
        df_total.style.set_properties(
            subset=['с {} по {}'.format(date_s, date_f),
                    'Отгрузка',
                    'Подбор',
                    'Доставка',
                    'Инвентаризация',
                    'Приемка',
                    'Перенос',
                    'Пст в шт.',
                    'Пст в зал',
                    'Самовывоз',
                    'Всего'],
            **{
                "text-align": "center",
                "font-weight": "bold",
                "font-size": "18px",
                "border": "1px solid black"
            })
    except Exception as ex:
        logger.error('{}\n{}'.format(ex, df))
        QMessageBox.critical(self, 'Ошибка',
                             'Ошибка итогов')
        self.restart()
    return df_total


def table_df_create(self, df):
    df_sort = df.sort_values(by='Всего', ascending=False)
    df_colors = df_sort.style.background_gradient(axis=0,
                                                  subset=['Отгрузка', 'Подбор',
                                                          'Доставка',
                                                          'Инвентаризация', 'Приемка',
                                                          'Перенос',
                                                          'Самовывоз',
                                                          'Пст в шт.',
                                                          'Пст в зал',
                                                          'Всего'],
                                                  cmap='YlGn'
                                                  ).set_properties(
        subset=['Табель', 'Отгрузка', 'Подбор',
                'Доставка',
                'Инвентаризация', 'Приемка',
                'Перенос',
                'Пст в шт.',
                'Пст в зал',
                'Самовывоз',
                'Всего'],
        **{
            "text-align": "center",
            "font-weight": "bold",
            "font-size": "16px",
            "border": "1px solid black"
        })
    return df_sort, df_colors


def write_exsel(self, df=None, df_dost=None, df_dost_s=None, df_input=None):
    try:
        with pd.ExcelWriter('Статистика.xlsx', engine='xlsxwriter', datetime_format='DD/MM/YY HH:MM:SS') as writer:
            (max_row, max_col) = df.shape
            workbook = writer.book
            cell_format = workbook.add_format({'align': 'center', 'valign': 'top', 'font_size': 14})

            df_total = total_df_create(self, df, date_s=date_start, date_f=date_finish)
            df_total.to_excel(writer, sheet_name='Таблица', header=True, index=False, na_rep='')

            df_table = table_df_create(self, df)
            df_table[1].to_excel(writer, sheet_name='Таблица', index=False, na_rep='', startrow=3)

            worksheet = writer.sheets['Таблица']
            set_column(df_table[0], worksheet, cell_format=cell_format, num=3)

            try:
                chart = workbook.add_chart({'type': 'doughnut'})

                chart.add_series({
                    'categories': '=Таблица!B4:I4',
                    'values': '=Таблица!B2:I2',
                    'name': 'Все операции за период',
                    'border': {'color': 'black'},
                    'points': [
                        {'fill': {'color': brews['Set1'][0]}},
                        {'fill': {'color': brews['Set1'][1]}},
                        {'fill': {'color': brews['Set1'][2]}},
                        {'fill': {'color': brews['Set1'][3]}},
                        {'fill': {'color': brews['Set1'][4]}},
                        {'fill': {'color': brews['Set1'][5]}},
                        {'fill': {'color': brews['Set1'][6]}},
                        {'fill': {'color': brews['Set1'][7]}},
                    ],
                })
                worksheet.insert_chart('A{}'.format(max_row + 6), chart,
                                       {
                                           'x_scale': 0.8, 'y_scale': 0.8
                                       })
            except Exception as ex:
                logger.error(ex)
                QMessageBox.critical(self, 'Ошибка',
                                     'Ошибка записи в файл')
                self.restart()
            try:
                chart = workbook.add_chart({'type': 'column'})
                chart.add_series({
                    'categories': '=Таблица!A5:A{}'.format(max_row + 4),
                    'values': '=Таблица!K5:K{}'.format(max_row + 4),
                    'name': 'Все операции сотрудника',
                    'border': {'color': 'black'},
                    'points': [{'fill': {'color': brews['Set1'][random.randint(0, 7)]}} for _ in range(max_row - 1)],

                })
                worksheet.insert_chart('D{}'.format(max_row + 6), chart,
                                       {
                                           'x_scale': 1.95, 'y_scale': 0.8
                                       })
            except Exception as ex:
                logger.error(ex)
                QMessageBox.critical(self, 'Ошибка',
                                     'Ошибка отрисовки графика')
                self.restart()

            try:
                df_dost.style.set_properties(
                    **{
                        "text-align": "left",
                        "font-size": "14px",
                        "border": "1px solid black"
                    }).to_excel(writer, sheet_name='Подбор доставка', index=False, na_rep='')
                worksheet2 = writer.sheets['Подбор доставка']
                (max_row, max_col) = df_dost.shape
                worksheet2.autofilter(0, 0, max_row, max_col - 1)
                worksheet2.autofit()
            except Exception as ex:
                logger.error(ex)

            try:
                df_dost_s.style.set_properties(
                    **{
                        "text-align": "left",
                        "font-size": "14px",
                        "border": "1px solid black"
                    }). \
                    to_excel(writer, sheet_name='Подбор самовывоз', index=False, na_rep='')
                worksheet3 = writer.sheets['Подбор самовывоз']
                (max_row, max_col) = df_dost_s.shape
                worksheet3.autofilter(0, 0, max_row, max_col - 1)
                worksheet3.autofit()
            except Exception as ex:
                logger.error(ex)
            df_input["Время завершения"] = pd.to_datetime(df_input["Время завершения"]).dt.date
            for day in date_list_day:
                try:
                    df_temp = df_input[df_input["Время завершения"] == day]
                    df_user_for_day = user_oper(self, df_temp)

                    df_total = total_df_create(self, df_user_for_day, date_s=day, date_f=day)
                    df_total.to_excel(writer, sheet_name=f'{day}', header=True, index=False, na_rep='')

                    df_table = table_df_create(self, df_user_for_day)
                    df_table[1].to_excel(writer, sheet_name=f'{day}', index=False, na_rep='', startrow=3)

                    worksheet = writer.sheets[f'{day}']
                    set_column(df_table[0], worksheet, cell_format=cell_format, num=3)
                except Exception as ex:
                    logger.error('{} {}'.format(ex, day))

    except Exception as ex:
        logger.error('Ошибка записи результата {}'.format(ex))
        QMessageBox.critical(self, 'Ошибка',
                             'Ошибка записи результата {}'.format(ex))
        self.restart()


def set_column(df, worksheet, cell_format=None, num=0):
    (max_row, max_col) = df.shape
    worksheet.autofilter(num, 0, max_row, max_col - 1)
    worksheet.set_column('A:A', 24, cell_format)
    worksheet.set_column('B:K', 16, cell_format)



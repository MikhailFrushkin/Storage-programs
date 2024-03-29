import datetime
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
        df = pd.read_excel(file_statistic, na_values='')
        df = df[df['Статус'] == 'Завершено']
        df = df[~(df['Тип документа'].isnull())
                & ~(df['Время создания'].isnull())
                & ~(df['Исполнитель'].isnull())
                & ~(df['Время завершения'].isnull())]
        list_drop = ['ИД документа',
                     'Статус',
                     'Кем создано',
                     'Устройство',
                     'Зона выдачи заказа',
                     'Код поставщика',
                     'Название поставщика',
                     'ОМТ',
                     'Перенесено в ЖП',
                     ]
        for i in list_drop:
            try:
                df = df.drop(columns=i)
            except Exception as ex:
                logger.debug(f'нет данного поля в документе {i} {ex}')
        df.fillna(0, inplace=True)
        users = sorted(list(df['Исполнитель'].unique()))
        date_list = sorted(df['Время завершения'].to_list())
        date_start = date_list[0].date()
        date_finish = date_list[-1].date()
        date_list_day = sorted(set([i.date() for i in date_list]))
        df['Время сборки в рабочих часах'] = df.apply(
            lambda x: calculate_work_hours(self, pd.to_datetime(x['Время создания']),
                                           pd.to_datetime(x['Время завершения'])), axis=1)

        df_dost = df[df['Название документа'].str.contains('Подбор Доставка', na=False, regex=True)]
        df_dost_s = df[df['Название документа'].str.contains('Подбор Самовывоз', na=False, regex=True)]
        df_dost_dost = df[df['Тип документа'] == 'Доставка']
        show_graf(df)

    except Exception as ex:
        logger.error('Ошибка при чтении файла отслеживания заданий тсд {}\n{}'.format(file_statistic, ex))
        QMessageBox.critical(self, 'Ошибка',
                             'Ошибка при чтении файла отслеживания заданий тсд {}\n{}'.format(file_statistic, ex))
        self.restart()

    result_df = user_oper(self, df)
    write_exсel(self, df=result_df, df_dost=df_dost, df_dost_s=df_dost_s, df_dost_dost=df_dost_dost, df_input=df)


def calculate_work_hours(self, start, end):
    # Создаем список дат-временных значений между началом и концом задачи
    minutes = pd.date_range(start=start, end=end, freq='T')

    # Оставляем только рабочие часы (между 9:00 и 21:00)
    work_minutes = minutes[(minutes.hour >= self.spinBox.value()) & (minutes.hour < self.spinBox_2.value())]

    # Фильтруем только те минуты, которые находятся между началом и концом задачи
    task_minutes = work_minutes[(work_minutes >= start) & (work_minutes <= end)]

    # Вычисляем общее время в часах, минутах и секундах
    total_seconds = len(task_minutes) * 60
    total_minutes, remaining_seconds = divmod(total_seconds, 60)
    total_hours, remaining_minutes = divmod(total_minutes, 60)

    return f'{total_hours:02d}:{remaining_minutes:02d}:{remaining_seconds:02d}'


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


def table_df_create(df):
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


def write_exсel(self, df=None, df_dost=None, df_dost_s=None, df_dost_dost=None, df_input=None):
    try:
        with pd.ExcelWriter('Статистика.xlsx', engine='xlsxwriter', datetime_format='DD/MM/YY HH:MM:SS') as writer:
            (max_row, max_col) = df.shape
            workbook = writer.book
            cell_format = workbook.add_format({'align': 'center', 'valign': 'top', 'font_size': 14})

            df_total = total_df_create(self, df, date_s=date_start, date_f=date_finish)
            df_total.to_excel(writer, sheet_name='Таблица', header=True, index=False, na_rep='')

            df_table = table_df_create(df)
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
                df_dost = df_dost.reset_index()
                df_dost['Время создания'] = df_dost['Время создания'].astype('string')
                df_dost['Время завершения'] = df_dost['Время завершения'].astype('string')
                df_dost = df_dost.drop(columns=['index'])
                df_dost.to_excel(writer, sheet_name='Подбор доставка', index=False, na_rep='')

                worksheet2 = writer.sheets['Подбор доставка']
                (max_row, max_col) = df_dost.shape
                worksheet2.autofilter(0, 0, max_row, max_col - 1)
                red_format = writer.book.add_format({'bg_color': '#FFC7CE'})
                worksheet2 = writer.sheets['Подбор доставка']
                (max_row, max_col) = df_dost.shape
                worksheet2.autofilter(0, 0, max_row, max_col - 1)
                for idx, row in df_dost.iterrows():
                    time_obj = datetime.datetime.strptime(row['Время сборки в рабочих часах'], '%H:%M:%S')
                    delta_obj = datetime.timedelta(hours=time_obj.hour, minutes=time_obj.minute,
                                                   seconds=time_obj.second)
                    seconds = delta_obj.total_seconds()
                    if seconds > 3600:
                        for col_num, value in enumerate(row):
                            if col_num in [0, 1, 2, 3, 4, 5, 6]:
                                worksheet2.write(idx + 1, col_num, value, red_format)
                column_max_lengths = []
                for i, column in enumerate(df_dost.columns):
                    column_length = max(df_dost[column].astype(str).map(len).max(), len(column)) + 2
                    column_max_lengths.append(column_length)
                    worksheet2.set_column(i, i, column_length)
            except Exception as ex:
                logger.error(ex)

            try:
                red_format = writer.book.add_format({'bg_color': '#FFC7CE'})
                df_dost_s = df_dost_s.reset_index()
                df_dost_s['Время создания'] = df_dost_s['Время создания'].astype('string')
                df_dost_s['Время завершения'] = df_dost_s['Время завершения'].astype('string')
                df_dost_s = df_dost_s.drop(columns=['index'])

                df_dost_s.to_excel(writer, sheet_name='Подбор самовывоз', index=False, na_rep='')
                worksheet3 = writer.sheets['Подбор самовывоз']
                (max_row, max_col) = df_dost_s.shape
                worksheet3.autofilter(0, 0, max_row, max_col - 1)
                for idx, row in df_dost_s.iterrows():
                    time_obj = datetime.datetime.strptime(row['Время сборки в рабочих часах'], '%H:%M:%S')
                    delta_obj = datetime.timedelta(hours=time_obj.hour, minutes=time_obj.minute,
                                                   seconds=time_obj.second)
                    seconds = delta_obj.total_seconds()
                    if seconds > 1800:
                        for col_num, value in enumerate(row):
                            if col_num in [0, 1, 2, 3, 4, 5, 6]:
                                worksheet3.write(idx + 1, col_num, value, red_format)
                column_max_lengths = []
                for i, column in enumerate(df_dost_s.columns):
                    column_length = max(df_dost_s[column].astype(str).map(len).max(), len(column)) + 2
                    column_max_lengths.append(column_length)
                    worksheet3.set_column(i, i, column_length)

            except Exception as ex:
                logger.error(ex)

            df_input["Время завершения"] = pd.to_datetime(df_input["Время завершения"]).dt.date
            for day in date_list_day:
                try:
                    df_temp = df_input[df_input["Время завершения"] == day]
                    df_user_for_day = user_oper(self, df_temp)

                    df_total = total_df_create(self, df_user_for_day, date_s=day, date_f=day)
                    df_total.to_excel(writer, sheet_name=f'{day}', header=True, index=False, na_rep='')

                    df_table = table_df_create(df_user_for_day)
                    df_table[1].to_excel(writer, sheet_name=f'{day}', index=False, na_rep='', startrow=3)
                    worksheet = writer.sheets[f'{day}']
                    set_column(df_table[0], worksheet, cell_format=cell_format, num=3)
                except Exception as ex:
                    logger.error('{} {}'.format(ex, day))

            # try:
            #     time_dost_dost = time_to_operations(df_dost_dost)
            #     time_dost_s = time_to_operations(df_dost_s)
            #     time_dost = time_to_operations(df_dost)
            #
            #     QMessageBox.information(self, 'Сводка',
            #                             'Время сбора подбор самовывоз: {}\n'
            #                             'Время сбора подбор доставка: {}\n'
            #                             'Время сбора маршрутов на доставку: {}\n'
            #                             .format(time_dost_s, time_dost, time_dost_dost))
            #
            # except Exception as ex:
            #     logger.error(ex)

    except Exception as ex:
        logger.error('Ошибка записи результата {}'.format(ex))
        QMessageBox.critical(self, 'Ошибка',
                             'Ошибка записи результата {}'.format(ex))
        self.restart()


def time_to_operations(df):
    result = '00:00:00'
    try:
        df['Время сборки в рабочих часах'] = pd.to_timedelta(df['Время сборки в рабочих часах'])
        total_time = df['Время сборки в рабочих часах'].sum()
        result = '{:02d}:{:02d}:{:02d}'.format(int(total_time.total_seconds() // 3600),
                                               int(total_time.total_seconds() % 3600 // 60),
                                               int(total_time.total_seconds() % 60),
                                               )
    except Exception as ex:
        logger.error(ex)
    return result


def set_column(df, worksheet, cell_format=None, num=0):
    (max_row, max_col) = df.shape
    worksheet.autofilter(num, 0, max_row, max_col - 1)
    worksheet.set_column('A:A', 24, cell_format)
    worksheet.set_column('B:K', 16, cell_format)


def show_graf(df):
    import pandas as pd
    import matplotlib.pyplot as plt
    import numpy as np

    colors = ['b', 'g', 'r', 'c', 'm', 'y', 'k', 'w', '#FFA07A', '#BA55D3', '#00FFFF', '#FFD700', '#808080', '#800000',
              '#008080', '#FFFF00', '#FF00FF', '#800080', '#00FF00']
    df['Время создания'] = pd.to_datetime(df['Время создания'])
    df['Время завершения'] = pd.to_datetime(df['Время завершения'])

    type_doc = ['Отгрузка',
                'Подбор', 'Внутрискладское перемещение',
                'Приемка', 'Доставка', 'Локальное перемещение', 'Самовывоз', 'Инвентаризация'
                ]
    df = df[df['Тип документа'].isin(type_doc)]
    operation_types = df['Тип документа'].unique()

    # Define the working hours as a list of tuples
    working_hours = [('05:00', '17:00')]

    # Создание списка времени с шагом 1 минута в рамках рабочего времени
    time_list = pd.date_range(start=df['Время создания'].min().strftime('%Y-%m-%d 09:00'),
                              end=df['Время завершения'].max().strftime('%Y-%m-%d 21:00'),
                              freq='1T')

    # Создание объекта оси
    fig, axes = plt.subplots(nrows=len(operation_types), ncols=1, sharex=True, figsize=(16, 16))
    for i, operation_type in enumerate(operation_types):
        # Создаем фрейм с данными для текущего типа операции
        temp_df = df[df['Тип документа'] == operation_type]
        operation_list = []
        for _, row in temp_df.iterrows():
            operation_list.append((row['Время создания'], row['Время завершения']))

        active_list = []
        for t in time_list:
            active = 0
            for start, end in operation_list:
                if start <= t <= end:
                    active = 1
                    break
            active_list.append(active)

        time_array = np.array(time_list)
        active_array = np.array(active_list)
        idx = np.where(active_array == 1)[0]
        axes[i].plot_date(time_array[idx], active_array[idx], linestyle='', marker='o', color=colors[i % len(colors)],
                          label=operation_type)

        axes[i].set_ylabel(operation_type)

    # Общие параметры для всех subplots
    plt.xlabel('Время')
    fig.tight_layout()
    plt.show()



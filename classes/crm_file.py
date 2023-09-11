import re

import openpyxl
import pandas as pd
import numpy as np
from datetime import datetime
class CrmFile:
    TECH_AGENTS = ['Тихонова Елена Петровна', 'Щербец Елена Игоревна']
    TECH_PLACES = ['Кладовая', 'Машиноместо']
    FULL_FRAME_COLUMNS = ['Договор, #', '№ Договора', 'Тип договора', 'Дата заключения', 'Дата_рег договора', 'Дата раст.-я',
 'Дата АПП', 'Дата ПАПП', 'Контрагент',  'Цена_дог,руб (дубли)', 'Площадь общая по договору (дубли)', 'Цена_дог,руб(без дублей)',
 'Площадь общая по договору (без дублей)', 'Адрес дома', 'Застройка', 'Очередь', 'Помещение Тип', 'Помещение Под Тип',
 'Помещение Студ', 'Корпус Номер', 'Этаж', '№кв.', 'Площадь общая по обмерам БТИ', '# пом.', 'Помещение', 'Правообладатель Название',
 'Прав.Тип', '## плт. граф.', '### плт. факт', 'Дата_плт.', 'Сумма_плт., руб (без дублей)', 'График Платежей Задолженность (без дублей)',
 'График Платежей Дней Просрочки', 'Дата платежа факт', 'Сумма, руб']
    ADDITIONAL_FRAME_COLUMNS = ['Квартал регистрации договора', 'Год регистрации', 'Квартал_Год регистрации',
                                'Квартал расторжения договора', 'Год расторжения', 'Квартал_Год расторжения', 'Тип договора',
                                'Очередь', 'Дом', 'Учитывается(нет/да)', 'Контрагент', 'Тип контрагента (Партнер или нет)',
                                'Договор', 'Площадь', 'Сумма', 'Комментарий', 'Расторгнут?(да/нет)', 'Корректировка м.2',
                                'Корректировка тыс.руб.']
    def __init__(self, path, type_file, name, is_empty, spt_data, period):
        if is_empty:
            self.path = path
            self.type_file = type_file
            self.is_empty = is_empty
            self.sheet_name = name
            self.spt_data = spt_data
            self.period = period
        else:
            self.path = path
            self.type_file = type_file
            self.is_empty = is_empty
            self.spt_data = spt_data
            self.period = period
            self.column = 1
            self.additional_column = 37
            self.sheet_name = name
            self.df = self.load_data()
            self.full_df = self.__create_full_frame()
            self.additional_frame = self.create_additional_frame()
            self.filter_frame()

    def load_data(self):
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        raw_values = list(ws.values)
        headers = self.clear_columns(raw_values[0])
        raw_data = pd.DataFrame(raw_values[1:], columns=headers)
        raw_data.fillna('missing', inplace=True)
        raw_data = self.__change_type_to_str(raw_data)
        raw_data = self.__custom_fillna(raw_data)
        raw_data['Цена_дог,руб(без дублей)'] = pd.to_numeric(raw_data['Цена_дог,руб(без дублей)'], errors='ignore')
        return raw_data

    def filter_frame(self):
        idx = [i for i in range(len(self.df)) if not (self.df['Цена_дог,руб(без дублей)'][i] != 0 and self.df['Тип договора'][i] in ('ДДУ', 'ДКП'))]
        self.additional_frame.iloc[idx] = ''
        # for i in range(len(self.df)):
        #     if not (self.df['Цена_дог,руб(без дублей)'][i] != 0 and self.df['Тип договора'][i] in ('ДДУ', 'ДКП')):
        #         self.additional_frame.iloc[i] = ''
        #     # if not (self.df['Цена_дог,руб(без дублей)'][i] != 0):
        #     #     self.additional_frame.iloc[i] = ''


    def create_additional_frame(self):
        add_df = pd.DataFrame(columns=self.ADDITIONAL_FRAME_COLUMNS)
        add_df['Тип договора'] = self.df['Тип договора']
        if self.spt_data == '':
            add_df['Дом'] = self.df['Корпус Номер']
        else:
            house_dict = self.create_house_dict(self.spt_data)
            add_df['Дом'] = self.df['Корпус Номер'].apply(lambda x: house_dict.get(x))
        add_df['Контрагент'], add_df['Договор'] = self.df['Контрагент'], self.df['№ Договора']
        add_df = self.__fill_date_columns(add_df)

        add_df['Очередь'] = self.df['Очередь'].apply(self.__fill_queue)
        add_df['Учитывается(нет/да)'] = np.where(self.df['Контрагент'] == 'ООО "САМОЛЕТ-НЕДВИЖИМОСТЬ МСК"', 'Нет', 'Да')
        add_df['Тип контрагента (Партнер или нет)'] = np.where(self.df['Прав.Тип'] == 'Партнеры', 'Да', 'Нет')
        # add_df['Площадь'] = np.where(add_df['Учитывается(нет/да)']=='Да', self.df['Площадь общая по договору (без дублей)'], 0)
        add_df['Площадь'] = [f'=IF(AT{i + 5}="Да",M{i + 5},0)' for i in range(len(add_df))]
        add_df['Сумма'] = [f'=IF(AT{i + 5}="Да",L{i + 5},0)/1000' for i in range(len(add_df))]
        # add_df['Корректировка м.2'] = [f'=-AX{5+i}' if self.df['Контрагент'][i] in self.TECH_AGENTS
        #                                or self.df['Помещение Под Тип'][i] in self.TECH_PLACES else ''
        #                                for i in range(len(add_df))
        #                                ]
        # add_df['Корректировка тыс.руб.'] = [f'=-AY{5 + i}' if self.df['Контрагент'][i] in self.TECH_AGENTS else ''
        #                                     for i in range(len(add_df))
        #                                ]
        add_df.fillna('missing', inplace=True)
        add_df = self.__custom_fillna(add_df)
        add_df['Расторгнут?(да/нет)'] = np.where(add_df['Год расторжения'] == '', 'Нет', 'Да')
        add_df['Корректировка м.2'] = [f'=-AX{5+i}' if add_df['Расторгнут?(да/нет)'][i] == 'Да' else ''
                                       for i in range(len(add_df))
                                       ]
        add_df['Корректировка тыс.руб.'] = [f'=-AY{5 + i}' if add_df['Расторгнут?(да/нет)'][i] == 'Да' else ''
                                            for i in range(len(add_df))
                                       ]
        add_df['Комментарий'] = [f'Техническая продажа' if self.df['Контрагент'][i] in self.TECH_AGENTS else ''
                                            for i in range(len(add_df))]
        return add_df

    @staticmethod
    def create_house_dict(data_path):
        data_houses = pd.read_excel(data_path)
        houses_dict = dict(data_houses.values)
        # house_dict = dict()
        # with open(data_path, 'r', encoding='utf8') as fl:
        #     for line in fl:
        #         chars = line.rstrip().split('\t')
        #         house_dict[chars[0]] = chars[1]
        return houses_dict

    @staticmethod
    def __fill_queue(string):
        numbers = re.findall(r'\d+', string)
        if numbers != []:
            return numbers[0]
        else:
            return string
    # @staticmethod
    # def __fill_date_columns(add_df, data):
    #     for i in range(len(add_df)):
    #         if data['Дата_рег договора'][i]!= '':
    #             string = data['Дата_рег договора'][i].split(' ')[0]
    #             if '-' in data['Дата_рег договора'][i]:
    #                 add_df['Квартал регистрации договора'][i] = '1' if datetime.strptime(string,
    #                                                                                      '%Y-%m-%d').date().month <= 6 else '2'
    #                 add_df['Год регистрации'][i] = str(
    #                     datetime.strptime(string, '%Y-%m-%d').date().year)
    #                 add_df['Квартал_Год регистрации'][i] = add_df['Квартал регистрации договора'][i] + '_' + \
    #                                                        add_df['Год регистрации'][i]
    #             else:
    #                 add_df['Квартал регистрации договора'][i] = '1' if datetime.strptime(data['Дата_рег договора'][i], '%d.%m.%Y').date().month<=6 else '2'
    #                 add_df['Год регистрации'][i] = str(datetime.strptime(data['Дата_рег договора'][i], '%d.%m.%Y').date().year)
    #                 add_df['Квартал_Год регистрации'][i] = add_df['Квартал регистрации договора'][i] + '_' + add_df['Год регистрации'][i]
    #             if data['Дата раст.-я'][i]!= '':
    #                 string = data['Дата раст.-я'][i].split(' ')[0]
    #                 if '-' in data['Дата раст.-я'][i]:
    #                     add_df['Квартал расторжения договора'][i] = '1' if datetime.strptime(string,
    #                                                                                          '%Y-%m-%d').date().month <= 6 else '2'
    #                     add_df['Год расторжения'][i] = str(
    #                         datetime.strptime(string, '%Y-%m-%d').date().year)
    #                     add_df['Квартал_Год расторжения'][i] = add_df['Квартал расторжения договора'][i] + '_' + \
    #                                                            add_df['Год расторжения'][i]
    #                 else:
    #                     add_df['Квартал расторжения договора'][i] = '1' if datetime.strptime(string,
    #                                                                                          '%d.%m.%Y').date().month <= 6 else '2'
    #                     add_df['Год расторжения'][i] = str(datetime.strptime(string, '%d.%m.%Y').date().year)
    #                     add_df['Квартал_Год расторжения'][i] = add_df['Квартал расторжения договора'][i] + '_' + \
    #                                                    add_df['Год расторжения'][i]
    #     return add_df
    @staticmethod
    def get_period(string, period):
        if string != '':
            date_string = string.split(' ')[0]
            if '-' in date_string:
                pattern = '%Y-%m-%d'
            else:
                pattern = '%d.%m.%Y'
            if period == 'Полугодие':
                if datetime.strptime(date_string, pattern).date().month <= 6:
                    return '1'
                else:
                    return '2'
            elif period == 'Месяц':
                return str(datetime.strptime(date_string, pattern).date().month)
            elif period == 'Квартал':
                return str(pd.Timestamp(datetime.strptime(date_string, pattern).date()).quarter)
            else:
                return str(datetime.strptime(date_string, pattern).date().year)
        else:
            return string


    def __fill_date_columns(self, add_df):
        if self.period != 'Год':
            add_df['Квартал регистрации договора'] = self.df['Дата_рег договора'].apply(self.get_period, args=[self.period])
            add_df['Год регистрации'] = self.df['Дата_рег договора'].apply(self.get_period, args=['Год'])
            add_df['Квартал_Год регистрации'] = add_df['Квартал регистрации договора'] + '_' + \
                                                               add_df['Год регистрации']

            add_df['Квартал расторжения договора'] = self.df['Дата раст.-я'].apply(self.get_period, args=[self.period])
            add_df['Год расторжения'] = self.df['Дата раст.-я'].apply(self.get_period, args=['Год'])
            add_df['Квартал_Год расторжения'] = add_df['Квартал расторжения договора'] + '_' + \
                                                   add_df['Год расторжения']
        else:
            add_df['Квартал регистрации договора'] = ''
            add_df['Год регистрации'] = self.df['Дата_рег договора'].apply(self.get_period, args=[self.period])
            add_df['Квартал_Год регистрации'] = add_df['Год регистрации']

            add_df['Квартал расторжения договора'] = ''
            add_df['Год расторжения'] = self.df['Дата раст.-я'].apply(self.get_period, args=[self.period])
            add_df['Квартал_Год расторжения'] = add_df['Год расторжения']

        return add_df

    @staticmethod
    def clear_columns(column_list):
        column_list = list(map(lambda x: x.replace("\n", "").strip(), column_list))
        return column_list

    @staticmethod
    def __change_type_to_str(raw_df):
        columns = ['Дата заключения', 'Дата_рег договора', 'Дата раст.-я', 'Дата АПП', 'Дата ПАПП','Дата_плт.', 'Дата платежа факт']
        for col in columns:
            if col in raw_df.columns:
                raw_df[col] = raw_df[col].apply(str)
        return raw_df

    @staticmethod
    def __custom_fillna(raw_data):
        for col in raw_data.columns:
            raw_data[col] = np.where(raw_data[col]=='missing', '', raw_data[col])
        return raw_data


    def __create_full_frame(self):
        full_frame = pd.DataFrame(columns=list(map(lambda x: x.replace(' ', ''), self.FULL_FRAME_COLUMNS)))
        for col in self.df.columns:
            if col.replace(' ', '') in full_frame.columns:
                full_frame[col.replace(' ', '')] = self.df[col]
            else:
                full_frame[col.replace(' ', '')] = ''
        full_frame.fillna('', inplace=True)
        return full_frame

# # # #
# cl = CrmFile(r'C:\Users\k.burov\Desktop\Сверка CRM\Сверка CRM\Пример работы\Новое Путилково\Путилково, CRM.xlsx', 'crm', 'fdfdf', False, '')
# print(cl.df.columns)



import numpy as np
import pandas as pd
import openpyxl
from datetime import datetime
import re

class AccountFile:
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
            if spt_data != '':
                self.house_list = self.create_house_list(spt_data)
            self.column = 2
            self.additional_column = 19
            self.sheet_name = name
            self.df = self.__create_raw_dataframe()
            self.split_frames = self.__create_split_dataframe()
            self.edit_dataframe()
            self.df.reset_index(inplace=True)
            self.df.drop(['index'], axis=1, inplace=True)
            self.additional_frame = self.__create_additional_frame()
            # self.period = ['1_2023', '2_2023']
            self.period = self.get_periods()

    def get_periods(self):
        periods = sorted(list(set(self.additional_frame['Квартал_Год'].to_list())))
        return periods

    def edit_dataframe(self):
        self.df = self.df

    @staticmethod
    def create_house_list(data_path):
        data_houses = pd.read_excel(data_path)
        # оставить только уникальные элементы
        house_list = list(data_houses['ToBe'])
        # with open(data_path, 'r', encoding='utf8') as fl:
        #     for line in fl:
        #         chars = line.rstrip().split('\t')
        #         house_list.append(chars[1])
        return house_list


    def get_spt_house(self, string):
        for house in self.house_list:
            if string in house:
                return house
        return string

    @staticmethod
    def __get_init_index(raw_data):
        for i, row in enumerate(raw_data.rows):
            try:
                datetime.strptime(row[0].value, '%d.%m.%Y')
                return i
            except Exception as exp:
                continue
    @classmethod
    def __cut_raw_data(cls, raw_data):
        init_idx = cls.__get_init_index(raw_data)
        raw_data = list(raw_data.values)
        cutted_data = raw_data[init_idx:]
        cutted_data = cutted_data[:-1]
        return cutted_data

    def __create_raw_dataframe(self):
        wb = openpyxl.load_workbook(self.path, read_only=True)
        ws = wb[wb.sheetnames[0]]
        raw_df = self.__drop_na_columns(pd.DataFrame(self.__cut_raw_data(ws)))
        raw_df.drop(raw_df.columns[-1], axis=1, inplace=True)
        raw_df = self.__set_columns(raw_df, self.type_file)
        raw_df = self.__split_string_columns(raw_df)
        raw_df.drop(['Документ', 'Аналитика ДТ', 'Аналитика КТ'], axis = 1, inplace=True)
        # raw_df = self.__drop_na_columns(raw_df, 0.9, False)
        if len(raw_df.columns)<15:
            raw_df['10_Аналитика КТ'] = ''
        raw_df = raw_df[sorted(list(raw_df.columns), key=lambda x: int(x.split('_')[0]))]
        raw_df.fillna(0, inplace=True)
        raw_df = self.__change_type_to_str(raw_df)
        return raw_df
    @staticmethod
    def __change_type_to_str(raw_df):
        columns = ['11_Счет ДТ','13_Счет КТ']
        for col in columns:
            raw_df[col] = raw_df[col].apply(str)
        return raw_df
    @staticmethod
    def __drop_na_columns(raw_df, coeff = 0.7, option=True):
        if option:
            na_col = raw_df.isna().all()
            na_col = na_col[na_col==True].index.to_list()
        else:
            na_col = raw_df.isna().sum()
            na_col = na_col[(na_col / len(raw_df)) > coeff].index.to_list()
        raw_df.drop(na_col, axis=1, inplace=True)
        return raw_df


    @staticmethod
    def __set_columns(raw_df, type_file):
        if len(raw_df.columns) == 8:
            if type_file == 'AccSales':
                raw_df.columns = ['1_Период', 'Документ', 'Аналитика ДТ', 'Аналитика КТ',
                              '11_Счет ДТ', '13_Счет КТ', '14_Сумма КТ', '15_Тип операции']
                raw_df['12_Сумма ДТ'] = 0
            else:
                raw_df.columns = ['1_Период', 'Документ', 'Аналитика ДТ', 'Аналитика КТ',
                                  '11_Счет ДТ', '12_Сумма ДТ', '13_Счет КТ', '15_Тип операции']
                raw_df['14_Сумма КТ'] = 0
        else:
            raw_df.columns = ['1_Период', 'Документ', 'Аналитика ДТ', 'Аналитика КТ',
             '11_Счет ДТ', '12_Сумма ДТ', '13_Счет КТ', '14_Сумма КТ', '15_Тип операции']
        return raw_df

    def __split_string_columns(self, raw_df):
        string_columns = ['Документ', 'Аналитика ДТ', 'Аналитика КТ']
        idx = 1
        for column in string_columns:
            temp_df = raw_df[column].apply(lambda x: x.split('\n') if pd.notnull(x) else np.nan)
            temp_df = temp_df.apply(pd.Series)
            temp_df = self.__drop_na_columns(temp_df)
            temp_df = self.__drop_na_columns(temp_df, 0.99, False)
            temp_df.columns = [str(idx+i+1)+'_'+column for i in range(len(temp_df.columns))]
            raw_df = raw_df.join(temp_df)
            idx+=len(temp_df.columns)
        return raw_df

    def __create_split_dataframe(self):
        string_df = self.df.drop(['11_Счет ДТ', '12_Сумма ДТ', '13_Счет КТ', '14_Сумма КТ', '15_Тип операции'], axis=1)
        numeric_df = self.df[['11_Счет ДТ', '12_Сумма ДТ', '13_Счет КТ', '14_Сумма КТ', '15_Тип операции']]
        return (string_df, numeric_df)

    def __create_additional_frame(self):
        additional_df = pd.DataFrame(columns=['Полугодие', 'Год', 'Квартал_Год', 'Очередь', 'Дом', 'Тип', 'Контрагент',
                                              'Тип контрагента', 'Договор', 'Сумма'])
        additional_df = self.fill_sum_column(additional_df)
        self.df.fillna('', inplace=True)
        additional_df = self.__fill_date_columns(additional_df, self.df, self.period)
        additional_df = self.fill_queue_and_house(additional_df)
        additional_df = self.fill_type_column(additional_df)
        if self.type_file == 'AccPay':
            additional_df['Контрагент'] = np.where(self.df['12_Сумма ДТ'] != 0, self.df['4_Аналитика ДТ'], self.df['7_Аналитика КТ'])
        else:
            additional_df['Контрагент'] = self.df['4_Аналитика ДТ']
        additional_df['Тип контрагента'] = ''
        additional_df['Договор'] = self.df['5_Аналитика ДТ']
        return additional_df
        # additional_df['Сумма'] = self.__fill_sum_column(additional_df, self.df)


    def fill_sum_column(self, add_df):
        add_df['Сумма'] = (self.df['12_Сумма ДТ'] - self.df['14_Сумма КТ']) / 1000
        return add_df

    def fill_type_column(self, add_df):
        add_df['Тип'] = 'Прочее движение'
        add_df.loc[self.df['11_Счет ДТ'] == '51', 'Тип'] = 'Поступление ДС'
        add_df.loc[self.df['13_Счет КТ'] == '51', 'Тип'] = 'Возврат ДС'
        add_df.loc[self.df['3_Документ'].str.lower().str.contains('заключение|изменение условий', na=False), 'Тип'] = 'Заключение'
        add_df.loc[self.df['3_Документ'].str.lower().str.contains('расторжение|аннулирование', na=False ), 'Тип'] = 'Расторжение'
        return add_df

    @staticmethod
    def __fill_date_columns(add_df, data, period):
        if period == 'Полугодие':
            add_df['Полугодие'] = data['1_Период'].apply(lambda x: '1' if datetime.strptime(x, '%d.%m.%Y').date().month<=6 else '2')
            add_df['Год'] = data['1_Период'].apply(lambda x: str(datetime.strptime(x, '%d.%m.%Y').date().year))
            add_df['Квартал_Год'] = add_df['Полугодие'] + '_' + add_df['Год']
        elif period == 'Месяц':
            add_df['Полугодие'] = data['1_Период'].apply(lambda x: str(datetime.strptime(x, '%d.%m.%Y').date().month))
            add_df['Год'] = data['1_Период'].apply(lambda x: str(datetime.strptime(x, '%d.%m.%Y').date().year))
            add_df['Квартал_Год'] = add_df['Полугодие'] + '_' + add_df['Год']
        elif period == 'Год':
            add_df['Полугодие'] = ''
            add_df['Год'] = data['1_Период'].apply(lambda x: str(datetime.strptime(x, '%d.%m.%Y').date().year))
            add_df['Квартал_Год'] = add_df['Год']
        else:
            add_df['Полугодие'] = data['1_Период'].apply(lambda x: str(pd.Timestamp(datetime.strptime(x, '%d.%m.%Y').date()).quarter))
            add_df['Год'] = data['1_Период'].apply(lambda x: str(datetime.strptime(x, '%d.%m.%Y').date().year))
            add_df['Квартал_Год'] = add_df['Полугодие'] + '_' + add_df['Год']
        return add_df


    def fill_queue_and_house(self, add_df):

        add_df['Очередь'] = [self.get_queue(self.df['9_Аналитика КТ'][i]) if self.df['12_Сумма ДТ'][i]==0
                                 else self.get_queue(self.df['6_Аналитика ДТ'][i]) for i in range(len(self.df))]
        if self.spt_data == '':
            add_df['Дом'] = [self.get_house(self.df['9_Аналитика КТ'][i]) if self.df['12_Сумма ДТ'][i] == 0
                                 else self.get_house(self.df['6_Аналитика ДТ'][i]) for i in range(len(self.df))]
        else:

            add_df['Дом'] = [self.get_spt_house(self.df['9_Аналитика КТ'][i].split(' ')[-1]) if self.df['12_Сумма ДТ'][i] == 0
                             else self.get_spt_house(self.df['6_Аналитика ДТ'][i].split(' ')[-1]) for i in range(len(self.df))]

        return add_df

    @staticmethod
    def get_queue(string):
        if isinstance(string, str):
            numbers = re.findall(r'\d+.?\d*', string)
            if len(numbers)>=2:
                return numbers[0]
            else:
                return string
        else:
            return string
    @staticmethod
    def get_house(string):
        if isinstance(string, str):
            numbers = re.findall(r'\d+.?\d*', string)
            if len(numbers)>=2:
                return numbers[-1]
            else:
                return string

class AccountPayment(AccountFile):
    def __init__(self, path, type_file, name, is_empty, spt_data, period):
        super().__init__(path, type_file, name, is_empty, spt_data, period)

class AccountSales(AccountFile):
    DOCUMNETS = ['реализация', 'корректировка реализации', 'отчет комитенту', 'передача']
    NOMENCLATURE = ['аренда', "вознаграждение агента", 'услуги связи', 'сертификат']
    def __init__(self, path, type_file, name, is_empty, spt_data, period):
        super().__init__(path, type_file, name, is_empty, spt_data, period)

    def fill_queue_and_house(self, add_df):
        add_df['Очередь'] = self.df['7_Аналитика КТ'].apply(self.get_queue)
        if self.spt_data == '':
            add_df['Дом'] = self.df['7_Аналитика КТ'].apply(self.get_house)
        else:
            add_df['Дом'] = [self.get_spt_house(self.df['7_Аналитика КТ'][i].split(' ')[-1]) for i in range(len(self.df))]
        return add_df

    def edit_dataframe(self):
        index_names = self.df[(self.df['11_Счет ДТ'].str.startswith("62"))
                              & (self.df['13_Счет КТ'].str.startswith("90"))
                              & (self.df['2_Документ'].str.lower().str.contains('|'.join(self.DOCUMNETS), na=False))
                              & (~self.df['9_Аналитика КТ'].str.lower().str.contains('|'.join(self.NOMENCLATURE), na=False))
                              & (~self.df['7_Аналитика КТ'].str.lower().str.contains('|'.join(self.NOMENCLATURE), na=False))
                              & (~self.df['8_Аналитика КТ'].str.lower().str.contains('|'.join(self.NOMENCLATURE), na=False))].index
        self.df = self.df.iloc[index_names]


    def fill_sum_column(self, add_df):
        add_df['Сумма'] = (-1)*round((self.df['12_Сумма ДТ'] - self.df['14_Сумма КТ']) / 1000,2)
        return add_df

    def fill_type_column(self, add_df):
        add_df['Тип'] = 'Заключение'
        return add_df


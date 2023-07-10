import pandas as pd
import re
import string
import win32com.client
from collections import namedtuple
from classes.crm_file import CrmFile
from classes.excel_file import AccountSales, AccountPayment

RowsDict = {'По данным 1С': 6, 'По данным CRM': 5, 'Разница': 4,
            'Продажи кв.м. (накопительный итог) без учета ВГО и дополнительных корректировок': 6,
            'Корректировка кв.м.': 5,
            'Продажи кв.м. (накопительный итог) с учетом ВГО и дополнительных корректировок': 5,
            'Продажи тыс. руб. (накопительный итог) без учета ВГО и дополнительных корректировок': 6,
            'Корректировка тыс.руб.': 5,
            'Продажи тыс. руб. (накопительный итог) с учетом ВГО и дополнительных корректировок': 5,
            'Партнерские продажи, кв. м.': 5, 'Партнерские продажи, тыс. руб.': 5, }

EditFormulasDict = {'Продажи кв.м. (накопительный итог) без учета ВГО и дополнительных корректировок': 7,
            'Продажи кв.м. (накопительный итог) с учетом ВГО и дополнительных корректировок': 6,
            'Продажи тыс. руб. (накопительный итог) без учета ВГО и дополнительных корректировок': 7,
            'Продажи тыс. руб. (накопительный итог) с учетом ВГО и дополнительных корректировок': 6,
            'Партнерские продажи, кв. м.': 6, 'Партнерские продажи, тыс. руб.': 6,
}


def add_rows(ws, data_path):
    EndRow = get_last_row_from_column(ws, 'B', True)
    new_data = pd.read_excel(data_path)
    for key, value in RowsDict.items():
        print(key)
        Start = get_split_row(ws, EndRow, key) + value
        End = get_split_row(ws, EndRow, 'ИТОГО', Start) - 1
        EndRow += len(new_data)
        for _ in range(len(new_data)):
            ws.Rows(End).Insert(-4121)
            ws.Rows(End - 1).Copy()
            ws.Rows(End).PasteSpecial(-4123)
            ws.Rows(End).PasteSpecial(-4122)
            End += 1

        ws.Range(ws.Cells(End - len(new_data), 3),  # Cell to start the "paste"
                    ws.Cells(End - len(new_data) + len(new_data.index) - 1,
                                3 + len(new_data.columns) - 1)  # No -1 for the index
                    ).Value = new_data.values

def get_excel_range(number):
    alphabet = string.ascii_uppercase
    if isinstance(number, (int, float)):
        if number>=26:
            first_letter = alphabet[(number//26)-1]
            letter = f'{first_letter}{alphabet[number%26]}'
        else:
            letter = alphabet[number]
        return letter
    else:
        alpha_list = list(number)
        if len(number)>1:
            number = (alphabet.find(alpha_list[0])+1)*26 + (alphabet.find(alpha_list[1]) + 1)
        else:
            number = alphabet.find(alpha_list[0])+1
        return number

def get_split_row(ws, EndRow, text = '', StartRow = 1, option = True):
    split_row = 0
    if option:
        for i in range(StartRow, EndRow+1):
            if ws.Range(f'B{i}').Value == text: # 'СВЕРКА 1С и CRM'
                split_row = i
                break
    else:
        for i in range(StartRow, EndRow+1):
            if re.match(r'\d_\d{4}', str(ws.Range(f'E{i}').Value)) != None \
                    and ws.Range(f'E{i}').Value == re.match(r'\d_\d{4}', str(ws.Range(f'E{i}').Value)).group():
                split_row = i - StartRow
                break
    return split_row

def get_split_col(ws, SearchRow, StartRow = 1, EndRow = 200):
    value_list = ws.Range(f'{SearchRow}:{SearchRow}').Value[0]
    res_col = 0
    for i in range(StartRow, EndRow):
        if value_list[i] == None:
            res_col = i
            break
    return res_col
def get_last_row_from_column(ws, column='', option = True):
    if option:
        max_end_row = ws.Range("{0}{1}".format(column, ws.Rows.Count)).End(-4162).Row
    else:
        columns = list(string.ascii_uppercase)
        columns.extend(['AA', 'AB'])
        max_end_row = max([ws.Range("{0}{1}".format(col, ws.Rows.Count)).End(-4162).Row for col in columns])
    return max_end_row

def get_last_num(ws):
    EndRow = get_last_row_from_column(ws, 'Q', True)
    last_num = ws.Cells(EndRow, 17).Value
    if isinstance(last_num, (int, float)):
        return last_num
    else:
        return 0
def create_sheet_dict(wb):
    user_data = namedtuple('UserData', ['sheetname', 'object'])
    sheet_dict = dict()
    for sheet in wb.Sheets:
        if 'ДДУ' in sheet.Name:
            sheet_dict['AccPay'] = user_data(sheet.Name, AccountPayment)
        elif 'ДКП' in sheet.Name:
            sheet_dict['AccSales'] = user_data(sheet.Name, AccountSales)
        elif 'CRM' in sheet.Name:
            sheet_dict['CRM'] = user_data(sheet.Name, CrmFile)
    return sheet_dict

def create_cumsum_column(sheet, df):
    last_num = get_last_num(sheet)
    df['cum_sum'] = df['12_Сумма ДТ'] - df['14_Сумма КТ']
    df['cum_sum'] = df['cum_sum'].cumsum()
    df['cum_sum'] = df['cum_sum'] + last_num
    return df


def create_custom_range(ws, project, range_type, EndRow, SplitRow):
    custom_range = namedtuple('range', ['type', 'copy', 'past'])
    if range_type in ('AccPay', "AccSales"):
        prj_row = get_split_row(ws, SplitRow, project)
        first_col = get_split_col(ws, prj_row, 1, 100)
        last_col = get_split_col(ws, prj_row, first_col + 1, 100) + 1
        if range_type == 'AccPay':
            copy_range = f'{get_excel_range(first_col-1)}1:{get_excel_range(first_col-1)}{SplitRow}'
            past_range = f'{get_excel_range(first_col)}1:{get_excel_range(first_col)}{SplitRow}'
        else:
            copy_range = f'{get_excel_range(last_col-2)}1:{get_excel_range(last_col-2)}{SplitRow}'
            past_range = f'{get_excel_range(last_col-1)}1:{get_excel_range(last_col-1)}{SplitRow}'
    else:
        prj_row = get_split_row(ws, EndRow, 'ИТОГО', SplitRow + 2)
        last_col = get_split_col(ws, prj_row, 5)

        copy_range = f'{get_excel_range(last_col-10)}{SplitRow+4}:{get_excel_range(last_col-1)}{EndRow}'
        past_range = f'{get_excel_range(last_col)}{SplitRow+4}:{get_excel_range(last_col+9)}{EndRow}'

    return custom_range(range_type, copy_range, past_range)

def check_range(ws, rng, user_rng, table_type, StartCol = 5):
    EndCol = get_split_col(ws, rng, StartCol)
    if table_type == 'AccSales':
        StartCol = EndCol + 1
        EndCol = get_split_col(ws, rng, StartCol)
    raw_values = ws.Range(f'{rng}:{rng}').Value[0][StartCol:EndCol]
    if user_rng not in raw_values:
        return True
    else:
        return False


def main_func(user_values, pg_bar, out):
    check_report = True
    files = []
    pivot_sheet = dict()
    Excel = win32com.client.Dispatch("Excel.Application")
    Excel.DisplayAlerts = False
    Excel.Visible = False
    wb = Excel.Workbooks.Open(user_values['SummaryFile'])
    sheet_dict = create_sheet_dict(wb)
    out.Update('Считывание файлов')
    pg_bar.Update(1)
    for k, v in sheet_dict.items():
        if user_values[k] == '':
            is_empty = True
        else:
            is_empty = False
        files.append(v.object(user_values[k], k, v.sheetname, is_empty, user_values['spt']))

    try:
        out.Update('Запись в файл')
        pg_bar.Update(2)

        for i, obj in enumerate(files):
            if files[i].type_file != 'CRM' and files[i].is_empty == False:
                sheet = wb.Worksheets(files[i].sheet_name)
                StartRow = get_last_row_from_column(sheet, option=False) + 2
                df = files[i].df
                add_df = files[i].additional_frame
                if files[i].type_file == 'AccPay':
                    df = create_cumsum_column(sheet, df)
            elif files[i].type_file == 'CRM' and files[i].is_empty == False:
                sheet = wb.Worksheets(files[i].sheet_name)
                if sheet.FilterMode:
                    sheet.ShowAllData()
                StartRow = get_last_row_from_column(sheet, option=False) + 2
                sheet.Range(f"B5:BE{StartRow}").ClearContents()
                df = files[i].full_df
                add_df = files[i].additional_frame
                StartRow = 5
            if files[i].is_empty == False:
                if files[i].type_file != 'CRM':
                    for col in ['L:L', 'N:N', 'W:W']:
                        sheet.Range(col).NumberFormat = '@'
                        sheet.Range(col).Replace(',', '.')
                    for col in ['M:M', 'Q:Q', 'AB:AB', 'O:O']:
                        sheet.Range(col).NumberFormat = '# ##0'
                else:
                    for col in ['T:T', 'U:U', 'AR:AR', 'AS:AS']:
                        sheet.Range(col).NumberFormat = '@'
                        sheet.Range(col).Replace(',', '.')


                sheet.Range(sheet.Cells(StartRow, files[i].column),  # Cell to start the "paste"
                            sheet.Cells(StartRow + len(df.index) - 1,
                                        files[i].column + len(df.columns) - 1)  # No -1 for the index
                            ).Value = df.values

                sheet.Range(sheet.Cells(StartRow, files[i].additional_column),  # Cell to start the "paste"
                            sheet.Cells(StartRow + len(add_df.index) - 1,
                                        files[i].additional_column + len(add_df.columns) - 1)  # No -1 for the index
                            ).Value = add_df.values
                UpdateEndRow = get_last_row_from_column(sheet, 'B', True)
                if files[i].type_file != 'CRM':
                    sheet.Range(f'B{6}:AB{6}').Copy()
                    sheet.Range(f'B{StartRow}:AB{UpdateEndRow}').PasteSpecial(-4122)


            if obj.type_file in ('AccPay', 'AccSales'):
                if not obj.is_empty:
                    pivot_sheet[obj.type_file] = obj.period
                else:
                    pivot_sheet[obj.type_file] = []

        for k, v in pivot_sheet.items():
            if v == []:
                pivot_sheet[k] = max(pivot_sheet.values(), key=lambda x: len(x))

        sheet = wb.Worksheets([ws.Name for ws in wb.Sheets][0])
        pivot_sheet['DownTable'] = max(pivot_sheet.values(), key=lambda x: len(x))

        EndRow = get_last_row_from_column(sheet, 'B', True)
        SplitRow = get_split_row(sheet, EndRow, 'СВЕРКА 1С и CRM') - 2

        DownTableDict = {element: get_split_row(sheet, EndRow, element, SplitRow) + get_split_row(sheet, EndRow, StartRow=get_split_row(sheet, EndRow, element, SplitRow), option=False)
                         for element in ['По данным 1С', 'По данным CRM', 'Разница']}

        # list_for_bottom = [get_split_row(sheet, EndRow, element, SplitRow) for element in
        #                        ('По данным 1С', 'По данным CRM', 'Разница')]

        UpTableDict = {element: get_split_row(sheet, SplitRow, element) + get_split_row(sheet, SplitRow, StartRow=get_split_row(sheet, SplitRow, element), option=False)
                       for element in ['Продажи тыс. руб. (накопительный итог) без учета ВГО и дополнительных корректировок',
                               'Продажи кв.м. (накопительный итог) без учета ВГО и дополнительных корректировок']}
        out.Update('Оформление СВОДа: добавление столбцов')
        pg_bar.Update(3)
        for k, v in pivot_sheet.items():
            for element in v:
                if k in ('AccPay', 'AccSales'):
                    check_rng = check_range(sheet, UpTableDict[list(UpTableDict.keys())[0]], element, k)
                else:
                    check_rng = check_range(sheet, DownTableDict[list(DownTableDict.keys())[0]], element, k)
                if check_rng:
                    obj = create_custom_range(sheet, user_values['prj'][0], k, EndRow, SplitRow) # user_values['project']
                    if k == 'AccPay':
                        sheet.Range(obj.past).Insert()

                    sheet.Range(obj.copy).Copy()
                    sheet.Range(obj.past).PasteSpecial(-4123)
                    sheet.Range(obj.past).PasteSpecial(-4122)
                    if k == 'DownTable':
                        sheet.Range(obj.past).EntireColumn.AutoFit()

                    if k in ('AccPay', "AccSales"):
                        for key, value in UpTableDict.items():
                            temp_range = re.sub(r'\d+', str(value), obj.past)
                            sheet.Range(temp_range).Value = element
                    else:
                        for key, value in DownTableDict.items():
                            temp_range = re.sub(r'\d+', str(value), obj.past)
                            sheet.Range(temp_range).Value = element

            if k == 'AccPay' and check_rng:
                for key, value in EditFormulasDict.items():
                    InitRow = get_split_row(sheet, EndRow, user_values['prj'][0], get_split_row(sheet, EndRow, key))
                    SumRow = get_split_row(sheet, EndRow, 'ИТОГО', InitRow)-1
                    letter = re.split(r'\d+', obj.past.split(':')[0])[0]
                    num_letter = get_excel_range(letter) + 2
                    end_letter = get_split_col(sheet, InitRow, num_letter + 1)
                    for col in range(num_letter, end_letter):
                        init_letter = get_excel_range(col)
                        num_col = get_excel_range(init_letter)
                        for i in range(InitRow, SumRow):
                            recent_column = re.findall(r'\+([A-Z]+:[A-Z]+)', sheet.Range(f'{init_letter}{i}').Formula)[0]
                            if get_excel_range(recent_column.split(':')[0]) != num_col-1:

                                sheet.Range(f'{init_letter}{i}').Formula = re.sub(r'\+([A-Z]+:[A-Z]+)',
                                                                                  f'+{get_excel_range(num_col-2)}:{get_excel_range(num_col-2)}',
                                                                                  sheet.Range(f'{init_letter}{i}').Formula)

                pass
        if user_values['-IN2-']:
            out.Update('Оформление СВОДа: добавление строк')
            pg_bar.Update(4)
            add_rows(sheet, user_values['new_data'])

    except Exception as exp:
        check_report = exp.args
    finally:
        wb.Save()
        wb.Close()
        Excel.Quit()
        return check_report

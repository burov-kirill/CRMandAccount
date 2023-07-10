import sys
import time

import win32com.client
import PySimpleGUI as sg
PROJECT_NAMES = ['ПУТИЛКОВО', 'АЛХИМОВО', 'СПУТНИК', 'ВЕРЕЙСКАЯ', "ГОРКИ-ПАРК", "ДОЛИНА ЯУЗЫ", "ЕГОРОВО-ПАРК",
                 "ЗАРЕЧЬЕ-ПАРК", "КВАРТАЛ ИВАКИНО", "ЛЮБЕРЦЫ", "МОЛЖАНИНОВО", "МЫТИЩИ-ПАРК", "НЕКРАСОВКА",
                 "НОВОДАНИЛОВСКАЯ", "НОВОЕ ВНУКОВО", "ОСТАФЬЕВО", "ПРИБРЕЖНЫЙ ПАРК", "ПРИГОРОД ЛЕСНОЕ", "ПЯТНИЦКИЕ ЛУГА",
                 "РУБЛЕВСКИЙ КВАРТАЛ", "ТОМИЛИНО", "ТРОПАРЕВО-ПАРК"]
def init_panel():
    layout = [
            [sg.Listbox(PROJECT_NAMES, key='prj', select_mode = 'LISTBOX_SELECT_MODE_SINGLE',
                        size = (30, 5), sbar_trough_color='#007bfb', sbar_frame_color='#007bfb',
                        sbar_arrow_color='#ffffff', sbar_relief='RELIEF_FLAT',
                        highlight_background_color='#007bfb', enable_events=True)],
            [sg.Text('Данные для редактирования', visible=False, key='spt_text', background_color='#007bfb', font='bold')],
            [sg.Input(key='spt', visible=False), sg.FileBrowse(key='spt_browse', visible=False,
                                                           button_color='#007bfb', button_text='Выбрать')],

            [sg.Text('Выбрать файл с карточкой счета 76', background_color='#007bfb', font='bold')],
            [sg.Input(key='AccPay'), sg.FileBrowse(button_color='#007bfb', button_text='Выбрать')],
            [sg.Text('Выбрать файл с карточкой счета 90', background_color='#007bfb', font='bold')],
            [sg.Input(key='AccSales'), sg.FileBrowse(button_color='#007bfb', button_text='Выбрать')],
            [sg.Text('Выбрать файл с данными CRM', background_color='#007bfb', font='bold')],
            [sg.Input(key='CRM'), sg.FileBrowse(button_color='#007bfb', button_text='Выбрать')],
            [sg.Text('Выбрать сводный файл', background_color='#007bfb', font='bold')],
            [sg.Input(key='SummaryFile'), sg.FileBrowse(button_color='#007bfb', button_text='Выбрать')],

            [sg.Radio('Есть дополнение', "RADIO1", default=False, key="-IN2-", enable_events=True,
                  background_color='#007bfb'),
            sg.Radio('Нет дополнений', "RADIO1", default=True, enable_events=True, background_color='#007bfb',
                  key='-IN1-')],
            [sg.Input(key='new_data', visible=False),
                sg.FileBrowse(key='new_data_browse', visible=False, button_color='#007bfb', button_text='Выбрать')],
            [sg.OK(button_color='#007bfb', button_text='Далее'), sg.Cancel(button_color='#007bfb', button_text='Выход')]

        ]
    yeet = sg.Window('Сверка БИТ и CRM', background_color='#007bfb', layout=layout)

    while True:
        event, values = yeet.read()
        if event in ('Выход', sg.WIN_CLOSED):
            sys.exit()
        elif '-IN2-' in event:
            yeet['new_data'].Update(visible=True)
            yeet['new_data_browse'].Update(visible=True)
            yeet.refresh()
        elif '-IN1-' in event:
            yeet['new_data'].Update(visible=False)
            yeet['new_data_browse'].Update(visible=False)
            yeet.refresh()
        elif 'prj' in event and values['prj'][0] == 'СПУТНИК':
            yeet['spt_text'].Update(visible=True)
            yeet['spt'].Update(visible=True)
            yeet['spt_browse'].Update(visible=True)
            yeet.refresh()
        elif 'prj' in event and values['prj'][0] != 'СПУТНИК':
            yeet['spt_text'].Update(visible=False)
            yeet['spt'].Update(visible=False)
            yeet['spt_browse'].Update(visible=False)
            yeet.refresh()
        elif event == 'Далее':
            break

    yeet.close()
    check_values = check_user_values(user_values=values)
    if check_values:
        return values
    else:
        check_input_error = input_error_panel()
        if check_input_error:
            return init_panel()


def check_user_values(user_values):
    keys = ['AccPay', 'SummaryFile', 'CRM']
    if any(map(lambda x: user_values[x] == '' or user_values[x] == [], keys)):
        return False
    else:
        return True


def input_error_panel():
    event = sg.popup('Ошибка ввода', 'При вводе данных возникла ошибка.\nВы хотите повторить ввод данных?',
                     background_color='#007bfb', button_color=('white', '#007bfb'),
                     title='Ошибка', custom_text=('Да', 'Нет'))
    if event == 'Да':
        return True
    else:
        sys.exit()

def end_panel(path):
    event = sg.popup('Сверка завершена\nОткрыть обработанный файл?', background_color='#007bfb',
                     button_color=('white', '#007bfb'),
                     title='Завершение работы', custom_text=('Да', 'Нет'))
    if event == 'Да':
        Excel = win32com.client.Dispatch("Excel.Application")
        Excel.Visible = True
        Excel.Workbooks.Open(Filename = path)
        time.sleep(5)
        del Excel
    else:
        sys.exit()

def error_panel(exp_desc):
    event = sg.popup_ok(f'При обработке данных возникла следующая ошибка:\n{exp_desc}',
                     background_color='#007bfb', button_color=('white', '#007bfb'),
                     title='Внутренняя ошибка')
    if event == 'OK':
        sys.exit()
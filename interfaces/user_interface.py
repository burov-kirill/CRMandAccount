import sys
import time

import win32com.client
import PySimpleGUI as sg

from update_scheme.config import VERSION
from update_scheme.update import check_version, call_updater

sg.LOOK_AND_FEEL_TABLE['SamoletTheme'] = {
                                        'BACKGROUND': '#007bfb',
                                        'TEXT': '#FFFFFF',
                                        'INPUT': '#FFFFFF',
                                        'TEXT_INPUT': '#000000',
                                        'SCROLL': '#FFFFFF',
                                        'BUTTON': ('#FFFFFF', '#007bfb'),
                                        'PROGRESS': ('#354d73', '#FFFFFF'),
                                        'BORDER': 1, 'SLIDER_DEPTH': 0,
                                        'PROGRESS_DEPTH': 0, }
def get_projects_list():
    project_list = []
    with open('__PROJECTS__.txt', 'r', encoding='utf8') as file:
        for line in file.readlines():
            project_list.append(line.replace('\n', ''))
    return project_list
#check if word
def set_new_project(project_name):
    with open('__PROJECTS__.txt', 'a', encoding='utf8') as file:
        file.write(f'\n{project_name}')

# PROJECT_NAMES = ['ПУТИЛКОВО', 'АЛХИМОВО', 'СПУТНИК', 'ВЕРЕЙСКАЯ', "ГОРКИ-ПАРК", "ДОЛИНА ЯУЗЫ", "ЕГОРОВО-ПАРК",
#                  "ЗАРЕЧЬЕ-ПАРК", "КВАРТАЛ ИВАКИНО", "ЛЮБЕРЦЫ", "МОЛЖАНИНОВО", "МЫТИЩИ-ПАРК", "НЕКРАСОВКА",
#                  "НОВОДАНИЛОВСКАЯ", "НОВОЕ ВНУКОВО", "ОСТАФЬЕВО", "ПРИБРЕЖНЫЙ ПАРК", "ПРИГОРОД", "ПЯТНИЦКИЕ ЛУГА",
#                  "РУБЛЕВСКИЙ КВАРТАЛ", "ТОМИЛИНО", "ТРОПАРЕВО-ПАРК"]


def init_panel():
    sg.theme('SamoletTheme')
    PROJECT_NAMES = get_projects_list()
    UPD_FRAME = [[sg.Button('Проверка', key='check_upd'), sg.Text('Нет обновлений', key='not_upd_txt'),
                  sg.Push(),
                  sg.pin(sg.Text('Доступно обновление', justification='center', visible=False, key='upd_txt', background_color='#007bfb', font='bold')),
                  sg.Push(),
                sg.pin(sg.Button('Обновить', key='upd_btn',  visible=False))],
    ]
    PRJ_FRAME = [[sg.Input(do_not_clear=True, size=(30, 1), enable_events=True, key='_INPUT_')],
                 [sg.Listbox(PROJECT_NAMES, key='prj', select_mode = 'LISTBOX_SELECT_MODE_SINGLE',
                        size = (30, 5), sbar_trough_color='#007bfb', sbar_frame_color='#007bfb',
                        sbar_arrow_color='#ffffff', sbar_relief='RELIEF_FLAT',
                        highlight_background_color='#007bfb', enable_events=True),
                  sg.Push(),
                  sg.Button('Добавить новый проект', key='new_prj'),
                  sg.Push()]]

    # DOC_FRAME = [
    #         [sg.Text('Счет 76', background_color='#007bfb', font='bold')],
    #         [sg.Input(key='AccPay'), sg.FileBrowse(button_color='#007bfb', button_text='Выбрать')],
    #         [sg.Text('Счет 90', background_color='#007bfb', font='bold')],
    #         [sg.Input(key='AccSales'), sg.FileBrowse(button_color='#007bfb', button_text='Выбрать')],
    #         [sg.Text('Данные CRM', background_color='#007bfb', font='bold')],
    #         [sg.Input(key='CRM'), sg.FileBrowse(button_color='#007bfb', button_text='Выбрать')],
    #         [sg.Text('Сводный файл', background_color='#007bfb', font='bold')],
    #         [sg.Input(key='SummaryFile'), sg.FileBrowse(button_color='#007bfb', button_text='Выбрать')],
    #
    #         [sg.Radio('Есть дополнение', "RADIO1", default=False, key="-IN2-", enable_events=True,
    #               background_color='#007bfb'),
    #         sg.Radio('Нет дополнений', "RADIO1", default=True, enable_events=True, background_color='#007bfb',
    #               key='-IN1-')],
    #         [sg.Input(key='new_data', visible=False),
    #             sg.FileBrowse(key='new_data_browse', visible=False, button_color='#007bfb', button_text='Выбрать')],
    #         [sg.pin(sg.Column(layout=[[sg.Text('Данные для редактирования', visible=False, key='spt_text', background_color='#007bfb',
    #                   font='bold')],
    #          [sg.Input(key='spt', visible=False),
    #           sg.FileBrowse(key='spt_browse', button_color='#007bfb', visible=False, button_text='Выбрать')]], key='--COL--',
    #                       visible=False), shrink=True)]
    # ]

    NEW_DOC_FRAME = [
        [sg.Column([
                    [sg.pin(sg.Checkbox('Добавить строки', background_color='#007bfb', enable_events=True, key='--ADD_STRING--'), shrink=True)],
                    [sg.pin(sg.Text('Номенклатура', visible=False, key='spt_text', background_color='#007bfb', font='bold'))],
                    [sg.pin(sg.Input(key='spt', visible=False)), sg.pin(sg.FileBrowse(key='spt_browse',button_color='#007bfb',
                                                                                           visible=False,button_text='Выбрать'))],
                    [sg.Text('Счет 76', background_color='#007bfb', font='bold')],
                    [sg.Input(key='AccPay'), sg.FileBrowse(button_color='#007bfb', button_text='Выбрать')],
                    [sg.Text('Счет 90', background_color='#007bfb', font='bold')],
                    [sg.Input(key='AccSales'), sg.FileBrowse(button_color='#007bfb', button_text='Выбрать')],
                    [sg.Text('Данные CRM', background_color='#007bfb', font='bold')],
                    [sg.Input(key='CRM'), sg.FileBrowse(button_color='#007bfb', button_text='Выбрать')],
                    [sg.Text('Сводный файл', background_color='#007bfb', font='bold')],
                    [sg.Input(key='SummaryFile'), sg.FileBrowse(button_color='#007bfb', button_text='Выбрать')],
                    [sg.pin(sg.Text('Новые строки', visible=False, key='new_data_text', background_color='#007bfb', font='bold'))],
                    [sg.pin(sg.Input(key='new_data', visible=False)), sg.pin(sg.FileBrowse(key='new_data_browse', visible=False,
                                                                                    button_color='#007bfb', button_text='Выбрать'))],
                    ],key='--DOC_COL--', scrollable = True, vertical_scroll_only = True, size=(430,300), background_color='#007bfb')]
    ]
    layout = [
            [sg.Frame(layout=UPD_FRAME, title='Обновление',background_color='#007bfb', key='--UPD_FRAME--',size=(470, 60))],
            [sg.Frame(layout=PRJ_FRAME, title='Выбор проекта', background_color='#007bfb', size=(470, 150), )],
            [sg.Frame(layout=NEW_DOC_FRAME, title='Выбор документов', background_color='#007bfb', size=(470, 350))],
            [sg.OK(button_color='#007bfb', button_text='Далее'), sg.Cancel(button_color='#007bfb', button_text='Выход')]

        ]
    yeet = sg.Window(f'Сверка БИТ и CRM {VERSION}', background_color='#007bfb', layout=layout)
    check, upd_check = False, True
    while True:
        event, values = yeet.read(timeout=10)
        if check:
            upd_check = check_version()
            check = False
        if event in ('Выход', sg.WIN_CLOSED):
            sys.exit()
        if event == 'check_upd':
            check = True
        if not upd_check:
            yeet['not_upd_txt'].Update(visible=False)
            yeet['upd_txt'].Update(visible=True)
            yeet['upd_btn'].Update(visible=True)
        if event == 'upd_btn':
            yeet.close()
            call_updater('pocket')
        if values['_INPUT_'] != '' and event == '_INPUT_':  # if a keystroke entered in search field
            search = values['_INPUT_']
            new_values = [x for x in PROJECT_NAMES if search in x]  # do the filtering
            yeet.Element('prj').Update(new_values)

        elif values['_INPUT_'] == '' and (event == 'new_prj' or len(yeet.Element('prj').get_list_values())<len(PROJECT_NAMES)):
            yeet.Element('prj').Update(PROJECT_NAMES)
        # elif event == 'prj':
        #     sg.PopupOK('kdkd')
        if 'new_prj' in event:
            set_new_project(values['_INPUT_'])
            yeet.Element('_INPUT_').Update('')
            PROJECT_NAMES = get_projects_list()
            yeet.Element('prj').Update(PROJECT_NAMES)
        elif 'prj' in event and values['prj'][0] == 'СПУТНИК':
            yeet['spt_text'].Update(visible=True)
            yeet['spt'].Update(visible=True)
            yeet['spt_browse'].Update(visible=True)
            yeet.refresh()
            yeet['--DOC_COL--'].contents_changed()
        elif 'prj' in event and values['prj'][0] != 'СПУТНИК':
            yeet['spt_text'].Update(visible=False)
            yeet['spt'].Update(visible=False)
            yeet['spt_browse'].Update(visible=False)
            yeet.refresh()
            yeet['--DOC_COL--'].contents_changed()
        elif event == '--ADD_STRING--' and values['--ADD_STRING--'] == True:
            yeet['new_data_text'].Update(visible=True)
            yeet['new_data'].Update(visible=True)
            yeet['new_data_browse'].Update(visible=True)
            yeet.refresh()
            yeet['--DOC_COL--'].contents_changed()
        elif event == '--ADD_STRING--' and values['--ADD_STRING--'] == False:
            yeet['new_data_text'].Update(visible=False)
            yeet['new_data'].Update(visible=False)
            yeet['new_data_browse'].Update(visible=False)
            yeet.refresh()
            yeet['--DOC_COL--'].contents_changed()

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
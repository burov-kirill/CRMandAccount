import sys
import time

import win32com.client
import PySimpleGUI as sg
from copy import copy
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

PERIODS = ['Месяц', 'Квартал', 'Полугодие', 'Год']
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
    new_value_list = []
    drop_prj = ''
    UPD_FRAME = [[sg.Button('Проверка', key='check_upd'), sg.Text('Нет обновлений', key='not_upd_txt'),
                  sg.Push(),
                  sg.pin(sg.Text('Доступно обновление', justification='center', visible=False, key='upd_txt', background_color='#007bfb', font='bold')),
                  sg.Push(),
                sg.pin(sg.Button('Обновить', key='upd_btn',  visible=False))],
    ]
    PRJ_FRAME = [[sg.Input(do_not_clear=True, size=(30, 1), enable_events=True, key='_INPUT_'), sg.Push(), sg.pin(sg.Button('Один проект', key = '-ONLY_ONE_PRJ-', visible=False)),
                  sg.pin(sg.Button('Несколько проектов', key = '-SEVERAL_PRJ-')),  sg.Push()],
                 [sg.Listbox(PROJECT_NAMES, key='prj', select_mode = 'LISTBOX_SELECT_MODE_SINGLE',
                        size = (30, 5), sbar_trough_color='#007bfb', sbar_frame_color='#007bfb',
                        sbar_arrow_color='#ffffff', sbar_relief='RELIEF_FLAT',
                        highlight_background_color='#007bfb', enable_events=True),
                  sg.Push(),
                  sg.pin(sg.Button('Добавить новый проект', key='new_prj', visible=True)),
                  sg.Listbox(new_value_list, size=(30, 5), enable_events=True, key='-upd_prj-',
                              visible=False),
                  sg.Push()],
                 [sg.Button('Удалить проект', key = '-DROP_PRJ-', visible=False), sg.Text('Начальный\nпериод'), sg.Combo(PERIODS, default_value='Полугодие', key='--FROM_PERIOD--'), sg.Push(),
                 sg.Text('Конечный\nпериод'), sg.Combo(PERIODS,default_value='Полугодие', key='--TO_PERIOD--')]
                 ]


    NEW_DOC_FRAME = [
        [sg.Column([
                    [sg.pin(sg.Checkbox('Добавить строки', background_color='#007bfb', enable_events=True, key='--ADD_STRING--'), shrink=True), sg.Push(),
                     sg.pin(sg.Checkbox('Создать новый файл', background_color='#007bfb', enable_events=True, key='--CREATE_FILE--'), shrink=True),
                     sg.pin(sg.Checkbox('Ревью файла', background_color='#007bfb', enable_events=True, key='--REVIEW--'), shrink=True)],

                    [sg.pin(sg.Column([[sg.Text('Номенклатура', key='spt_text', background_color='#007bfb',
                                 font='bold')],
                            [sg.Input(key='spt'), sg.FileBrowse(key='spt_browse', button_color='#007bfb', button_text='Выбрать')]], visible=False, key = 'spt_col'))],
                    [sg.pin(sg.Column([
                        [sg.Text('Папка для сохранения', background_color='#007bfb', font='bold')],
                        [sg.Input(key='save_folder'), sg.FolderBrowse(button_color='#007bfb', button_text='Выбрать')]],
                             key='save_folder_col', visible=False))],
                    # [sg.Text('Счет 76', background_color='#007bfb', font='bold')],
                    [sg.pin(sg.Column([
                            [sg.Text('Счет 76', background_color='#007bfb', font='bold', key='AccPayTxt')],
                            [sg.Input(key='AccPay'), sg.FileBrowse(button_color='#007bfb', button_text='Выбрать', key='AccPayFile')],
                            [sg.Text('Счет 90', background_color='#007bfb', font='bold', key='AccSalesTxt')],
                            [sg.Input(key='AccSales'), sg.FileBrowse(button_color='#007bfb', button_text='Выбрать', key='AccSalesFile')],
                            [sg.Text('Данные CRM', background_color='#007bfb', font='bold', key='CRMTxt')],
                            [sg.Input(key='CRM'), sg.FileBrowse(button_color='#007bfb', button_text='Выбрать', key='CRMFile')],
                            [sg.Text('Сводный файл', background_color='#007bfb', font='bold', key='SummaryFileTxt')],
                            [sg.Input(key='SummaryFile'),sg.FileBrowse(button_color='#007bfb', button_text='Выбрать', key='SummaryFl')],
                            [sg.Text('Новые строки', visible=False, key='new_data_text', background_color='#007bfb',
                                 font='bold')],
                            [sg.Input(key='new_data', visible=False), sg.FileBrowse(key='new_data_browse', visible=False,
                                                                                button_color='#007bfb',
                                                                                button_text='Выбрать')]
                                    ], key='-FILE_PANELS-', visible=True))],
                    [sg.pin(sg.Column([
                            # [sg.Text('Номенклатура', visible=False, key='spt_text_fld', background_color='#007bfb',
                            #      font='bold')],
                            # [sg.Input(key='spt_fld', visible=False), sg.FileBrowse(key='spt_browse_fld', button_color='#007bfb',
                            #                                                visible=False, button_text='Выбрать')],
                            [sg.Text('Счет 76', background_color='#007bfb', font='bold', key='AccPayFldTxt')],
                            [sg.Input(key='AccPayFld'), sg.FolderBrowse(button_color='#007bfb', button_text='Выбрать', key='AccPayFolder')],
                            [sg.Text('Счет 90', background_color='#007bfb', font='bold',  key='AccSalesFldTxt')],
                            [sg.Input(key='AccSalesFld'), sg.FolderBrowse(button_color='#007bfb', button_text='Выбрать', key='AccSalesFolder')],
                            [sg.Text('Данные CRM', background_color='#007bfb', font='bold', key='CRMFldTxt')],
                            [sg.Input(key='CRMFld'), sg.FolderBrowse(button_color='#007bfb', button_text='Выбрать', key='CRMFolder')],
                            [sg.Text('Сводный файл', background_color='#007bfb', font='bold', key = 'SummaryFldTxt')],
                            [sg.Input(key='SummaryFld'), sg.FolderBrowse(button_color='#007bfb', button_text='Выбрать', key='SummaryFolder')],
                            [sg.Text('Новые строки', key='new_data_text_fld', background_color='#007bfb', font='bold', visible=False)],
                            [sg.Input(key='new_data_fld', visible=False), sg.FolderBrowse(key='new_data_browse_fld', visible=False,
                                                                                button_color='#007bfb',
                                                                                button_text='Выбрать')],

                    ],
                        key='-FOLDER_PANELS-', visible=False))],
                    ],key='--DOC_COL--', scrollable = True, vertical_scroll_only = True, size=(430,300), background_color='#007bfb')]
    ]
    layout = [
            [sg.Frame(layout=UPD_FRAME, title='Обновление',background_color='#007bfb', key='--UPD_FRAME--',size=(470, 60))],
            [sg.Frame(layout=PRJ_FRAME, title='Выбор проекта', background_color='#007bfb', size=(470, 200), )],
            [sg.Frame(layout=NEW_DOC_FRAME, title='Выбор документов', background_color='#007bfb', size=(470, 350))],
            [sg.OK(button_color='#007bfb', button_text='Далее'), sg.Cancel(button_color='#007bfb', button_text='Выход')]

        ]
    yeet = sg.Window(f'Сверка БИТ и CRM {VERSION}', background_color='#007bfb', layout=layout)
    check, upd_check, many_files = False, True, False
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
            new_values = [x for x in PROJECT_NAMES if search.upper() in x]  # do the filtering
            yeet.Element('prj').Update(new_values)
        if '-SEVERAL_PRJ-' in event:
            many_files = True
            yeet.Element('-SEVERAL_PRJ-').Update(visible=False)
            yeet.Element('-ONLY_ONE_PRJ-').Update(visible=True)
            yeet.Element('new_prj').Update(visible=False)
            yeet.Element('-upd_prj-').Update(visible=True)
            yeet.Element('-DROP_PRJ-').Update(visible=True)
            yeet.Element('-FILE_PANELS-').Update(visible=False)
            yeet.Element('-FOLDER_PANELS-').Update(visible=True)
            yeet.refresh()
            yeet['--DOC_COL--'].contents_changed()
        if '-ONLY_ONE_PRJ-' in event:
            many_files = False
            yeet.Element('-ONLY_ONE_PRJ-').Update(visible=False)
            yeet.Element('-SEVERAL_PRJ-').Update(visible=True)
            yeet.Element('new_prj').Update(visible=True)
            yeet.Element('-upd_prj-').Update(visible=False)
            yeet.Element('-DROP_PRJ-').Update(visible=False)
            yeet.Element('-FILE_PANELS-').Update(visible=True)
            yeet.Element('-FOLDER_PANELS-').Update(visible=False)
            yeet.refresh()
            yeet['--DOC_COL--'].contents_changed()
        if values['_INPUT_'] == '' and (event == 'new_prj' or len(yeet.Element('prj').get_list_values())<len(PROJECT_NAMES)):
            yeet.Element('prj').Update(PROJECT_NAMES)
        if 'prj' in event:
            selection = values[event][0]
            if selection not in new_value_list:
                new_value_list.append(selection)
                yeet.Element('-upd_prj-').update(new_value_list)
        if 'upd_prj' in event:
            drop_prj = values['-upd_prj-'][0]
        if '-DROP_PRJ-' in event:
            if drop_prj:
                new_value_list.remove(drop_prj)
                yeet.Element('-upd_prj-').update(new_value_list)
        if values['--CREATE_FILE--']:
            if many_files:
                yeet.Element('SummaryFldTxt').Update(visible=False)
                yeet.Element('SummaryFld').Update(visible=False)
                yeet.Element('SummaryFolder').Update(visible=False)
            else:
                yeet.Element('SummaryFileTxt').Update(visible=False)
                yeet.Element('SummaryFile').Update(visible=False)
                yeet.Element('SummaryFl').Update(visible=False)
            yeet.Element('save_folder_col').Update(visible=True)
            yeet.refresh()
            yeet['--DOC_COL--'].contents_changed()
        if not values['--CREATE_FILE--']:
            yeet.Element('save_folder_col').Update(visible=False)
            if many_files:
                yeet.Element('SummaryFldTxt').Update(visible=True)
                yeet.Element('SummaryFld').Update(visible=True)
                yeet.Element('SummaryFolder').Update(visible=True)
            else:
                yeet.Element('SummaryFileTxt').Update(visible=True)
                yeet.Element('SummaryFile').Update(visible=True)
                yeet.Element('SummaryFl').Update(visible=True)
            yeet.refresh()
            yeet['--DOC_COL--'].contents_changed()
        if event in '--REVIEW--':
            if values['--REVIEW--']:
                if values['--TO_PERIOD--'] == values['--FROM_PERIOD--']:
                    temp_date_list = copy(PERIODS)
                    temp_date_list.remove(values['--TO_PERIOD--'])
                    yeet.Element('--TO_PERIOD--').Update(temp_date_list[1])
            else:
                if values['--TO_PERIOD--'] != values['--FROM_PERIOD--']:
                    yeet.Element('--TO_PERIOD--').Update(values['--FROM_PERIOD--'])
            if not many_files:
                yeet.Element('AccPayTxt').Update(visible=not values['--REVIEW--'])
                yeet.Element('AccPay').Update(visible=not values['--REVIEW--'])
                yeet.Element('AccPayFile').Update(visible=not values['--REVIEW--'])

                yeet.Element('AccSalesTxt').Update(visible=not values['--REVIEW--'])
                yeet.Element('AccSales').Update(visible=not values['--REVIEW--'])
                yeet.Element('AccSalesFile').Update(visible=not values['--REVIEW--'])

                yeet.Element('CRMTxt').Update(visible=not values['--REVIEW--'])
                yeet.Element('CRM').Update(visible=not values['--REVIEW--'])
                yeet.Element('CRMFile').Update(visible=not values['--REVIEW--'])
            else:

                yeet.Element('AccPayFldTxt').Update(visible=not values['--REVIEW--'])
                yeet.Element('AccPayFld').Update(visible=not values['--REVIEW--'])
                yeet.Element('AccPayFolder').Update(visible=not values['--REVIEW--'])

                yeet.Element('AccSalesFldTxt').Update(visible=not values['--REVIEW--'])
                yeet.Element('AccSalesFld').Update(visible=not values['--REVIEW--'])
                yeet.Element('AccSalesFolder').Update(visible=not values['--REVIEW--'])

                yeet.Element('CRMFldTxt').Update(visible=not values['--REVIEW--'])
                yeet.Element('CRMFld').Update(visible=not values['--REVIEW--'])
                yeet.Element('CRMFolder').Update(visible=not values['--REVIEW--'])
            yeet.refresh()
            yeet['--DOC_COL--'].contents_changed()

        if 'new_prj' in event:
            set_new_project(values['_INPUT_'])
            yeet.Element('_INPUT_').Update('')
            PROJECT_NAMES = get_projects_list()
            yeet.Element('prj').Update(PROJECT_NAMES)
        elif 'prj' in event and values['prj'][0] == 'СПУТНИК':
            yeet['spt_col'].Update(visible=True)
            yeet.refresh()
            yeet['--DOC_COL--'].contents_changed()
        elif 'prj' in event and values['prj'][0] != 'СПУТНИК':
            if not many_files:
                yeet['spt_col'].Update(visible=False)
            elif many_files and 'СПУТНИК' not in new_value_list:
                yeet['spt_col'].Update(visible=False)

            yeet.refresh()
            yeet['--DOC_COL--'].contents_changed()
        elif event == '--ADD_STRING--' and values['--ADD_STRING--'] == True:
            if not many_files:
                yeet['new_data_text'].Update(visible=True)
                yeet['new_data'].Update(visible=True)
                yeet['new_data_browse'].Update(visible=True)
            else:
                yeet['new_data_text_fld'].Update(visible=True)
                yeet['new_data_fld'].Update(visible=True)
                yeet['new_data_browse_fld'].Update(visible=True)
            yeet.refresh()
            yeet['--DOC_COL--'].contents_changed()
        elif event == '--ADD_STRING--' and values['--ADD_STRING--'] == False:
            if not many_files:
                yeet['new_data_text'].Update(visible=False)
                yeet['new_data'].Update(visible=False)
                yeet['new_data_browse'].Update(visible=False)
            else:
                yeet['new_data_text_fld'].Update(visible=False)
                yeet['new_data_fld'].Update(visible=False)
                yeet['new_data_browse_fld'].Update(visible=False)
            yeet.refresh()
            yeet['--DOC_COL--'].contents_changed()

        elif event == 'Далее':
            break

    yeet.close()
    # check_values = check_user_values(user_values=values)
    check_values = True
    values['prj_list'] = new_value_list
    values['many_files'] = many_files
    values['prj'] = "".join(values['prj'])
    if check_values:
        return values
    else:
        check_input_error = input_error_panel()
        if check_input_error:
            return init_panel()


def check_user_values(user_values):
    if (user_values['AccPay'] == '' and user_values['--FROM_PERIOD--'] == user_values['--TO_PERIOD--'])\
            or (user_values['CRM'] == '' and user_values['--FROM_PERIOD--'] == user_values['--TO_PERIOD--']) or\
            (user_values['SummaryFile'] == '' and user_values['prj'] == '') or \
            (user_values['--CREATE_FILE--'] and user_values['--FROM_PERIOD--'] != user_values['--TO_PERIOD--']):
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

def end_panel(path = '', opt = True):
    if opt:
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
    elif not opt:
        succed_prj = '\n'.join([k for k, v in path.items() if v == True])
        failed_prj = '\n'.join([k for k, v in path.items() if v != True])
        sg.popup_auto_close(f'Сверка завершена\nУспешно завершенные проекты:\n{succed_prj}\n'
                            f'Необработанные проекты:\n{failed_prj}',
                         title='Завершение работы', auto_close_duration = 20)


def error_panel(exp_desc):
    event = sg.popup_ok(f'При обработке данных возникла следующая ошибка:\n{exp_desc}',
                     background_color='#007bfb', button_color=('white', '#007bfb'),
                     title='Внутренняя ошибка')
    if event == 'OK':
        sys.exit()
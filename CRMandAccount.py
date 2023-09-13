import os
import sys
from time import time
from settings.logs import log
from functions.manage import main_func
import PySimpleGUI as sg
from interfaces.user_interface import init_panel, end_panel, error_panel
from functions.search_function import create_projects_list
from update_scheme.update import killProcess, get_subpath

if __name__ == '__main__':
    try:
        pid = int(sys.argv[2])
    except:
        pass
    else:
        killProcess(pid)
        os.chdir(get_subpath(sys.argv[0], 1))
        # shutil.rmtree(f'{get_subpath(sys.argv[0], 1)}\\config', ignore_errors=True)
    start = time()
    user_values = init_panel()
    if user_values['--ADD_STRING--']:
        steps = 4
    else:
        steps = 3
    progressbar = [[sg.ProgressBar(steps, size=(50, 10),  orientation='h', key='pg_bar')]]
    outputwin = [[sg.Output(key='out')]]
    layout = [
        [sg.Frame('Прогресс', layout=progressbar, background_color='#007bfb', size=(300, 50), key='prg_frame')],
        [sg.Frame('Процессы', layout=outputwin,  background_color='#007bfb', size=(300, 50))]
    ]
    window = sg.Window('Работа', layout=layout, finalize=True, element_justification='center', background_color='#007bfb')
    pg_bar = window['pg_bar']
    out = window['out']
    default_event = True
    prj_status = dict()
    check_report = True
    exp_desc = ''
    avg_time_list = []
    while True:
        event, values = window.read(timeout=5)
        if event in ('Выход', sg.WIN_CLOSED):
            sys.exit()
        elif default_event:
            if user_values['many_files']:
                prj_lst = create_projects_list(user_values)
                prj_status = {project: False for project in user_values['prj_list']}
                for i, prj in enumerate(prj_lst, 1):
                    window.Element('prg_frame').Update(f"{prj['prj']} {i} из {len(prj_lst)}")
                    try:
                        log.info(f'Начата обработка следующего проекта: {prj["prj"]}')
                        avg_start_time = time()
                        values = main_func(prj, pg_bar, out)
                    except Exception as exp:
                        log.info(f'При обработке следующего проекта {prj["prj"]} возникло исключение:\n')
                        log.exception(exp)
                        os.system('TASKKILL /F /IM excel.exe')
                    else:
                        log.info(f'Обработка проекта {prj["prj"]} успешно завершена')
                        avg_stop_time = time()
                        avg_time_list.append((avg_stop_time - avg_start_time))
                        prj_status[prj['prj']] = True
            else:
                try:

                    user_values = main_func(user_values, pg_bar, out)
                except Exception as exp:
                    exp_desc = exp
                    log.info(f'При обработке следующего файла {user_values["prj"]} возникло исключение:\n')
                    log.exception(exp)
                    check_report = False
                    os.system('TASKKILL /F /IM excel.exe')
                else:
                    log.info(f'Обработка проекта успешно завершена')
                    check_report = True
            break
    window.close()

    if check_report == True:
        stop = time()
        log.info(f'Общее время работы составило: {round((stop - start)/60,2)} минут')
        if not user_values['many_files']:
            end_panel(user_values['SummaryFile'])
        else:
            log.info(f'Среднее время обработки одного проекта составило: {round((sum(avg_time_list) / len(avg_time_list)) / 60, 2)} минут')
            end_panel(path = prj_status, opt=False)
    else:
        error_panel(exp_desc)
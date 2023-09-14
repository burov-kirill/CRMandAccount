import os
import queue
import sys
import threading
from threading import Thread
import time
from time import time as tm
from settings.logs import log
from functions.manage import main_func
import PySimpleGUI as sg
from interfaces.user_interface import init_panel, end_panel, error_panel
from functions.search_function import create_projects_list
from update_scheme.update import killProcess, get_subpath
from queue import Empty


class ThreadWithReturnValue(Thread):

    def __init__(self, group=None, target=None, name=None,
                 args=(), kwargs={}, Verbose=None):
        Thread.__init__(self, group, target, name, args, kwargs)
        self._return = None

    def run(self):
        if self._target is not None:
            self._return = self._target(*self._args,
                                        **self._kwargs)

    def join(self, *args):
        Thread.join(self, *args)
        return self._return

def tasks(project_list):
    global q
    prj_status = dict()
    for i, project in enumerate(project_list):
        q.put(project["prj"])
        log.info(f'Начата обработка следующего проекта: {project["prj"]}')
        avg_start_time = tm()
        try:
            main_func(project)
        except Exception as exp:
            log.info(f'При обработке следующего проекта {project["prj"]} возникло исключение:\n')
            log.exception(exp)
            os.system('TASKKILL /F /IM excel.exe')
            log.info(f'Файлы проекта {project["prj"]} принудительно закрыты\n')
        else:
            log.info(f'Обработка проекта {project["prj"]} успешно завершена')
            avg_stop_time = tm()
            avg_time_list.append((avg_stop_time - avg_start_time))
            prj_status[project['prj']] = True
        finally:
            q.put(i + 1)
            time.sleep(5)

    return prj_status

if __name__ == '__main__':
    try:
        pid = int(sys.argv[2])
    except:
        pass
    else:
        killProcess(pid)
        os.chdir(get_subpath(sys.argv[0], 1))
        # shutil.rmtree(f'{get_subpath(sys.argv[0], 1)}\\config', ignore_errors=True)
    start = tm()
    user_values = init_panel()
    # if user_values['--ADD_STRING--']:
    #     steps = 4
    # elif user_values['--REVIEW--']:
    #     steps = 2
    # else:
    #     steps = 3



    if user_values['many_files']:
        prj_lst = create_projects_list(user_values)
        prj_status = {project: False for project in user_values['prj_list']}
    else:
        prj_lst = [user_values]
    if prj_lst == []:
        log.info('Не найдены документы ни для одного файла')
        sys.exit()
        # pg_bar.Update(f"{prj['prj']} {i} из {len(prj_lst)}")
    progressbar = [[sg.ProgressBar(len(prj_lst), size=(50, 10),  orientation='h', key='pg_bar')]]
    outputwin = [[sg.Output(key='out')]]
    layout = [
        [sg.Frame('Прогресс', layout=progressbar, background_color='#007bfb', size=(300, 50), key='prg_frame')],
        [sg.Frame('Процессы', layout=outputwin,  background_color='#007bfb', size=(300, 50))]
    ]
    window = sg.Window('Работа', layout=layout, finalize=True, element_justification='center', background_color='#007bfb')
    pg_bar = window['pg_bar']
    out = window['out']
    default_event = True
    check_report = True
    exp_desc = ''
    avg_time_list = []
    q = queue.Queue()
    prj_status = dict()
    # tasks(project_list=prj_lst)
    worker_task = ThreadWithReturnValue(target=tasks, args=[prj_lst])
    worker_task.setDaemon(True)
    worker_task.start()
    while True:
        event, values = window.read(timeout=100)
        if event == 'Cancel' or event is None:
            os.system('TASKKILL /F /IM excel.exe')
            sys.exit()
        try:
            value = q.get_nowait()
        except Empty:
            continue
        else:
            if isinstance(value, int):
                pg_bar.UpdateBar(value)
                window.Element('prg_frame').Update(f"{value} из {len(prj_lst)}")
                if value == len(prj_lst):  #
                    break
            else:
                window.Element('out').Update(value)
    window.close()

    #         except Exception as exp:
    #             log.info(f'При обработке следующего проекта {prj["prj"]} возникло исключение:\n')
    #             log.exception(exp)
    #             os.system('TASKKILL /F /IM excel.exe')
    #         else:
    #             log.info(f'Обработка проекта {prj["prj"]} успешно завершена')
    #             avg_stop_time = time()
    #             avg_time_list.append((avg_stop_time - avg_start_time))
    #             prj_status[prj['prj']] = True
    # else:
    #     try:
    #         worker_task = threading.Thread(target=main_func, args=[user_values, q])
    #         worker_task.setDaemon(True)
    #         worker_task.start()
    #         while True:
    #             event, values = window.read(timeout=100)
    #             if event == 'Cancel' or event is None:
    #                 break
    #             try:
    #                 progress_value = q.get_nowait()
    #             except Empty:
    #                 continue
    #             else:  # Читать данные
    #                 pg_bar.UpdateBar(progress_value)
    #                 if progress_value == steps:  #
    #                     break
    #         window.close()
    #     except Exception as exp:
    #         exp_desc = exp
    #         log.info(f'При обработке следующего файла {user_values["prj"]} возникло исключение:\n')
    #         log.exception(exp)
    #         check_report = False
    #         os.system('TASKKILL /F /IM excel.exe')
    #     else:
    #         log.info(f'Обработка проекта успешно завершена')
    #         check_report = True

    # while True:
    #     event, values = window.read(timeout=5)
    #     if event in ('Выход', sg.WIN_CLOSED):
    #         sys.exit()
    #     elif default_event:
    #         if user_values['many_files']:
    #             prj_lst = create_projects_list(user_values)
    #             prj_status = {project: False for project in user_values['prj_list']}
    #             for i, prj in enumerate(prj_lst, 1):
    #                 window.Element('prg_frame').Update(f"{prj['prj']} {i} из {len(prj_lst)}")
    #                 try:
    #                     log.info(f'Начата обработка следующего проекта: {prj["prj"]}')
    #                     avg_start_time = time()
    #                     worker_task = threading.Thread(target=task_1)
    #                     worker_task.setDaemon(True)
    #                     worker_task.start()
    #                     values = main_func(prj, pg_bar, out, q)
    #                 except Exception as exp:
    #                     log.info(f'При обработке следующего проекта {prj["prj"]} возникло исключение:\n')
    #                     log.exception(exp)
    #                     os.system('TASKKILL /F /IM excel.exe')
    #                 else:
    #                     log.info(f'Обработка проекта {prj["prj"]} успешно завершена')
    #                     avg_stop_time = time()
    #                     avg_time_list.append((avg_stop_time - avg_start_time))
    #                     prj_status[prj['prj']] = True
    #         else:
    #             try:
    #                 worker_task = threading.Thread(target=main_func, args=[user_values, q])
    #                 worker_task.setDaemon(True)
    #                 worker_task.start()
    #                 while True:
    #                     event, values = window.read(timeout=100)
    #                     if event == 'Cancel' or event is None:
    #                         break
    #                     try:
    #                         progress_value = q.get_nowait()
    #                     except Empty:
    #                         continue
    #                     else:  # Читать данные
    #                         pg_bar.UpdateBar(progress_value)
    #                         if progress_value == steps:  #
    #                             break
    #             except Exception as exp:
    #                 exp_desc = exp
    #                 log.info(f'При обработке следующего файла {user_values["prj"]} возникло исключение:\n')
    #                 log.exception(exp)
    #                 check_report = False
    #                 os.system('TASKKILL /F /IM excel.exe')
    #             else:
    #                 log.info(f'Обработка проекта успешно завершена')
    #                 check_report = True
    #         break
    # window.close()
    prj_status = worker_task.join()
    if check_report == True:
        stop = tm()
        log.info(f'Общее время работы составило: {round((stop - start)/60,2)} минут')
        if not user_values['many_files']:
            end_panel(user_values['SummaryFile'])
        else:
            log.info(f'Среднее время обработки одного проекта составило: {round((sum(avg_time_list) / len(avg_time_list)) / 60, 2)} минут')
            end_panel(path = prj_status, opt=False)
    else:
        error_panel(exp_desc)
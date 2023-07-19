import sys

from manage import main_func
import PySimpleGUI as sg
from interfaces.user_interface import init_panel, end_panel, error_panel
from update_scheme.update import killProcess

if __name__ == '__main__':
    try:
        pid = int(sys.argv[2])
    except:
        pass
    else:
        killProcess(pid)
    user_values = init_panel()
    if user_values['--ADD_STRING--']:
        steps = 4
    else:
        steps = 3
    progressbar = [[sg.ProgressBar(steps, orientation='h', size=(21, 10), key='pg_bar')]]
    outputwin = [[sg.Output(size=(37, 2), key='out')]]
    layout = [
        [sg.Frame('Прогресс', layout=progressbar, background_color='#007bfb')],
        [sg.Frame('Процессы', layout=outputwin,  background_color='#007bfb')]
    ]
    window = sg.Window('Работа', layout=layout, finalize=True, element_justification='center', background_color='#007bfb')
    pg_bar = window['pg_bar']
    out = window['out']
    default_event = True
    while True:
        event, values = window.read(timeout=5)
        if event in ('Выход', sg.WIN_CLOSED):
            sys.exit()
        elif default_event:
            check_report = main_func(user_values, pg_bar, out)
            break
    window.close()

    if check_report == True:
        end_panel(user_values['SummaryFile'])
    else:
        error_panel(check_report)
import sys
import threading
import time
from time import sleep
import PySimpleGUI as sg
import subprocess
import requests
import ssl
import urllib


from pathlib import Path

FOLDER_NAME = ''
APP_NAME = f'{FOLDER_NAME}CRMandBIT.exe'
APP_URL = 'https://raw.githubusercontent.com/burov-kirill/CRMandAccount/master/dist/CRMandBIT.exe'
VERSIONS_NAME = f'{FOLDER_NAME}versions.txt'
VERSION_URL = 'https://raw.githubusercontent.com/burov-kirill/CRMandAccount/master/__VERS__.txt'
PERCENT = None
LOGIN = 'burov-kirill'
ACCESSTOKEN = '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'

def get_current_version():
    default_str = ''
    versions = list()
    versions_file = Path(VERSIONS_NAME)
    if versions_file.is_file():
        with open(VERSIONS_NAME, 'r') as file:
            for line in file:
                versions.append(line.replace('\n', ''))
    else:
        # Path(FOLDER_NAME).mkdir(parents=True, exist_ok=True)
        open(VERSIONS_NAME, 'a+')
    if versions != []:
        return versions[-1]
    else:
        return default_str

def get_latest_version():
    error_report = False
    desc = ''
    for _ in range(10):
        try:
            # auth = (LOGIN, ACCESSTOKEN)
            # res = urllib.request.urlopen(VERSION_URL, context=context)
            res = requests.get(VERSION_URL)
            time.sleep(2)
            # if res.getcode() == 200:
            #     web_data = res.read().decode('utf-8')
            #     return web_data.split('\n')[-1]
            if res.status_code == 200:
                return res.text.split('\n')[-1]
        except Exception as exp:
            error_report = True
            desc = exp
    if error_report:
        error_panel(desc)

def error_panel(desc):
    event = sg.popup_ok(f'При загрузке данных возникла ошибка: {desc}',
                     background_color='#007bfb', button_color=('white', '#007bfb'),
                     title='Ошибка загрузки')
    if event == 'OK':
        sys.exit()
def set_version(version):
    with open(VERSIONS_NAME, 'a+') as file:
        file.write(version)

def download_file(window):
    # auth = (LOGIN, ACCESSTOKEN)
    # with urllib.request.urlopen(APP_URL, context=context) as r:
    with requests.get(APP_URL, stream=True) as r:
        chunk_size = 64*1024
        total_length = int(r.headers.get('content-length'))
        total = total_length//chunk_size if total_length % chunk_size == 0 else total_length//chunk_size + 1
        with open(APP_NAME, 'wb') as f:
            # i = 0
            # while True:
            #     buffer = r.read(chunk_size)
            #     if not buffer:
            #         break
            #     data_wrote = f.write(buffer)
            #     PERCENT = int((i+1)/total*100)
            #     window.write_event_value('Next', PERCENT)
            #     i+=1
            # # an integer value of size of written data

            for i, chunk in enumerate(r.iter_content(chunk_size=chunk_size)):
                f.write(chunk)
                PERCENT = int((i+1)/total*100)
                window.write_event_value('Next', PERCENT)
def create_download_window(title='Загрузка исполняемого файла'):
    progress_bar = [
        [sg.ProgressBar(100, size=(40, 20), pad=(0, 0), key='Progress Bar', border_width = 0),
         sg.Text("  0%", size=(4, 1), key='Percent', background_color='#007bfb', border_width=0), ],
    ]

    layout = [
        [sg.pin(sg.Column(progress_bar, key='Progress', visible=True, background_color='#007bfb',
                          pad=(0, 0), element_justification='center'))],
    ]
    window = sg.Window(title, layout, size=(600, 40), finalize=True,
                       use_default_focus=False, background_color='#007bfb')
    progress_bar = window['Progress Bar']
    percent = window['Percent']
    progressB = window['Progress']
    default_event = True
    while True:
        event, values = window.read(timeout=10)
        if event == sg.WINDOW_CLOSED:
            break
        elif default_event:
            default_event = False
            count = 0
            progress_bar.update(current_count=0, max=100)
            thread = threading.Thread(target=download_file, args=(window,), daemon=True)
            thread.start()
        elif event == 'Next':
            count = values[event]
            progress_bar.update(current_count=count)
            percent.update(value=f'{count:>3d}%')
            window.refresh()
            if count == 100:
                sleep(1)
                break
    window.close()

def check_update():
    current_version = get_current_version()
    latest_version = get_latest_version()
    if current_version != latest_version:
        create_download_window()
    else:
        return True

# if __name__ == '__main__':
#     current_version = get_current_version()
#     latest_version = get_latest_version()
#     if current_version == latest_version:
#         app_file = Path(APP_NAME)
#         if not app_file.is_file():
#             create_download_window()
#         subprocess.run([APP_NAME])
#     else:
#         if current_version == '':
#             title = 'Первоначальная загрузка файла'
#         else:
#             title = 'Идет обновление до последней версии'
#         create_download_window(title)
#         set_version(latest_version)
#         subprocess.run([APP_NAME])
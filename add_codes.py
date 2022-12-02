import PySimpleGUI as sg
import pandas as pd
from openpyxl import Workbook, load_workbook
from exceptions import *
from functions import *
import traceback

layout = [
    [sg.B('настроить колонки', key='-set_cols-')],
    [sg.T('файл'), sg.Input(key='-excel-'), sg.FileBrowse(file_types=(("excel", "*.xlsx"),("excel", "*.xls")))],
    [sg.T('ряд начала'), sg.Input(size=5, key='-start-', default_text='2')],
    [sg.B('начать', key='-process-')]
    ]


try:
    CATALOG = load_catalog('cat.pkl')
except FileNotFoundError:
    print('каталог не найден')
    exit()
except pickle.PickleError:
    print('ошибка загрузки каталога')
    exit()

COLUMN_SETTINGS = dict()
window = sg.Window('add codes', layout)
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
    elif event == '-set_cols-':
        if (settings := set_columns(CATALOG.columns[:15], COLUMN_SETTINGS)) != None:
            COLUMN_SETTINGS = settings
    elif event == '-process-':
        if (file := values['-excel-']) == '':
            sg.popup('файл не выбран')
            continue
        start = values['-start-']
        if start.isdigit():
            start = int(start)
        else:
            sg.popup('начальный ряд задается числом >= 1')
            continue
        try:
            wb = load_workbook(file)
        except Exception:
            sg.popup('не удалось загрузить файл')
            continue
        ws = wb.active
        try:
            children = generate_children(CATALOG)
            add_children(ws, start, children, COLUMN_SETTINGS, CATALOG)
            *dirr, suff = file.split('.')
            wb.save('.'.join(dirr)+'_res.'+suff)
            sg.popup('Done!')
        except Exception as e:
            sg.popup(e)
            print(traceback.print_exc())
        finally:
            wb.close()


window.close()
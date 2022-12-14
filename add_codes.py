import PySimpleGUI as sg
import pandas as pd
from openpyxl import Workbook, load_workbook
from exceptions import *
import sys
import pickle
import traceback
from functions import *


group_tab = [
    [sg.Checkbox('копировать значения для добавленных', key='-copy-')],
    [sg.B('начать', key='-process_group-')]
]

add_tab = [
    [sg.B('настроить колонки', key='-set_cols-')],
    [sg.T('ряд начала'), sg.Input(size=5, key='-start-', default_text='2')],
    [sg.B('начать', key='-process-')]
]

layout = [
    [sg.T('файл'), sg.Input(key='-excel-'), sg.FileBrowse(file_types=(("excel", "*.xlsx"),("excel", "*.xls")))],
    [sg.TabGroup([[
        sg.Tab(layout=add_tab, key='-add_tab-', title='добавить'),
        sg.Tab(layout=group_tab, key='-group_tab-', title='группировать и добавить')
    ]], size=(450, 100))]
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
    if 'process' in event:
        if (file := values['-excel-']) == '':
            sg.popup('файл не выбран')
            continue
        *dirr, suff = file.split('.')
        save_name = '.'.join(dirr)+'_res.'+suff
    if event == '-process-':
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
            wb.save(save_name)
            sg.popup('Done!')
        except Exception as e:
            sg.popup(e)
            print(traceback.print_exc())
        finally:
            wb.close()
    if event == '-process_group-':
        try:
            df = pd.read_excel(file)
            if not check_df(df):
                sg.popup('выберите подходящий файл (первая колонка должны быть "Код тов.")')
                continue
            children = generate_children(CATALOG)
            res = group(CATALOG, children, values['-copy-'], df)
            initial = dict.fromkeys(res.columns, False)
            for n in df.columns[1:]:
                initial[n] = True
            columns = check_box_window(initial)
            res = res[columns]
            print(type(res))
            res.to_excel(save_name)
            sg.popup('Done!')
        except Exception as e:
            sg.popup(e)
            print(traceback.print_exc())


window.close()
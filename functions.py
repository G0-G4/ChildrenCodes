import pickle
import PySimpleGUI as sg
from openpyxl.utils.cell  import column_index_from_string, get_column_letter
from collections import defaultdict
from exceptions import *

def load_catalog(name):
    with open(name, 'rb') as f:
        cat = pickle.load(f)
    return cat

def set_columns(names, settings):
    res = {}
    layout = [
        [sg.T(name), sg.Input(get_column_letter(settings[name]+1) if name in settings else '', key=name)] for name in names
    ] + [[sg.B('ok')]]
    window =  sg.Window('setting up columns', layout, modal=True)
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            break
        elif event == 'ok':
            for key in values:
                if values[key] == '':
                    continue
                try:
                    res[key] = column_index_from_string(values[key]) - 1
                except ValueError as e:
                    sg.popup(e)
                    break
            else:
                window.close()
                return res
    window.close()
    
def generate_children(cat):
    children = defaultdict(list)
    for code, article in zip(cat['Код тов.'][1:], cat['Артикул'][1:]):
        mcode = code
        if '/' in str(article):
            mcode = article.split('/')[0].strip()
            if mcode.isdigit():
                mcode = int(mcode)
                if mcode != code:
                    children[mcode].append(code)
    return children

def get_values(cat, codes, names):
    df = cat[cat['Код тов.'].isin(codes)][names]
    return [list(df[name]) for name in names]

def get_range(c1: int, r1: int, c2: int, r2: int):
    c1, c2 = get_column_letter(c1), get_column_letter(c2)
    return f'{c1}{r1}:{c2}{r2}'

def insert_values(ws, start, end, columns, values):
    for i, row in enumerate(ws.iter_rows(min_row = start, max_row = end)):
        for j, col in enumerate(columns):
            val = values[j][i]
            if type(val) == str:
                val = val.strip()
            print('      inserting', val, 'into', (start+i, col+1))
            try:
                row[col].value = val
            except IndexError:
                raise InsertError(f'cant insert into {(start+i, col+1)}',)

def add_children(ws, start, children, settings, cat):
    if 'Код тов.' not in settings:
        raise CodeColumnNotFound('please provide column for mother code')
    code_id = settings['Код тов.']
    if code_id >= len(ws[1]):
        raise OutOfBounds('provided mother code column is not in the table')
    i = start
    while True:
        row = ws[i]
        if row[code_id].value == None:
            break
        try:
            code = int(row[code_id].value)
        except Exception as e:
            raise MotherCodeError(f'Wrong mother code {row[code_id].value}')
        if code in children:
            num = len(children[code])
            print(code, 'moving down', len(children[code]), 'rows')
            # ws.insert_rows(i+1, num) # insert_rows inserts before index, indexing starts from one there for add 1
            rang = get_range(1, i+1, ws.max_column, ws.max_row)
            ws.move_range(rang, rows = num, translate = True)
            values = get_values(cat, children[code], settings.keys())
            insert_values(ws, i+1, i+num, settings.values(), values)
            i += num
        i += 1
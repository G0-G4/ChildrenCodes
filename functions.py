import pickle
import PySimpleGUI as sg
import pandas as pd
import numpy as np
from openpyxl.utils.cell  import column_index_from_string, get_column_letter
from collections import defaultdict
from exceptions import *

def load_catalog(name):
    with open(name, 'rb') as f:
        cat = pickle.load(f)
    return cat

def set_columns(names, settings, check=True):
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
                    if check:
                        res[key] = column_index_from_string(values[key]) - 1
                    else:
                        res[key] = values[key] # for second tab no need for check
                except ValueError as e:
                    sg.popup(e)
                    break
            else:
                window.close()
                return res
    window.close()
    
def generate_children(cat):
    clones = defaultdict(list)
    for code, clone in zip(cat['Код тов.'][1:], cat['Клон'][1:]):
        if type(clone) == int and clone != code:
            clones[clone].append(code)
    return clones

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
            if i != ws.max_row:
                rang = get_range(1, i+1, ws.max_column, ws.max_row)
                print(rang)
                ws.move_range(rang, rows = num, translate = True)
            values = get_values(cat, children[code], settings.keys())
            insert_values(ws, i+1, i+num, settings.values(), values)
            i += num
        i += 1


# def mother_children_table(cat, codes, settings):
#     cat_grouped = cat.set_index(['Клон', 'Код тов.'])
#     cat_grouped.sort_index(inplace = True)
#     print(settings.keys())
#     return cat_grouped.loc[codes]

# def read_codes(file, settings):
#     return pd.read_excel(file).iloc[:, settings['Код тов.']]

def check_df(df):
    if df.columns[0]!= 'Код тов.':
        return False
    return True

def group(cat, clones, copy, df):
    add_children = []
    # searching childten to add
    for _, row in df.iterrows():
        if row['Код тов.'] in clones:
            for ch in clones[row['Код тов.']]:
                if ch not in df['Код тов.'].values:
                    add_children.append(ch)
    # merging
    df = pd.concat([df, (pd.DataFrame({'Код тов.': add_children}))])
    res = df.merge(cat, on=['Код тов.'], how='left', suffixes=('','catalog'))
    res = res.set_index(['Клон', 'Код тов.'])
    res.sort_index(inplace = True)
    # copying values
    if copy:
        for i in res.index.get_level_values(0):
            n = len(res.loc[i])
            for name in df.columns[1:]:
                res.loc[i, name] = np.full(n, res.loc[i].iloc[0][name])
    return res

def check_box_window(initial):
    layout = [
        [sg.Column([[sg.Checkbox(name, key=name, default=val)] for name, val in initial.items()], scrollable=True, size=(200, 300)), sg.B('ok')]
]
    window =  sg.Window('choose', layout, modal=True)
    while True:
        event, values = window.read()
        if event:
            window.close()
        if event == sg.WIN_CLOSED:
            return []
        if event == 'ok':
            return list(map(lambda x: x[0], filter(lambda x: x[1], values.items())))

if __name__ == '__main__':
    print(check_box_window({'a':1, 'b':0}))
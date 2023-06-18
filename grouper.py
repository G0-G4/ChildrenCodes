from exceptions import *
from openpyxl.utils.cell  import column_index_from_string, get_column_letter
from collections import defaultdict

CODE_COLUMN = 'код'
CLONE_COLUMN = 'клон'

class Grouper():
    def __init__(self, wb, start, children, columns, cat):
        print(columns)
        self.wb = wb
        self.ws = wb.active
        self.i = start - 1 # т.к потом будет move который сдвинет
        self.children = children
        self.columns = columns
        self.cat = cat
        self.code_id = columns[CODE_COLUMN]
        self.BIAS = 0
        self.all_codes = self.get_all_codes()
        if self.code_id >= len(self.ws[1]):
            raise OutOfBounds('provided mother code column is not in the table')
        
    def save(self, name):
        self.wb.save(name)

    def get_mother(self, code):
        if len(self.cat[self.cat[CODE_COLUMN] == code][CLONE_COLUMN]) != 1:
            return 0
        return int(self.cat[self.cat[CODE_COLUMN] == code][CLONE_COLUMN])
    
    def is_mother(self, code):
        return code == self.get_mother(code)
    
    def move(self):
        # add skipping recently added or moved codes, skip group
        # think about it
        self.i += 1
        row = self.ws[self.i]
        print('current i = ', self.i)
        if row[self.code_id].value == None:
            return False
        try:
            code = int(row[self.code_id].value)
            print('code = ', code)
        except Exception as e:
            print(f'Wrong mother code {row[self.code_id].value}')
        else:
            return code
        return -1
    
    def get_range(self, c1: int, r1: int, c2: int, r2: int):
        c1, c2 = get_column_letter(c1), get_column_letter(c2)
        return f'{c1}{r1}:{c2}{r2}'
    
    def move_range(self, rows, frm, to = None):
        if not to:
            to = self.ws.max_row
            self.BIAS += rows # add to BIAS if moving all cells below
        rang = self.get_range(1, frm, self.ws.max_column, to)
        self.ws.move_range(rang, rows)
        print('moving ', rang, 'by ', rows, 'BIAS = ', self.BIAS)

    def get_biased_index(self, code):
        return self.all_codes[code] + self.BIAS
    
    def move_row(self, row, rows):
        self.move_range(rows=rows, frm=row, to=row)
        self.move_range(rows=-1, frm=row + 1)
    
    def move_mother(self, mother):
        self.move_range(rows=1, frm=self.i)
        mother_id = self.get_biased_index(mother)
        print('moving mother ', mother, 'biased id ', mother_id)
        # self.move_range(rows=self.i-mother_id, frm=mother_id, to=mother_id)
        # self.move_range(rows=-1, frm=mother_id + 1)
        self.move_row(mother_id, self.i - mother_id)
        print('skipping next')
        self.move()

    def get_values(self, codes, names):
        df = self.cat[self.cat[CODE_COLUMN].isin(codes)][names]
        return [list(df[name]) for name in names]

    def add_mother(self, mother):
        self.move_range(rows=1, frm=self.i)
        values = self.get_values([mother], self.columns.keys())
        print('inserting mother into ', self.i)
        self.insert_values(self.i, self.i, self.columns.values(), values)
        print('skipping next')
        self.move()

    def move_children(self, mother):
        brothers = self.children[mother]
        self.move_range(rows = len(brothers), frm=self.i) # free space for brothers
        for brother in brothers:
            ii = self.get_biased_index(brother)
            self.move_row(row = ii, rows = self.i - ii)
            self.move()


    def add_children(self, mother):
        ...


    def insert_values(self, start, end, columns, values):
        for i, row in enumerate(self.ws.iter_rows(min_row = start, max_row = end)):
            for j, col in enumerate(columns):
                val = values[j][i]
                if type(val) == str:
                    val = val.strip()
                print('      inserting', val, 'into', (start+i, col+1))
                try:
                    row[col].value = val
                except IndexError:
                    raise InsertError(f'cant insert into {(start+i, col+1)}',)

    def process(self):
        while code := self.move():
            mother = code
            if not self.is_mother(code):
                mother = self.get_mother(code)
            if mother == 0:
                continue
            if mother in self.all_codes:
                self.move_mother(mother)
            else:
                self.add_mother(mother)
            self.move_children(mother)
            self.add_children(mother)

    def get_all_codes(self):
        print('getting all codes')
        i = self.i
        codes = defaultdict(list)
        while code := self.move():
            codes[code] = self.i
        self.i = i
        print('=================')
        return codes

    

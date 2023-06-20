'''
load and pickle catalog
add to 1c program
'''

import pandas as pd
import pickle
import sys
if len(sys.argv) < 2:
    print('please provide a file')
    exit()
print('reading file (it will take some time)...')
cat = pd.read_excel(sys.argv[1])
cat.columns = list(map(str.strip, cat.columns))
cat['клон'] = pd.to_numeric(cat['Артикул'].str.split('/').str[0], errors='coerce')
cat['клон'] = cat['клон'].fillna(0)
cat['клон'] = cat['клон'].astype(int)
cat.loc[cat['Код'] == cat['Артикул'], 'клон'] = cat.loc[cat['Код'] == cat['Артикул'], 'Код']
print('saving file...')
with open('cat.pkl', 'wb') as f:
    pickle.dump(cat, f)
print('Done!')
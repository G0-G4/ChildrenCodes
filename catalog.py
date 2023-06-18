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
cat['клон'] = cat['артикул'].str.split('/').str[0]
cat['клон'] = cat['клон'].fillna(0)
cat.loc[cat['код'] == cat['артикул'], 'клон'] = cat.loc[cat['код'] == cat['артикул'], 'код']
print('saving file...')
with open('cat.pkl', 'wb') as f:
    pickle.dump(cat, f)
print('Done!')
'''
load and pickle catalog
'''

import pandas as pd
import pickle
import sys
if len(sys.argv) < 2:
    print('please provide a file')
    exit()
print('reading file (it will take some time)...')
cat = pd.read_excel(sys.argv[1])
names = dict(zip(cat.columns, cat.loc[[0]].values.flatten()))
cat.rename(columns = names, inplace = True)
print('saving file...')
with open('cat.pkl', 'wb') as f:
    pickle.dump(cat[1:], f)
print('Done!')
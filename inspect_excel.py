import os
import pandas as pd
files = ['E-Stop 1.0.xlsx', 'E-Stop 1.0 Rule.xlsx']
for fn in files:
    exists = os.path.exists(fn)
    print('---', fn, exists)
    if exists:
        df = pd.read_excel(fn)
        print('columns:', list(df.columns))
        print('first rows:')
        for row in df.head(5).to_dict(orient='records'):
            print(row)

import pandas as pd
import os
files = [
    'input_output/output/6.xlsx',
    'input_output/PLCTags_121001.xlsx',
    'input_output/PLCTags_259001.xlsx',
    'examples/e_stop_rules.xlsx',
    'examples/rules_sample.xlsx'
]
for fn in files:
    print('\nFILE', fn)
    if os.path.exists(fn):
        try:
            df = pd.read_excel(fn)
            print('shape', df.shape)
            print(df.head(20).to_string(index=False))
        except Exception as e:
            print('ERROR reading', fn, type(e).__name__, e)
    else:
        print('MISSING')

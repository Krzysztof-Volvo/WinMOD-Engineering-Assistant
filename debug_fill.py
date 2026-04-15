from main import fill_missing_formula_values, get_column_name
import pandas as pd

file_path = 'examples/e_stop_rules.xlsx'
df = pd.read_excel(file_path)
print('before', df.to_dict(orient='records'))
print('rule col', get_column_name(df, 'rule'))
df2 = fill_missing_formula_values(df, file_path, 0)
print('after', df2.to_dict(orient='records'))
print('rule values', df2['rule'].tolist())

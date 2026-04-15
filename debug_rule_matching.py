import pandas as pd
from main import apply_macro_rules, read_excel_best_header, has_macro_rules

signals = pd.read_excel('examples/signals_sample.xlsx')
macros = read_excel_best_header('examples/macros_sample.xlsx', expected_columns=['Symbol', 'macro signal', 'macro name', 'macro_type', 'macro type'])
rules = read_excel_best_header('examples/e_stop_rules.xlsx', expected_columns=['field', 'operator', 'value', 'macro signal', 'rule'], require_non_empty_columns=['rule', 'macro signal'])
print('signals cols', list(signals.columns))
print('macros cols', list(macros.columns))
print('rules cols', list(rules.columns))
print('rules head', rules.to_dict(orient='records'))
print('has_macro_rules', has_macro_rules(rules))
matched = apply_macro_rules(signals, macros, rules, output_consumer=print)
print('matched rows', len(matched))
print(matched.head().to_dict(orient='records'))

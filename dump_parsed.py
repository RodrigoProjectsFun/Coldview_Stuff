import pandas as pd
from SEACHACCOUNT import parse_fast

df = parse_fast('test_repro.txt')
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)

with open('dump_output.txt', 'w', encoding='utf-8') as f:
    f.write("Columns:\n")
    for c in df.columns:
        f.write(f"{c}: {repr(df[c].iloc[0])}\n")

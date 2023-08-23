import pandas as pd

df = pd.read_csv('DOP_KEP.txt', sep='\t', encoding='windows-1252')
df.to_excel('output.xlsx', index=False)
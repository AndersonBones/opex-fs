import pandas as pd

path = r"C:\Users\Anderson\Desktop\OPEX\opex-fs\src\data\M_OPEX.xlsx"
df = pd.read_excel(path, decimal=',', sheet_name="M_OPEX", index_col=False)


df = df.iloc[: , 1:] # remove a coluna Cost Center

# remove as duas primeiras linhas (Data e Year)
df.drop(index=0, inplace=True)
df.drop(index=1, inplace=True)



df.to_excel('result.xlsx', header=True, sheet_name='Opex', index=False)


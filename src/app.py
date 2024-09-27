import pandas as pd

path = r"C:\Users\anderson.bones\Desktop\Projetos\Python\opex-fs\src\data\M_OPEX.xlsx"
df = pd.read_excel(path, decimal=',', sheet_name="M_OPEX", index_col=False, )
import os

cc = [
    
]



# remove as duas primeiras linhas (Data e Year)
df.drop(index=0, inplace=True)
df.drop(index=1, inplace=True)

df.rename(columns={"Unnamed: 0": "Totais_label"}, inplace=True)
df.rename(columns={"Version":"Conta"}, inplace=True)



df.drop(df[(df.Conta == "Account")].index, inplace=True) # Remove account

df.drop(df[(df.Totais_label != "Totais")].index, inplace=True)

df = df.iloc[: , 1:] # remove a coluna Cost Center


df = df.drop(columns=['Fct - Previous'])


output="result.xlsx"
df.to_excel('result.xlsx', header=True, sheet_name='Opex', index=False)


os.popen(output)

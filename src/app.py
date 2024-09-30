import pandas as pd
from IPython.display import display, HTML
import numpy as np


path = r"C:\Users\Anderson\Desktop\OPEX\opex-fs\src\data\M_OPEX.xlsx"
df = pd.read_excel(path, decimal=',', sheet_name="M_OPEX", index_col=False, )
import os


def remove_rows():
    # remove as duas primeiras linhas (Data e Year)
    df.drop(index=0, inplace=True)
    #df.drop(index=1, inplace=True)


remove_rows()

def rename_column():
    df.rename(columns={"Unnamed: 0": "Totais_label"}, inplace=True)
    df.rename(columns={"Version":"Conta"}, inplace=True)


rename_column()




df = df.drop(columns=['Fct - Previous'])

df.iloc[:, 0].fillna(0)



df.drop(df[(df.Totais_label != "Totais") & (df.Totais_label.index > 2)].index, inplace=True)

df = df.iloc[: , 1:] # remove a coluna Cost Center

df = df.astype({"Conta":str})

df.Conta[df.Conta == 'Non-Labor'] =  'Despesas Operacionais'


for index, conta in enumerate(df.iloc[:, 0]):
    if str(conta)[0].isnumeric() == True:
      print(df.iloc[[index]])



output="result.xlsx"
df.to_excel('result.xlsx', header=True, sheet_name='Opex', index=False)


os.popen(output)

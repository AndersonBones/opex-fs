import pandas as pd
from IPython.display import display, HTML
import numpy as np
import os
import openpyxl as xl
from utils import auto_adjust_column, get_ytd_budget_df
from modules.despesas import Despesas
from modules.ytd_budget import YTD_Budget
import datetime
import re
from modules.formatting import Formatting



class Opex:



    def __init__(self, input_path) -> None:
    
       
        self.input_path = input_path # diretório da base exportada do SAC

        self.wb = xl.load_workbook(self.input_path) 
    
        self.sheet_name=self.wb.sheetnames[0]
    
        self.opex_df = pd.read_excel(input_path, decimal=',', sheet_name=self.sheet_name, index_col=False)


    def remove_rows(self):
        # remove a primeira linha Data
        self.opex_df.drop(index=0, inplace=True)

    def rename_columns(self):
        self.opex_df.rename(columns={"Unnamed: 0": "Totais_label"}, inplace=True)
        self.opex_df.rename(columns={"Version":"Conta"}, inplace=True)
        self.opex_df = self.opex_df.astype({"Conta":str})


    def remove_column(self, column_name): # remove qualquer: nome da coluna como base
        self.opex_df= self.opex_df.drop(columns=[column_name])


    def remove_null_values(self): # substitui null por zero
        self.opex_df.iloc[:, 0].fillna(0)
        



    def remove_centro_custo_rows(self):
        # remove as linhas dos centros de custo  
        self.opex_df.drop(self.opex_df[(self.opex_df.Totais_label != "Totais") & (self.opex_df.Totais_label.index > 2)].index, inplace=True)


    def remove_cost_center(self):
        self.opex_df = self.opex_df.iloc[: , 1:] # remove a coluna Cost Center

        
    def rename_despesas_opecionais_header(self):
        # renomeia o nome da coluna para Despesas Operacionais
        self.opex_df.replace({'Conta':{'Non-Labor':'Despesas Operacionais'}}, inplace=True)




    def insert_ytd_columns(self): # insere as colunas YTD referente ao Budget e Forecast 
        self.opex_df.insert(15, "Δ YTD Budget", np.nan)
        self.opex_df.insert(16, "Δ YTD real", np.nan)
        self.opex_df.insert(17, "Δ YTD", np.nan)
        self.opex_df['Δ Fct x Bdg'] = self.opex_df['Budget'] - self.opex_df['Forecast'] # Orçamento x realizado
        


    def set_ytd_budget_by_conta_razao(self): # configura o YTD do budget de cada conta 

        ytd_budget_df = get_ytd_budget_df() # obtem as contas da base Budget referente a cada conta

        for index, conta_razao in enumerate(self.opex_df['Conta'].str.extract(pat='(^[0-9]{8})').values): # looping na coluna Conta. retorna o valor e indice das contas
            
            for account_index, account in enumerate(ytd_budget_df['Account']): # looping nas contas da base budget

                if conta_razao.tolist()[0] == str(account): # localiza as contas cadastradas na base budget
                    self.opex_df.iat[index, 15] = ytd_budget_df.iat[account_index, 14] # plota o budget acumulado de cada conta localizada na base budget
        


    def set_ytd_budget_total_by_despesa(self): #configura o YTD do budget de cada grupo de despesas 
        despesas = Despesas(self.opex_df)

        ytd_budget = 0
        despesa_index = despesas.get_index_row_by_despesa(column='Conta') # obtém os indices do nome do grupo de despesas 

        for index in range(0, len(despesa_index)-1): #looping pelas contas de cada grupo de despesa 
            ytd_budget+=self.opex_df.iloc[despesa_index[index]-2:despesa_index[index+1]-1, 15].sum()

            # soma os valores entre os intervalos dos indices das linhas do grupo de despesa
            self.opex_df.iloc[despesa_index[index]-2, 15] = self.opex_df.iloc[despesa_index[index]-2:despesa_index[index+1]-1, 15].sum()

        ytd_budget+=self.opex_df.iloc[despesa_index[-1]-2:, 15].sum()
        self.opex_df.iloc[despesa_index[-1]-2, 15] = self.opex_df.iloc[despesa_index[-1]-2:, 15].sum() #soma os valores entre os intervalos dos indices das linhas do ultimo grupo de despesa

        self.opex_df.iloc[2, 15] = ytd_budget # plota o valor YTD BUDGET na linha de Despesas Operacionais 
        

    def set_delta_ytd_fct(self): # configura o forecast acumulado
        first_month_index_safra = 4
        current_month_index = datetime.date.today().month - first_month_index_safra

        self.opex_df['Δ YTD real'] = self.opex_df.iloc[2:, 3:current_month_index+3].sum(skipna=True, axis=1) # refatorar urgentemente


    def set_ytd(self):
        self.opex_df['Δ YTD'] = self.opex_df['Δ YTD real'] - self.opex_df['Δ YTD Budget'] 
        self.opex_df.replace({'Percentual':{np.inf:0}}, inplace=True)


    def set_style(self, export_file_path):
        formatting = Formatting(export_file_path)
        formatting.run()

    def save_file(self, export_file_path):
        
        writer = pd.ExcelWriter(path=export_file_path, engine="xlsxwriter")

        self.opex_df.to_excel(writer, header=True, sheet_name='Opex', index=False)

        auto_adjust_column(self.opex_df, writer)

        self.set_style(export_file_path)
  

       

        
    def run(self):
     
        self.remove_rows()
        self.rename_columns()
        self.remove_column(column_name="Fct - Previous")
        self.remove_column(column_name='Δ Fct x Fct')

        self.remove_centro_custo_rows()
        self.remove_cost_center()
        self.rename_despesas_opecionais_header()

        self.insert_ytd_columns()

        # set YTD Budget
        self.set_ytd_budget_by_conta_razao()
        self.set_ytd_budget_total_by_despesa()

        # set YTD Forecast
        self.set_delta_ytd_fct()
        self.set_ytd()


        # update Base
        ytd_budget = YTD_Budget(self.opex_df)
        ytd_budget.update_ytd_budget()
        

        

        


       

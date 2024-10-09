import pandas as pd
from IPython.display import display, HTML
import numpy as np
import os
from utils import auto_adjust_column, get_budget_dataframe
from modules.accounts import Accounts
import datetime
import re



class Opex:



    def __init__(self, path) -> None:
        self.path = path 
        self.df= pd.read_excel(path, decimal=',', sheet_name="M_OPEX", index_col=False, )


    def remove_rows(self):
        # remove a primeira linha Data
        self.df.drop(index=0, inplace=True)


    def rename_column(self): #renomeia o nome das colunas
        self.df.rename(columns={"Unnamed: 0": "Totais_label"}, inplace=True)
        self.df.rename(columns={"Version":"Conta"}, inplace=True)
    
    def set_str_type_column(self): # altera o tipo de dados da coluna Conta
        self.df = self.df.astype({"Conta":str})


    def remove_column(self, column_name): # remove qualquer: nome da coluna como base
        """
            :column_name: nome da coluna
        """
        self.df= self.df.drop(columns=[column_name])


    def delta_fct_x_bgd(self): # configura o delta do Budget x Forecast
        self.df['Δ Fct x Bdg'] = self.df['Budget'] - self.df['Forecast'] # Orçamento x realizado


    def remove_null_values(self): # remove valores nulos
        self.df.iloc[:, 0].fillna(0)



    def remove_cc_rows(self):
        # remove as linhas dos centros de custo  
        self.df.drop(self.df[(self.df.Totais_label != "Totais") & (self.df.Totais_label.index > 2)].index, inplace=True)


    def remove_cost_center(self):
        self.df = self.df.iloc[: , 1:] # remove a coluna Cost Center

        
    def rename_despesas_opecionais_header(self):
        # renomeia o nome da coluna para Despesas Operacionais
        self.df.Conta[self.df.Conta == 'Non-Labor'] =  'Despesas Operacionais'


    


    def insert_ytd_columns(self): # insere as colunas YTD referente ao Budget e Forecast 
        self.df.insert(15, "Δ YTD Budget", np.nan)
        self.df.insert(16, "Δ YTD real", np.nan)
        self.df.insert(17, "Δ YTD", np.nan)


    def set_delta_ytd_bugdet_subtotal(self): # configura o YTD do budget de cada conta 

        ytd_budget = get_budget_dataframe() # obtem as contas da base Budget referente a cada conta

        for index, item in enumerate(self.df['Conta'].str.extract(pat='(^[0-9]{8})').values): # looping na coluna Conta. retorna o valor e indice das contas
            for account_index, account in enumerate(ytd_budget['Account']): # looping nas contas da base budget
                if item.tolist()[0] == str(account): # localiza as contas cadastradas na base budget
                    self.df.iat[index, 15] = ytd_budget.iat[account_index, 14] # plota o budget acumulado de cada conta localizada na base budget
        


    def set_delta_ytd_bugdet_total(self): #configura o YTD do budget de cada grupo de despesas 
        accounts = Accounts(self.df)
        account_total = accounts.get_total_bgd_fct_by_account() # obtem a identificação e os valores do grupo de desepesas 

        account_label = account_total['account_label'] # obtem a identificação do grupo de desepesas
        index_list = []

        for account in account_label:   
            index_list.append(self.df.loc[self.df['Conta'] == account, 'Conta'].index[0]) # indices das linhas dos grupos de despesas

        index_list = index_list[3:]
        for index in range(0, len(index_list)-1):
            # soma os valores entre os intervalos dos indices das linhas do grupo de despesa
            self.df.iloc[index_list[index]-2, 15] = self.df.iloc[index_list[index]-2:index_list[index+1]-1, 15].sum()


        self.df.iloc[index_list[-1]-2, 15] = self.df.iloc[index_list[-1]-2:, 15].sum() #soma os valores entre os intervalos dos indices das linhas do ultimo grupo de despesa

    # def set_delta_ytd_bugdet_total(self):
    #     account_label_pattern = re.compile('^[a-zA-Z]{5,90}')

    #     self.df.loc[self.df['Conta'].str.contains(pat=account_label_pattern, regex=True) == True, 'Δ YTD Budget']

        

    def set_delta_ytd_fct(self): # configura o forecast acumulado
        account_label_pattern = re.compile('^[0-9]{8} +[a-zA-Z]{5,90}')

        first_month_index_safra = 4
        current_month_index = datetime.date.today().month - first_month_index_safra

        self.df['Δ YTD real'] = self.df.iloc[3:, 3:current_month_index+3].sum(skipna=True, axis=1) # refatorar urgentemente


            
    def set_ytd(self):
        self.df['Δ YTD'] = self.df['Δ YTD Budget'] - self.df['Δ YTD real']


    def save_file(self):
        output="Opex.xlsx"    

        writer = pd.ExcelWriter(path=output, engine="xlsxwriter")

        self.df.to_excel(writer, header=True, sheet_name='Opex', index=False)

        ytd_budget = get_budget_dataframe()

        ytd_budget.to_excel(writer, header=True, sheet_name='Relatório', index=False)

        auto_adjust_column(self.df, writer)
       

        
    def run(self):
        self.remove_rows()
        self.rename_column()
        self.remove_column(column_name="Fct - Previous")
        self.remove_column(column_name='Δ Fct x Fct')
        self.set_str_type_column()
 
        self.remove_cc_rows()

        self.remove_cost_center()

        self.rename_despesas_opecionais_header()

        self.delta_fct_x_bgd()

        self.insert_ytd_columns()

        self.set_delta_ytd_bugdet_subtotal()
        self.set_delta_ytd_bugdet_total()

        self.set_delta_ytd_fct()



        self.set_ytd()

  

        self.save_file()

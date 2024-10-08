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


    def rename_column(self):
        self.df.rename(columns={"Unnamed: 0": "Totais_label"}, inplace=True)
        self.df.rename(columns={"Version":"Conta"}, inplace=True)
    
    def set_str_type_column(self):
        self.df = self.df.astype({"Conta":str})


    def remove_column(self, column_name):
        """
            :column_name: nome da coluna
        """
        self.df= self.df.drop(columns=[column_name])


    def delta_fct_bgd(self):
        self.df['Δ Fct x Bdg'] = self.df['Budget'] - self.df['Forecast']


    def remove_null(self):
        self.df.iloc[:, 0].fillna(0)



    def remove_cc_rows(self):
        # remove cc rows 
        self.df.drop(self.df[(self.df.Totais_label != "Totais") & (self.df.Totais_label.index > 2)].index, inplace=True)


    def remove_cost_center(self):
        self.df = self.df.iloc[: , 1:] # remove a coluna Cost Center

        
    def set_despesas_opecionais_header(self):
        # set Despesas Operacionais header
        self.df.Conta[self.df.Conta == 'Non-Labor'] =  'Despesas Operacionais'


    


    def insert_ytd_columns(self):
        self.df.insert(15, "Δ YTD Budget", np.nan)
        self.df.insert(16, "Δ YTD real", np.nan)
        self.df.insert(17, "Δ YTD", np.nan)


    def set_delta_ytd_bugdet_subtotal(self):
        accounts = Accounts(self.df)
        account_subtotal = accounts.get_subtotal_bgd_fct_by_account()
        account_id_pattern = re.compile('^[0-9]{8}')
        account_total_label_pattern = re.compile('^[a-zA-Z]')

        account_total = accounts.get_total_bgd_fct_by_account()

        first_month_index_safra = 4
        current_month_index = datetime.date.today().month - first_month_index_safra

        ytd_budget = get_budget_dataframe()


        for index, item in enumerate(self.df['Conta'].str.extract(pat='(^[0-9]{8})').values):
            for account_index, account in enumerate(ytd_budget['Account']):
                
                if item.tolist()[0] == str(account):
                    self.df.iat[index, 15] = ytd_budget.iat[account_index, 14]
        

    def split_df_by_account(self):
        dfs = []
        accounts = Accounts(self.df)
        account_total = accounts.get_total_bgd_fct_by_account()

        account_label = account_total['account_label']
        index_list = []

        for account in account_label:
          
            index_list.append(self.df.loc[self.df['Conta'] == account, 'Conta'].index[0])

            
            print(self.df.loc[self.df['Conta'] == account, 'Conta']) # dividir em varios dataframes se baseando no indice dos totais
                                                                     # apos isso, obter apenas a coluna do budget acumulado, e por fim somar os valores
        
        
        


    def set_delta_ytd_bugdet_total(self):
        account_label_pattern = re.compile('^[a-zA-Z]{5,90}')

        self.df.loc[self.df['Conta'].str.contains(pat=account_label_pattern, regex=True) == True, 'Δ YTD Budget']

        

    def set_delta_ytd_budget2(self):
       
        accounts = Accounts(self.df)
        account_subtotal = accounts.get_subtotal_bgd_fct_by_account()
        account_id_pattern = re.compile('^[0-9]{8}')
        account_total_label_pattern = re.compile('^[a-zA-Z]')

        account_total = accounts.get_total_bgd_fct_by_account()

        first_month_index_safra = 4
        current_month_index = datetime.date.today().month - first_month_index_safra
        ytd_budget = get_budget_dataframe()


        self.df.loc[self.df['Conta'].astype(str).str.extract(pat='(^[0-9]{8})') == ytd_budget['Account'].astype(str), 'Δ YTD Budget'] = 'ss'

        for index, item in enumerate(ytd_budget['Account']):

            ytd_budget_value = ytd_budget.loc[ytd_budget['Account'] == item, 'YTD Budget'].tolist()

           

            # account = self.df.loc[self.df['Conta'].str.extract(pat='(^[0-9]{8})') == ytd_budget['Account'].astype(str).str.extract(pat='(^[0-9]{8})'), 'Δ YTD Budget']

           
            self.df.loc[self.df['Conta'].str.extract(pat='(^[0-9]{8})') == ytd_budget['Account'], 'Δ YTD Budget'] = ytd_budget_value[0] if len(ytd_budget_value) > 0 else 0
        

           
        for index, item in enumerate(account_subtotal['account_label']):
             self.df.loc[self.df['Conta'] == item, 'Δ YTD Budget'] = (account_subtotal['account_bdg_total'][index] / 12) * current_month_index

        for index, item in enumerate(account_total['account_label']):
            self.df.loc[self.df['Conta'] == item, 'Δ YTD Budget'] = (account_total['account_bdg_total'][index] / 12) * current_month_index


    def set_delta_ytd_fct(self):
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
        self.remove_null()
        self.remove_cc_rows()

        self.remove_cost_center()

        self.set_despesas_opecionais_header()

        self.delta_fct_bgd()

        self.insert_ytd_columns()

        self.set_delta_ytd_bugdet_subtotal()
        self.set_delta_ytd_bugdet_total()

        self.set_delta_ytd_fct()

        self.split_df_by_account()

        self.set_ytd()

  

        self.save_file()

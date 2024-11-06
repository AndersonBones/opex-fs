import re
import numpy as np
import pandas as pd
from utils import regex_pattern

class Despesas:
    def __init__(self, df) -> None:
        self.df = df


    def get_centro_custo(self):# obtém os dados dos centros de custo
        pattern = regex_pattern("Centro de custo")  # regex pattern

        try:
            centro_custo_df = pd.DataFrame(data=self.df.iloc[:,0].str.extract(pattern).dropna()).rename(columns={0:"Centro de custo"})
            return centro_custo_df
        
        except Exception as err:
            return None
        


    def get_index_row_by_despesa(self, column="Conta"):
        despesa_pattern = regex_pattern("Despesa") # obtém o regex referente o nome do grupo de despesa

        try:
            index_list = self.df.loc[self.df['Conta'].str.contains(despesa_pattern, regex=True) == True, column].index.tolist() #
            
            return index_list[3:]
        except Exception as err:
            return None
    

    def get_index_row_by_conta_razao(self, column="Conta"):
        conta_razao_pattern = regex_pattern("Conta razão num") # obtém o regex referente o nome do grupo de despesa
    
        try:
            index_list = self.df.loc[self.df['Conta'].str.contains(conta_razao_pattern, regex=True) == True, column].index.tolist() #
            return index_list
        except Exception as err:
            return None
    

    def get_despesa(self): 
        
        # columns 
        despesa = []
        budget = []
        conta_razao = []
        forecast = []
        ytd_delta_budget = []
        ytd_delta_realizado = []

        despesa_index_list = self.get_index_row_by_despesa('Conta') # obtém os indices do nome do grupo de despesas 


        df_despesa = pd.DataFrame(data={ # declara o dataframe
            "Despesa":[], 
            "Conta":[], 
            "Budget":[], 
            "Forecast":[], 
            "Fct x Bdg":[], 
            "Percentual":[],
            "Δ YTD Budget":[],
            "Δ YTD real":[],
        })


        for index in range(0, len(despesa_index_list)-1): #looping pela lista de indices 

            # soma os valores dentro do intervalo do grupo de despesa
            for idx, acc in enumerate(self.df.iloc[despesa_index_list[index]-1:despesa_index_list[index+1]-2, 0].tolist()): #looping pelas contas de cada grupo de despesa

                despesa.append(self.df.iloc[despesa_index_list[index]-2, 0]) # armazena o nome do grupo de despesa que se repete ate que o looping passe por outro grupo de despesa
                conta_razao.append(acc) # armazena cada conta razao referente aquela despesa

                budget.append(self.df.iloc[despesa_index_list[index]-1:despesa_index_list[index+1]-2, 1].tolist()[idx]) # armazena o budget definido para aquela conta razao
                forecast.append(self.df.iloc[despesa_index_list[index]-1:despesa_index_list[index+1]-2, 2].tolist()[idx])# armazena o forecast previsto para aquela conta razao

                ytd_delta_budget.append(self.df.iloc[despesa_index_list[index]-1:despesa_index_list[index+1]-2, 15].tolist()[idx]) # armazena o budget acumulado do ano para aquela conta razao
                ytd_delta_realizado.append(self.df.iloc[despesa_index_list[index]-1:despesa_index_list[index+1]-2, 16].tolist()[idx])# armazena o realizado acumulado do ano para aquela conta razao
        

        pd.set_option('future.no_silent_downcasting', True)

        # looping pelo ultimo grupo de despesa
        for idx, acc in enumerate(self.df.iloc[despesa_index_list[-1]-1:, 0]):
            despesa.append(self.df.iloc[despesa_index_list[-1]-2, 0])
            conta_razao.append(self.df.loc[self.df['Conta'] == acc, 'Conta'].tolist()[0])

            budget.append(self.df.loc[self.df['Conta'] == acc, 'Budget'].fillna(0).tolist()[0])
            forecast.append(self.df.loc[self.df['Conta'] == acc, 'Forecast'].fillna(0).tolist()[0])

            ytd_delta_budget.append(self.df.loc[self.df['Conta'] == acc, 'Δ YTD Budget'].fillna(0).tolist()[0])
            ytd_delta_realizado.append(self.df.loc[self.df['Conta'] == acc, 'Δ YTD real'].fillna(0).tolist()[0])
            

        #armazena as listas em sua respectiva colunas do dataframe
        df_despesa['Despesa'] = despesa
        df_despesa['Conta'] = conta_razao
        df_despesa['Budget'] = budget
        df_despesa['Forecast'] = forecast

        df_despesa['Δ YTD Budget'] = ytd_delta_budget
        df_despesa['Δ YTD real'] = ytd_delta_realizado

        df_despesa.fillna(0, inplace=True) # substitui os valores nulos por zero

        df_despesa['Fct x Bdg'] = df_despesa['Budget'] - df_despesa['Forecast']
        
        df_despesa['Percentual'] = df_despesa['Forecast'].div(df_despesa['Budget'])
    
        df_despesa.replace({'Percentual':{np.inf:0}}, inplace=True)
        df_despesa.replace({'Percentual':{-np.inf:0}}, inplace=True)

        df_despesa.fillna(0, inplace=True) # substitui os valores nulos por zero

        
        return df_despesa
    


    def get_column_from_despesa(self, column="Conta"):
        despesa_pattern = regex_pattern('Despesa')
        
        despesa_column = self.df.loc[self.df['Conta'].str.contains(despesa_pattern, regex=True) == True, column]

        if despesa_column.dtypes == float:
            return despesa_column.replace(np.nan, 0).tolist()
        else:
            return despesa_column.tolist()



    def get_conta_razao(self):
        regex_conta_razao = regex_pattern("Conta razão num")
        return self.opex_df['Conta'].str.extract(pat=regex_conta_razao)
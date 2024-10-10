import re
import numpy as np
import pandas as pd

class Accounts:
    def __init__(self, df) -> None:
        self.df = df


    def get_cc(self):
        """
        """

        centro_custo = {
            "cc_label":[],
            "cc":[]
        }

        # regex pattern
        pattern_cc = f'[0-9]+BIO'
        pattern_cc_label = f'[0-9]+BIO +Biomassa'


        # get cc
        for cc in self.df['Totais_label'].iloc[0:].values:
            if cc is not None and type(cc) == str and cc[0].isnumeric() == True:
                centro_custo['cc_label'].append(cc)

                # get cc by regex pattern
                extract_cc = re.findall(pattern_cc, cc, flags=re.IGNORECASE)
                centro_custo['cc'].append(extract_cc[0])

                # get cc label by regex pattern
                extract_cc_label = re.findall(pattern_cc_label, cc, flags=re.IGNORECASE)
                centro_custo['cc_label'].append(extract_cc_label[0])

        return centro_custo


    def get_despesas_category(self):


        account_total = self.get_total_bgd_fct_by_account() # obtem a identificação e os valores do grupo de desepesas 

        account_label = account_total['account_label'] # obtem a identificação do grupo de desepesas
        index_list = []

        for account in account_label:   
            index_list.append(self.df.loc[self.df['Conta'] == account, 'Conta'].index[0]) # indices das linhas dos grupos de despesas
        index_list = index_list[3:]


        category = {"Despesa":[], "Conta":[], "Budget":[], "Forecast":[], "Fct x Bdg":[], "Percentual":[]}
        df_category = pd.DataFrame(data=category)



        despesas_category = []
        budget_by_account = []
        account_by_category = []
        forecast_by_account = []

        for index in range(0, len(index_list)-1):
            # soma os valores entre os intervalos dos indices das linhas do grupo de despesa

            for idx, acc in enumerate(self.df.iloc[index_list[index]-1:index_list[index+1]-2, 0].tolist()):
                despesas_category.append(self.df.iloc[index_list[index]-2, 0])
                account_by_category.append(acc)
                budget_by_account.append(self.df.iloc[index_list[index]-1:index_list[index+1]-2, 1].tolist()[idx])
                forecast_by_account.append(self.df.iloc[index_list[index]-1:index_list[index+1]-2, 2].tolist()[idx])

        
        df_category['Despesa'] = despesas_category
        df_category['Conta'] = account_by_category
        df_category['Budget'] = budget_by_account
        df_category['Forecast'] = forecast_by_account
        df_category.fillna(0, inplace=True)

        df_category['Fct x Bdg'] = df_category['Budget'] - df_category['Forecast']
        df_category['Percentual'] = (df_category['Forecast'] / df_category['Budget'])

        df_category['Percentual'].replace([np.inf, -np.inf], 0, inplace=True)

        df_category.style.format({
            'Percentual': '{:,.2%}'.format,
        })






        return df_category
    

    def get_total_bgd_fct_by_account(self):
        account_total_label_pattern = re.compile('^[A-Za-záàâãéèêíïóôõöúçñÁÀÂÃÉÈÍÏÓÔÕÖÚÇÑ]')
        
        account_total = {
            "account_label":[],
            "account_bdg_total":[],
            "account_fct_total":[]
        }

        account_total["account_label"] = self.df.loc[self.df['Conta'].str.contains(account_total_label_pattern, regex=True) == True, 'Conta'].tolist()
        account_total["account_bdg_total"] = self.df.loc[self.df['Conta'].str.contains(account_total_label_pattern, regex=True) == True, 'Budget'].replace(np.nan, 0).tolist()
        account_total["account_fct_total"] = self.df.loc[self.df['Conta'].str.contains(account_total_label_pattern, regex=True) == True, 'Forecast'].replace(np.nan, 0).tolist()

        return account_total


    def get_subtotal_bgd_fct_by_account(self): # get subtotal values
        account_label_pattern = re.compile('^[0-9]{8} +[a-zA-Z]{5,90}') # 
        account_label_total_pattern = re.compile('^[a-zA-Z]{1,90}')
        account_id_pattern = re.compile('^[0-9]{8}')
        

        account_subtotal = {
            "account_label":[],
            "account_id":[],
            "account_bdg_total":[],
            "account_fct_total":[]
        }


        for value in self.df.loc[self.df['Conta'].str.contains(account_label_pattern, regex=True) == True, 'Conta'].tolist():
            account_subtotal["account_id"].append(re.findall(account_id_pattern, value)[0])

        account_subtotal['account_label'] = self.df.loc[self.df['Conta'].str.contains(account_label_pattern, regex=True) == True, 'Conta'].tolist()
        account_subtotal["account_bdg_total"] = self.df.loc[self.df['Conta'].str.contains(account_label_pattern, regex=True) == True, 'Budget'].replace(np.nan, 0).tolist()
        account_subtotal["account_fct_total"] = self.df.loc[self.df['Conta'].str.contains(account_label_pattern, regex=True) == True, 'Forecast'].replace(np.nan, 0).tolist()
    

        return account_subtotal
    

    def get_bdg_total(self):
        account_subtotal_label = re.compile('^[0-9]{8} +[a-zA-Z]{5,90}')
        account_total_label = re.compile('^[a-zA-Z]')
        account = {
            "account_label":[],
            "account_bdg_total":[],
            "account_fct_total":[]
        }

        account["account_label"] = self.df.loc[self.df['Conta'].str.contains(account_subtotal_label, regex=True) == True, 'Conta'].tolist()
        account["account_bdg_total"] = self.df.loc[self.df['Conta'].str.contains(account_subtotal_label, regex=True) == True, 'Budget'].replace(np.nan, 0).tolist()
        account["account_fct_total"] = self.df.loc[self.df['Conta'].str.contains(account_subtotal_label, regex=True) == True, 'Forecast'].replace(np.nan, 0).tolist()

        return account
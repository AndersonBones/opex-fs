import re
import numpy as np


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


    def despesas_category(self):

        # account category
        despesas_category = {
            "Manutenção":[
                    "40120101 Manutenção das instalações",
                    "40120102 Manutenção de veículos",
                    "40120103 Manutenção de máquinas e equipamentos"
                ],

            "Aluguéis e locações":[
                "40120201 Aluguel de veículos"
            ],

            "Despesas de viagens":[
                "40120401 Locomoção terrestre",
                "40120402 Passagens aéreas",
                "40120403 Despesa com hospedagem",
                "40120404 Refeições - Despesa de viagem",

            ],

            "Água, energia e comunicação":[
                "40120501 Energia elétrica",
                "40120503 Serviço de telefonia móvel"
            ],

            "Despesas com serviços de terceiros":[
                "40120601 Serviços de auditoria",
                "40120602 Serviços de consultoria",
                "40120605 Serviços de armazenagem",
                "40120603 Serviços de assessoria jurídica",
                "40120606 Serviços de industrialização Biomassa (picagem)",
                "40120612 Serviços Prestados por Pessoa Jurídica",
                "40120613 Serviços Manutenção e Assistência de Sistemas",
                "40120614 Serviços Análises"
            ],

            "Despesas tributárias":[
                "40120701 Licenciamento de veículos (IPVA e DPVAT)",
                "40120705 Despesas cartorárias",
                "40120706 Licenças e alvarás",
                "40120707 Taxas diversas",
                "40120801 Seguro de veículos",
                "40120804 Outros seguros",
                "40120901 Despesas com fretes"
            ],

            "Despesas gerais":[
                "40121102 Material de escritório",
                "40121104 Uniforme",
                "40121107 Bens de pequeno valor",
                "40121108 Propaganda e publicidade",
                "40121110 Equipamentos de Proteção Individual",
                "40121113 Multas Indedutíveis",
                "40121119 Licencas de uso",
                "40121120 Eventos e Confraternizações",
                "40121122 Materiais de suprimentos de informatica"
            ]
        }



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
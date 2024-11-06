from modules.despesas import Despesas
import pandas as pd
from utils import auto_adjust_column
import os

class YTD_Budget:
    def __init__(self, df):
        self.df = df
        self.ytd_budget_path = r"F:\BIOMASSA\00. Indicadores Gerenciais\03. Budget Biomassa\2024-2025\OPEX - Gasto Fixo\Budget 24'25\Base"
        self.ytd_budget_path = os.path.join(self.ytd_budget_path, "Opex - Budget.xlsx")


    def update_ytd_budget(self):
        # get desepesas category from Accounts
        despesas = Despesas(self.df)
        df_ytd_budget = despesas.get_despesa()


        
        try:
            with pd.ExcelWriter(self.ytd_budget_path, engine='openpyxl', mode='a') as writer: # faz a leitura da base YTD Budget
                
                wb = writer.book #obtém o workbook da planilha base

                if "Despesas e Contas" in wb.sheetnames: #verifica se existe a aba 'Despesas e Contas'
              
                    wb.remove(wb["Despesas e Contas"]) # remove a aba 'Despesas e Contas' caso exista 
                
                pd.set_option('float_format', '{:.3f}'.format) #

                df_ytd_budget.to_excel(writer, sheet_name='Despesas e Contas', index=False )
        except Exception as ex:
            print(f"Error:'{ex}'.")
            exit()
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            exit()

        # 
  
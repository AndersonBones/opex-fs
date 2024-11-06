import pandas as pd
import datetime
import os

def auto_adjust_column(df, writer):
    
    for column in df:
        column_length = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        
        for sheet in writer.sheets:
            writer.sheets[sheet].set_column(col_idx, col_idx, column_length+1.5)

    writer.close()
    


def get_ytd_budget_df():

    opex_path_base = r"F:\BIOMASSA\00. Indicadores Gerenciais\03. Budget Biomassa\2024-2025\OPEX - Gasto Fixo\Budget 24'25\Base"
    
    path=os.path.join(opex_path_base, "Opex - Budget.xlsx")

    pd.options.display.float_format = '{:.2f}'.format

    first_month_index_safra = 4
    current_month_index = datetime.date.today().month - first_month_index_safra

    df = pd.read_excel(path, sheet_name="Budget")

    df['YTD Budget'] = df.iloc[:, 1:current_month_index+1].sum(axis=1)

  
    return df


def regex_pattern(regex_label):
    pattern = {
        "Centro de custo":"(^[0-9]+BIO)",
        "Despesa":"^[A-Za-záàâãéèêíïóôõöúçñÁÀÂÃÉÈÍÏÓÔÕÖÚÇÑ]",
        "Conta razão":"(^[0-9]{8} +[a-zA-Z]{5,90})",
        "Conta razão num":"(^[0-9]{8})"

    }


    try:
        return pattern[regex_label]
    except Exception as err:
        return None


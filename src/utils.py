import pandas as pd
import datetime
import os

def auto_adjust_column(df, writer):

    for column in df:
        column_length = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        writer.sheets["Opex"].set_column(col_idx, col_idx, column_length+1)

    writer.close()
    


def get_budget_dataframe():

    cwd = os.getcwd()
    path=os.path.join(cwd, "data", "budget.xlsx")

    pd.options.display.float_format = '{:.2f}'.format

    first_month_index_safra = 4
    current_month_index = datetime.date.today().month - first_month_index_safra

    df = pd.read_excel(path, sheet_name="budget")
    
    df['YTD Budget'] = df.iloc[:, 1:current_month_index+1].sum(axis=1)

  
    return df


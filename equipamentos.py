import pandas as pd
import copy
from consumo import converter

def criar_equip(itens,writer):
    new_itens = copy.deepcopy(itens)
    i=0
    while i < len(new_itens['Potência']):
        new_itens["Potência"][i] = converter(new_itens['Potência'][i])
        new_itens["Fator de Potência"][i] = converter(new_itens['Fator de Potência'][i])
        new_itens["Quantidade"][i] = converter(new_itens['Quantidade'][i])
        i+=1
    df_equipamentos = pd.DataFrame(new_itens)
    df_equipamentos.to_excel(writer, sheet_name="Equipamentos", startrow=1, header=False, index=False)

    workbook = writer.book
    worksheet = writer.sheets["Equipamentos"]
    (max_row, max_col) = df_equipamentos.shape
    column_settings = [{"header": column} for column in df_equipamentos.columns]
    worksheet.add_table(0, 0, max_row, max_col - 1, {"columns": column_settings})
    worksheet.set_column(0, max_col - 1, 12)
    worksheet.autofit()
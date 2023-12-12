import pandas as pd

def criar_equip(itens,writer):
    df_equipamentos = pd.DataFrame(itens)
    df_equipamentos.to_excel(writer, sheet_name="Equipamentos", startrow=1, header=False, index=False)

    workbook = writer.book
    worksheet = writer.sheets["Equipamentos"]
    (max_row, max_col) = df_equipamentos.shape
    column_settings = [{"header": column} for column in df_equipamentos.columns]
    worksheet.add_table(0, 0, max_row, max_col - 1, {"columns": column_settings})
    worksheet.set_column(0, max_col - 1, 12)
    worksheet.autofit()
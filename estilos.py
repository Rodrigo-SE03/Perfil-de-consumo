import pandas as pd
def skip(val,num=1):
    return chr(ord(val)+num)

def consumo_equip_style(worksheet,workbook,categoria):
    merge_format = workbook.add_format(
    {
        "bold": 1,
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "fg_color": "#4f81bd",
        "font_color": "white",
        "border_color": "white"
    }
    )
    empty_format = workbook.add_format(
    {
        "bold": 1,
        "right": 1,
        "fg_color": "#4f81bd",
        "border_color": "white"
    }
    )
    if categoria == "Verde" or categoria == "Azul":
        worksheet.merge_range("C1:E1","Horas de Utilização Diária",merge_format)
        worksheet.merge_range("F1:H1","Consumo Diário Típico (kWh)",merge_format)
    elif categoria == 'Branca':
        worksheet.merge_range("C1:F1","Horas de Utilização Diária",merge_format)
        worksheet.merge_range("G1:J1","Consumo Diário Típico (kWh)",merge_format)
    else:
        worksheet.write_blank(0,2,'',empty_format)
        worksheet.write_blank(0,3,'',empty_format)
    worksheet.write_blank(0,0,'',empty_format)
    worksheet.write_blank(0,1,'',empty_format)

def tabelas_geral(worksheet,workbook,categoria,tarifas,custo,dias):
    pos = "G1"

    col = pos[0]
    row = int(pos[1])

    merge_format = workbook.add_format(
    {
        "bold": 1,
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "fg_color": "#4f81bd",
        "font_color": "white"
    }
    )
    if categoria == "Branca":
        worksheet.merge_range(f"{skip(col)}{row}:{skip(col,3)}{row}","Valores Diários",merge_format)
        worksheet.merge_range(f"{skip(col,6)}{row}:{skip(col,8)}{row}","Valores Mensais",merge_format)
    else:
        if categoria == 'Azul' or categoria == 'Verde': col = skip(col,2)
        worksheet.merge_range(f"{col}{row}:{skip(col,2)}{row}","Valores Diários",merge_format)
        worksheet.merge_range(f"{skip(col,5)}{row}:{skip(col,7)}{row}","Valores Mensais",merge_format)

    border = workbook.add_format({'border':1})
    rs_format = workbook.add_format({'num_format':'R$ #,##0.00','border':1})
    pot_format = workbook.add_format({'num_format':'#,##0.00 "kW"','border':1})
    nrg_format = workbook.add_format({'num_format':'#,##0.00 "kWh"','border':1})
    
    if categoria == 'Convencional':
        worksheet.write(f'{col}{row+1}',"Consumo",border)
        worksheet.write(f'{skip(col)}{row+1}',custo[0]/tarifas[0],nrg_format)
        worksheet.write(f'{skip(col,2)}{row+1}',custo[0],rs_format)

        worksheet.write(f'{skip(col,5)}{row+1}',"Consumo",border)
        worksheet.write(f'{skip(col,6)}{row+1}',dias*custo[0]/tarifas[0],nrg_format)
        worksheet.write(f'{skip(col,7)}{row+1}',dias*custo[0],rs_format)

    if categoria == 'Branca':
        worksheet.write(f'{skip(col)}{row+1}',"Consumo FP",border)
        worksheet.write(f'{skip(col,2)}{row+1}',custo[0]/tarifas[0],nrg_format)
        worksheet.write(f'{skip(col,3)}{row+1}',custo[0],rs_format)
        worksheet.write(f'{skip(col)}{row+2}',"Consumo I",border)
        worksheet.write(f'{skip(col,2)}{row+2}',custo[2]/tarifas[2],nrg_format)
        worksheet.write(f'{skip(col,3)}{row+2}',custo[2],rs_format)
        worksheet.write(f'{skip(col)}{row+3}',"Consumo P",border)
        worksheet.write(f'{skip(col,2)}{row+3}',custo[1]/tarifas[1],nrg_format)
        worksheet.write(f'{skip(col,3)}{row+3}',custo[1],rs_format)

        worksheet.write(f'{skip(col,6)}{row+1}',"Consumo FP",border)
        worksheet.write(f'{skip(col,7)}{row+1}',dias*custo[0]/tarifas[0],nrg_format)
        worksheet.write(f'{skip(col,8)}{row+1}',dias*custo[0],rs_format)
        worksheet.write(f'{skip(col,6)}{row+2}',"Consumo I",border)
        worksheet.write(f'{skip(col,7)}{row+2}',dias*custo[2]/tarifas[2],nrg_format)
        worksheet.write(f'{skip(col,8)}{row+2}',dias*custo[2],rs_format)
        worksheet.write(f'{skip(col,6)}{row+3}',"Consumo P",border)
        worksheet.write(f'{skip(col,7)}{row+3}',dias*custo[1]/tarifas[1],nrg_format)
        worksheet.write(f'{skip(col,8)}{row+3}',dias*custo[1],rs_format)
        worksheet.write(f'{skip(col,6)}{row+4}',"Total",border)
        worksheet.write(f'{skip(col,7)}{row+4}',"-",border)
        worksheet.write_formula(f'{skip(col,8)}{row+4}',f"=SUM({skip(col,8)}{row+1}:{skip(col,8)}{row+3})",rs_format)

    elif categoria == 'Verde' or categoria == 'Azul':
        worksheet.write(f'{col}{row+1}',"Consumo FP",border)
        worksheet.write(f'{skip(col)}{row+1}',custo[0],nrg_format)
        worksheet.write(f'{skip(col,2)}{row+1}',custo[2],rs_format)
        worksheet.write(f'{col}{row+2}',"Consumo P",border)
        worksheet.write(f'{skip(col)}{row+2}',custo[1],nrg_format)
        worksheet.write(f'{skip(col,2)}{row+2}',custo[3],rs_format)

        worksheet.write(f'{skip(col,5)}{row+1}',"Consumo FP",border)
        worksheet.write(f'{skip(col,6)}{row+1}',dias*custo[0],nrg_format)
        worksheet.write(f'{skip(col,7)}{row+1}',dias*custo[2],rs_format)
        worksheet.write(f'{skip(col,5)}{row+2}',"Consumo P",border)
        worksheet.write(f'{skip(col,6)}{row+2}',dias*custo[1],nrg_format)
        worksheet.write(f'{skip(col,7)}{row+2}',dias*custo[3],rs_format)
        if categoria == 'Verde':
            worksheet.write(f'{col}{row+3}',"Demanda",border)
            worksheet.write(f'{skip(col)}{row+3}',custo[4]/tarifas[2],pot_format)
            worksheet.write(f'{skip(col,2)}{row+3}',custo[4],rs_format)

            worksheet.write(f'{skip(col,5)}{row+3}',"Demanda",border)
            worksheet.write(f'{skip(col,6)}{row+3}',custo[4]/tarifas[2],pot_format)
            worksheet.write(f'{skip(col,7)}{row+3}',custo[4],rs_format)
            worksheet.write(f'{skip(col,5)}{row+4}',"Total",border)
            worksheet.write(f'{skip(col,6)}{row+4}',"-",border)
            worksheet.write_formula(f'{skip(col,7)}{row+4}',f"=SUM({skip(col,7)}{row+1}:{skip(col,7)}{row+3})",rs_format)
        else:
            worksheet.write(f'{col}{row+3}',"Demanda FP",border)
            worksheet.write(f'{skip(col)}{row+3}',custo[4],pot_format)
            worksheet.write(f'{skip(col,2)}{row+3}',custo[4]*tarifas[2],rs_format)
            worksheet.write(f'{col}{row+4}',"Demanda P",border)
            worksheet.write(f'{skip(col)}{row+4}',custo[5],pot_format)
            worksheet.write(f'{skip(col,2)}{row+4}',custo[5]*tarifas[3],rs_format)
            
            worksheet.write(f'{skip(col,5)}{row+3}',"Demanda FP",border)
            worksheet.write(f'{skip(col,6)}{row+3}',custo[4],pot_format)
            worksheet.write(f'{skip(col,7)}{row+3}',custo[4]*tarifas[2],rs_format)
            worksheet.write(f'{skip(col,5)}{row+4}',"Demanda P",border)
            worksheet.write(f'{skip(col,6)}{row+4}',custo[5],pot_format)
            worksheet.write(f'{skip(col,7)}{row+4}',custo[5]*tarifas[3],rs_format)
            worksheet.write(f'{skip(col,5)}{row+5}',"Total",border)
            worksheet.write(f'{skip(col,6)}{row+5}',"-",border)
            worksheet.write_formula(f'{skip(col,7)}{row+5}',f"=SUM({skip(col,7)}{row+1}:{skip(col,7)}{row+4})",rs_format)

def tabela_reativos(dict,writer,demr,categoria,consumo_mes):
    df_valores = pd.DataFrame(dict)
    df_valores.to_excel(writer,sheet_name="Reativos", startrow=3, header=False, index=False)
    workbook = writer.book
    worksheet = writer.sheets["Reativos"]
    table_format = workbook.add_format(
    {
        'num_format': '#,##0.00'
    }
    )
    (max_row, max_col) = df_valores.shape
    if categoria == "Verde":
        worksheet.set_column('A:J',10,table_format)
    else:
        worksheet.set_column('A:K',10,table_format)
    column_settings = [{"header": column} for column in df_valores.columns]
    worksheet.add_table(2, 0, max_row+2, max_col - 1, {"columns": column_settings})
    merge_format = workbook.add_format(
    {
        "bold": 1,
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "fg_color": "#4f81bd",
        "font_color": "white",
        "border_color": "white"
    }
    )
    merge_format2 = workbook.add_format(
    {
        "bold": 1,
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "fg_color": "#4f81bd",
        "font_color": "white",
    }
    )
    empty_format = workbook.add_format(
    {
        "bold": 1,
        "right": 1,
        "fg_color": "#4f81bd",
        "border_color": "white"
    }
    )
    rs_format = workbook.add_format({'num_format':'R$ #,##0.00','border':1})
    worksheet.merge_range("A1:A2","",empty_format)
    if categoria == "Verde":
        worksheet.merge_range("B1:D1","Valores Ativos",merge_format)
        worksheet.merge_range("C2:D2","Consumo",merge_format)
        worksheet.merge_range("E1:F2","Valores Reativos",merge_format)
        worksheet.merge_range("G1:H2","Fator de Potência",merge_format)
        worksheet.merge_range("I1:J1","Valores Calculados",merge_format)
        worksheet.write("B2","Demanda",merge_format)
        worksheet.write("I2","Demanda",merge_format)
        worksheet.write("J2","Consumo",merge_format)
        worksheet.merge_range("M1:N1","Acréscimo na Fatura",merge_format2)
        worksheet.write("M2","Demanda",rs_format)
        worksheet.write("N2",demr,rs_format)
        worksheet.write("M3","Consumo",rs_format)
        worksheet.write("N3",consumo_mes,rs_format)
        worksheet.write("M4","Total",rs_format)
        worksheet.write("N4",consumo_mes+demr,rs_format)
    else:
        worksheet.merge_range("B1:E1","Valores Ativos",merge_format)
        worksheet.merge_range("B2:C2","Demanda",merge_format)
        worksheet.merge_range("D2:E2","Consumo",merge_format)
        worksheet.merge_range("F1:G2","Valores Reativos",merge_format)
        worksheet.merge_range("H1:I2","Fator de Potência",merge_format)
        worksheet.merge_range("J1:K1","Valores Calculados",merge_format)
        worksheet.write("J2","Demanda",merge_format)
        worksheet.write("K2","Consumo",merge_format)
        worksheet.merge_range("N1:O1","Acréscimo na Fatura",merge_format2)
        worksheet.write("N2","Demanda FP",rs_format)
        worksheet.write("O2",demr[0],rs_format)
        worksheet.write("N3","Demanda P",rs_format)
        worksheet.write("O3",demr[1],rs_format)
        worksheet.write("N4","Consumo",rs_format)
        worksheet.write("O4",consumo_mes,rs_format)
        worksheet.write("N5","Total",rs_format)
        worksheet.write("O5",consumo_mes+demr[0]+demr[1],rs_format)
         
    worksheet.autofit()

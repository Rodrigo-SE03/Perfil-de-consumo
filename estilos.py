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

    # next = chr(ord(col)+1)

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
        worksheet.write(f'{skip(col)}{row+1}',custo[0]/tarifas[0],nrg_format)
        worksheet.write(f'{skip(col,2)}{row+1}',custo[0],rs_format)
        worksheet.write(f'{col}{row+2}',"Consumo P",border)
        worksheet.write(f'{skip(col)}{row+2}',custo[1]/tarifas[1],nrg_format)
        worksheet.write(f'{skip(col,2)}{row+2}',custo[1],rs_format)

        worksheet.write(f'{skip(col,5)}{row+1}',"Consumo FP",border)
        worksheet.write(f'{skip(col,6)}{row+1}',dias*custo[0]/tarifas[0],nrg_format)
        worksheet.write(f'{skip(col,7)}{row+1}',dias*custo[0],rs_format)
        worksheet.write(f'{skip(col,5)}{row+2}',"Consumo P",border)
        worksheet.write(f'{skip(col,6)}{row+2}',dias*custo[1]/tarifas[1],nrg_format)
        worksheet.write(f'{skip(col,7)}{row+2}',dias*custo[1],rs_format)
        if categoria == 'Verde':
            worksheet.write(f'{col}{row+3}',"Demanda",border)
            worksheet.write(f'{skip(col)}{row+3}',custo[3]/tarifas[2],pot_format)
            worksheet.write(f'{skip(col,2)}{row+3}',custo[3],rs_format)

            worksheet.write(f'{skip(col,5)}{row+3}',"Demanda",border)
            worksheet.write(f'{skip(col,6)}{row+3}',custo[3]/tarifas[2],pot_format)
            worksheet.write(f'{skip(col,7)}{row+3}',custo[3],rs_format)
            worksheet.write(f'{skip(col,5)}{row+4}',"Total",border)
            worksheet.write(f'{skip(col,6)}{row+4}',"-",border)
            worksheet.write_formula(f'{skip(col,6)}{row+4}',"=SUM(N2:N4)",rs_format)
        else:
            worksheet.write(f'{col}{row+3}',"Demanda FP",border)
            worksheet.write(f'{skip(col)}{row+3}',custo[3],pot_format)
            worksheet.write(f'{skip(col,2)}{row+3}',custo[3]*tarifas[2],rs_format)
            worksheet.write(f'{col}{row+4}',"Demanda P",border)
            worksheet.write(f'{skip(col)}{row+4}',custo[4],pot_format)
            worksheet.write(f'{skip(col,2)}{row+4}',custo[4]*tarifas[3],rs_format)
            
            worksheet.write(f'{skip(col,5)}{row+3}',"Demanda FP",border)
            worksheet.write(f'{skip(col,6)}{row+3}',custo[3],pot_format)
            worksheet.write(f'{skip(col,7)}{row+3}',custo[3]*tarifas[2],rs_format)
            worksheet.write(f'{skip(col,5)}{row+4}',"Demanda P",border)
            worksheet.write(f'{skip(col,6)}{row+4}',custo[4],pot_format)
            worksheet.write(f'{skip(col,7)}{row+4}',custo[4]*tarifas[3],rs_format)
            worksheet.write(f'{skip(col,5)}{row+5}',"Total",border)
            worksheet.write(f'{skip(col,6)}{row+5}',"-",border)
            worksheet.write_formula(f'{skip(col,7)}{row+5}',"=SUM(N2:N5)",rs_format)

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

def consumo_geral_valores_dm(worksheet,workbook,categoria,tarifas,custo,dias):
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
        worksheet.merge_range("H1:J1","Valores Diários",merge_format)
        worksheet.merge_range("M1:O1","Valores Mensais",merge_format)
    else:
        worksheet.merge_range("G1:I1","Valores Diários",merge_format)
        worksheet.merge_range("L1:N1","Valores Mensais",merge_format)

    border = workbook.add_format({'border':1})
    rs_format = workbook.add_format({'num_format':'R$ #,##0.00','border':1})
    pot_format = workbook.add_format({'num_format':'#,##0.00 "kW"','border':1})
    nrg_format = workbook.add_format({'num_format':'#,##0.00 "kWh"','border':1})
    
    if categoria == 'Convencional':
        worksheet.write('G2',"Consumo",border)
        worksheet.write('H2',custo[0]/tarifas[0],nrg_format)
        worksheet.write('I2',custo[0],rs_format)

        worksheet.write('L2',"Consumo",border)
        worksheet.write('M2',dias*custo[0]/tarifas[0],nrg_format)
        worksheet.write('N2',dias*custo[0],rs_format)

    if categoria == 'Branca':
        worksheet.write('H2',"Consumo FP",border)
        worksheet.write('I2',custo[0]/tarifas[0],nrg_format)
        worksheet.write('J2',custo[0],rs_format)
        worksheet.write('H3',"Consumo I",border)
        worksheet.write('I3',custo[2]/tarifas[2],nrg_format)
        worksheet.write('J3',custo[2],rs_format)
        worksheet.write('H4',"Consumo P",border)
        worksheet.write('I4',custo[1]/tarifas[1],nrg_format)
        worksheet.write('J4',custo[1],rs_format)
        # worksheet.write('H5',"Total",border)
        # worksheet.write('I5',"-",border)
        # worksheet.write_formula('J5',"=SUM(J2:J4)",rs_format)


        worksheet.write('M2',"Consumo FP",border)
        worksheet.write('N2',dias*custo[0]/tarifas[0],nrg_format)
        worksheet.write('O2',dias*custo[0],rs_format)
        worksheet.write('M3',"Consumo I",border)
        worksheet.write('N3',dias*custo[2]/tarifas[2],nrg_format)
        worksheet.write('O3',dias*custo[2],rs_format)
        worksheet.write('M4',"Consumo P",border)
        worksheet.write('N4',dias*custo[1]/tarifas[1],nrg_format)
        worksheet.write('O4',dias*custo[1],rs_format)
        worksheet.write('M5',"Total",border)
        worksheet.write('N5',"-",border)
        worksheet.write_formula('O5',"=SUM(O2:O4)",rs_format)

    elif categoria == 'Verde' or categoria == 'Azul':
        worksheet.write('G2',"Consumo FP",border)
        worksheet.write('H2',custo[0]/tarifas[0],nrg_format)
        worksheet.write('I2',custo[0],rs_format)
        worksheet.write('G3',"Consumo P",border)
        worksheet.write('H3',custo[1]/tarifas[1],nrg_format)
        worksheet.write('I3',custo[1],rs_format)

        worksheet.write('L2',"Consumo FP",border)
        worksheet.write('M2',dias*custo[0]/tarifas[0],nrg_format)
        worksheet.write('N2',dias*custo[0],rs_format)
        worksheet.write('L3',"Consumo P",border)
        worksheet.write('M3',dias*custo[1]/tarifas[1],nrg_format)
        worksheet.write('N3',dias*custo[1],rs_format)
        if categoria == 'Verde':
            worksheet.write('G4',"Demanda",border)
            worksheet.write('H4',custo[3]/tarifas[2],pot_format)
            worksheet.write('I4',custo[3],rs_format)
            # worksheet.write('G5',"Total",border)
            # worksheet.write('H5',"-",border)
            # worksheet.write_formula('I5',"=SUM(H2:H4)",rs_format)

            worksheet.write('L4',"Demanda",border)
            worksheet.write('M4',custo[3]/tarifas[2],pot_format)
            worksheet.write('N4',custo[3],rs_format)
            worksheet.write('L5',"Total",border)
            worksheet.write('M5',"-",border)
            worksheet.write_formula('N5',"=SUM(N2:N4)",rs_format)
        else:
            worksheet.write('G4',"Demanda FP",border)
            worksheet.write('H4',custo[3],pot_format)
            worksheet.write('I4',custo[3]*tarifas[2],rs_format)
            worksheet.write('G5',"Demanda P",border)
            worksheet.write('H5',custo[4],pot_format)
            worksheet.write('I5',custo[4]*tarifas[3],rs_format)
            # worksheet.write('G6',"Total",border)
            # worksheet.write('H6',"-",border)
            # worksheet.write_formula('I6',"=SUM(H2:H5)",rs_format)
            
            worksheet.write('L4',"Demanda FP",border)
            worksheet.write('M4',custo[3],pot_format)
            worksheet.write('N4',custo[3]*tarifas[2],rs_format)
            worksheet.write('L5',"Demanda P",border)
            worksheet.write('M5',custo[4],pot_format)
            worksheet.write('N5',custo[4]*tarifas[3],rs_format)
            worksheet.write('L6',"Total",border)
            worksheet.write('M6',"-",border)
            worksheet.write_formula('N6',"=SUM(N2:N5)",rs_format)
import pandas as pd
import calculo_tarifas

def get_hora(tempo):
    h = float(tempo.split(":")[0])
    m = float(tempo.split(":")[1])
    hora = h + (m/60)
    return hora

def calc_intervalo(categoria,inicio,fim,values):
    intervalo = 0
    h_ponta = int(values['-h_ponta-'])
    if categoria == 'Convencional':
        intervalo = get_hora(fim)-get_hora(inicio)
    elif categoria == 'Verde' or categoria == 'Azul':
        intervalo = [0,0]
        if get_hora(inicio) <= h_ponta and get_hora(fim) <= h_ponta:
            intervalo[0] = get_hora(fim)-get_hora(inicio)
            intervalo[1] = 0
        elif get_hora(inicio) <= h_ponta and get_hora(fim) >= h_ponta:
            intervalo[0] = h_ponta - get_hora(inicio)
            if get_hora(fim) <= (h_ponta+3):
                intervalo[1] = get_hora(fim) - h_ponta
            else: 
                intervalo[1] = 3
                intervalo[0] += get_hora(fim) - (h_ponta+3)
        elif get_hora(inicio) >= h_ponta:
            if get_hora(fim)>=(h_ponta+3):
                intervalo[1] = (h_ponta+3)-get_hora(inicio)
                intervalo[0] = get_hora(fim) - (h_ponta+3)
            else:
                intervalo[1] = get_hora(fim)-get_hora(inicio)
                intervalo[0] = 0
    if categoria == 'Branca':
        intervalo = [0,0,0]
        if get_hora(fim) <= h_ponta-1:
            intervalo[0] = get_hora(fim)-get_hora(inicio)
            intervalo[1] = 0
            intervalo[2] = 0

        elif get_hora(fim) <=h_ponta:
            if get_hora(inicio)<=h_ponta-1:
                intervalo[0] = (h_ponta-1) - get_hora(inicio)
                intervalo[1] = get_hora(fim) - (h_ponta-1)
            else:
                intervalo[0] = 0
                intervalo[1] = get_hora(fim) - get_hora(inicio)
            intervalo[2] = 0

        elif get_hora(fim) <= h_ponta+3:
            if get_hora(inicio) <= h_ponta-1:
                intervalo[0] = (h_ponta-1) - get_hora(inicio)
                intervalo[1] = 1
                intervalo[2] = get_hora(fim) - h_ponta
            elif get_hora(inicio) <= h_ponta:
                intervalo[0] = 0
                intervalo[1] = h_ponta - get_hora(inicio)
                intervalo[2] = get_hora(fim) - h_ponta
            else:
                intervalo[0] = 0
                intervalo[1] = 0
                intervalo[2] = get_hora(fim) - get_hora(inicio)

        elif get_hora(fim)<= h_ponta+4:
            if get_hora(inicio) <= h_ponta-1:
                intervalo[0] = (h_ponta-1) - get_hora(inicio) 
                intervalo[1] = 1 + get_hora(fim) - (h_ponta+3)
                intervalo[2] = 3
            elif get_hora(inicio) <= h_ponta:
                intervalo[0] = 0
                intervalo[1] = h_ponta - get_hora(inicio) + get_hora(fim) - h_ponta+3
                intervalo[2] = 3
            elif get_hora(inicio) <= get_hora+3:
                intervalo[0] = 0
                intervalo[1] = get_hora(fim) - (h_ponta+3)
                intervalo[2] = (h_ponta+3) - get_hora(inicio)
            else:
                intervalo[0] = 0
                intervalo[1] = get_hora(fim) - get_hora(inicio)
                intervalo[2] = 0

        else:
            if get_hora(inicio) <= h_ponta-1:
                intervalo[0] = (h_ponta-1) - get_hora(inicio) + get_hora(fim) - h_ponta+3
                intervalo[1] = 2
                intervalo[2] = 3
            elif get_hora(inicio) <= h_ponta:
                intervalo[0] = get_hora(fim) - (h_ponta+4)
                intervalo[1] = h_ponta - get_hora(inicio) + 1
                intervalo[2] = 3
            elif get_hora(inicio) <= get_hora+3:
                intervalo[0] = get_hora(fim) - (h_ponta+4)
                intervalo[1] = 1
                intervalo[2] = (h_ponta+3) - get_hora(inicio)
            elif get_hora(inicio) <= get_hora+4:
                intervalo[0] = get_hora(fim) - (h_ponta+4)
                intervalo[1] = (h_ponta+4) - get_hora(inicio)
                intervalo[2] = 0
            else:
                intervalo[0] = get_hora(fim) - get_hora(inicio)
                intervalo[1] = 0
                intervalo[2] = 0
    return intervalo

def select_consumo(itens,categoria,values):  
    h_ponta = int(values['-h_ponta-'])
    if categoria == 'Convencional':
        consumo_dict = {'Horas':[],'Minutos':[],'Potência - kW':[]}
        for h in range(0,24):
            for m in range(0,60):
                i=0
                pot=0
                while i < len(itens['Equipamentos']):
                    if get_hora(f'{h}:{m}')>=get_hora(itens['Início'][i]) and get_hora(f'{h}:{m}')<get_hora(itens['Fim'][i]):
                        pot += float(itens['Potência'][i])
                    i+=1
                consumo_dict['Potência - kW'].append(pot)
                consumo_dict['Horas'].append(h)
                consumo_dict['Minutos'].append(m)
    
    elif categoria == 'Branca':
        consumo_dict = {'Horas':[],'Minutos':[],'Potência FP - kW':[],'Potência P - kW':[],'Potência I - kW':[]}
        for h in range(0,24):
            for m in range(0,60):
                i=0
                pot_fp = 0
                pot_p = 0
                pot_i = 0
                while i < len(itens['Equipamentos']):
                    if get_hora(f'{h}:{m}')>=get_hora(itens['Início'][i]) and get_hora(f'{h}:{m}')<get_hora(itens['Fim'][i]):  
                        if h==(h_ponta-1) or h==(h_ponta+3):
                            pot_i += float(itens['Potência'][i])
                        elif h<(h_ponta-1) or h>=(h_ponta+4):
                            pot_fp += float(itens['Potência'][i])
                        else:
                            pot_p += float(itens['Potência'][i])
                    i+=1
                consumo_dict['Potência FP - kW'].append(pot_fp)
                consumo_dict['Potência P - kW'].append(pot_p)
                consumo_dict['Potência I - kW'].append(pot_i)
                consumo_dict['Horas'].append(h)
                consumo_dict['Minutos'].append(m)         
    
    else:
        consumo_dict = {'Horas':[],'Minutos':[],'Potência FP - kW':[],'Potência P - kW':[]}
        for h in range(0,24):
            for m in range(0,60):
                i=0
                pot_fp = 0
                pot_p = 0
                pot_i = 0
                while i < len(itens['Equipamentos']):
                    if get_hora(f'{h}:{m}')>=get_hora(itens['Início'][i]) and get_hora(f'{h}:{m}')<get_hora(itens['Fim'][i]):
                        if h<h_ponta or h>=(h_ponta+3):
                            pot_fp += float(itens['Potência'][i])
                        else:
                            pot_p += float(itens['Potência'][i])
                    i+=1
                consumo_dict['Potência FP - kW'].append(pot_fp)
                consumo_dict['Potência P - kW'].append(pot_p)
                consumo_dict['Horas'].append(h)
                consumo_dict['Minutos'].append(m)    
    
    return consumo_dict
    

def criar_consumo(itens,writer,categoria,tarifas,values):
    consumo_dict = select_consumo(itens,categoria,values)

    df_consumo = pd.DataFrame(consumo_dict)
    df_consumo.to_excel(writer, sheet_name="Consumo geral", startrow=1, header=False, index=False)
    workbook = writer.book
    worksheet = writer.sheets["Consumo geral"]
    (max_row, max_col) = df_consumo.shape
    column_settings = [{"header": column} for column in df_consumo.columns]
    worksheet.add_table(0, 0, max_row, max_col - 1, {"columns": column_settings})
    worksheet.set_column(0, max_col - 1, 12)
    print(tarifas)
    
    custo = calculo_tarifas.select_tarifa(tarifas,categoria,consumo_dict)
    print(tarifas)
    i=0
    for t in tarifas:
        if isinstance(t,str):
            tarifas[i] = float(t.replace(",","."))
        i+=1
    rs_format = workbook.add_format({'num_format':'R$ #,##0.00'})
    pot_format = workbook.add_format({'num_format':'0.00 "kW"'})
    nrg_format = workbook.add_format({'num_format':'0.00 "kWh"'})
    if categoria == 'Convencional':
        worksheet.write('G1',"Consumo:")
        
        worksheet.write('H1',custo[0]/tarifas[0],nrg_format)
        worksheet.write('I1',custo[0],rs_format)
    if categoria == 'Branca':
        worksheet.write('G1',"Consumo FP:")
        worksheet.write('H1',custo[0]/tarifas[0],nrg_format)
        worksheet.write('I1',custo[0],rs_format)
        worksheet.write('G2',"Consumo I:")
        worksheet.write('H2',custo[2]/tarifas[2],nrg_format)
        worksheet.write('I2',custo[2],rs_format)
        worksheet.write('G3',"Consumo P:")
        worksheet.write('H3',custo[1]/tarifas[1],nrg_format)
        worksheet.write('I3',custo[1],rs_format)

    elif categoria == 'Verde' or categoria == 'Azul':
        worksheet.write('G1',"Consumo FP:")
        worksheet.write('H1',custo[0]/tarifas[0],nrg_format)
        worksheet.write('I1',custo[0],rs_format)
        worksheet.write('G2',"Consumo P:")
        worksheet.write('H2',custo[1]/tarifas[1],nrg_format)
        worksheet.write('I2',custo[1],rs_format)
        if categoria == 'Verde':
            worksheet.write('G4',"Demanda")
            worksheet.write('H4',custo[3]/tarifas[2],pot_format)
            worksheet.write('I4',custo[3],rs_format)
        else:
            worksheet.write('G4',"Demanda FP")
            worksheet.write('H4',custo[3]/tarifas[2],pot_format)
            worksheet.write('I4',custo[3],rs_format)
            worksheet.write('G5',"Demanda P")
            worksheet.write('H5',custo[4]/tarifas[3],pot_format)
            worksheet.write('I5',custo[4],rs_format)
        
    
    worksheet.autofit()
    valores_equipamentos(itens,writer,categoria,values)

def valores_equipamentos(itens,writer,categoria,values):
    if categoria == 'Verde' or categoria == 'Azul':
        equip_dict = {"Equipamento":[],
                    "Potência (kW)":[],
                    "Horas - ponta":[],
                    "Horas - fora ponta":[],
                    "Total de horas":[],
                    "Consumo - ponta":[],
                    "Consumo - fora ponta":[],
                    "Consumo total":[]}
    elif categoria == 'Branca':
        equip_dict = {"Equipamento":[],
                    "Potência (kW)":[],
                    "Horas - ponta":[],
                    "Horas - intermediário":[],
                    "Horas - fora ponta":[],
                    "Total de horas":[],
                    "Consumo - ponta":[],
                    "Consumo - intermediário":[],
                    "Consumo - fora ponta":[],
                    "Consumo total":[]}
    else:
        equip_dict = {"Equipamento":[],
                    "Potência (kW)":[],
                    "Horas":[],
                    "Consumo":[]}
    
    i=0
    for equip in itens['Equipamentos']:
        equip_dict['Equipamento'].append(equip)
        equip_dict['Potência (kW)'].append(itens['Potência'][i])
        if categoria == 'Convencional':
            equip_dict['Horas'].append(calc_intervalo(categoria,itens['Início'][i],itens['Fim'][i],values))
            equip_dict['Consumo'].append(itens['Potência'][i]*calc_intervalo(categoria,itens['Início'][i],itens['Fim'][i],values))
        elif categoria == 'Verde' or categoria == 'Azul':
            equip_dict['Horas - ponta'].append(calc_intervalo(categoria,itens['Início'][i],itens['Fim'][i],values)[1])
            equip_dict['Horas - fora ponta'].append(calc_intervalo(categoria,itens['Início'][i],itens['Fim'][i],values)[0])
            equip_dict['Total de horas'].append(equip_dict['Horas - ponta'][i]+equip_dict['Horas - fora ponta'][i])
            equip_dict['Consumo - ponta'].append(itens['Potência'][i]*calc_intervalo(categoria,itens['Início'][i],itens['Fim'][i],values)[1])
            equip_dict['Consumo - fora ponta'].append(itens['Potência'][i]*calc_intervalo(categoria,itens['Início'][i],itens['Fim'][i],values)[0])
            equip_dict['Consumo total'].append(equip_dict['Potência (kW)'][i]*equip_dict['Total de horas'][i])
        else:
            equip_dict['Horas - ponta'].append(calc_intervalo(categoria,itens['Início'][i],itens['Fim'][i],values)[2])
            equip_dict['Horas - intermediário'].append(calc_intervalo(categoria,itens['Início'][i],itens['Fim'][i],values)[1])
            equip_dict['Horas - fora ponta'].append(calc_intervalo(categoria,itens['Início'][i],itens['Fim'][i],values)[0])
            equip_dict['Total de horas'].append(equip_dict['Horas - ponta'][i]+equip_dict['Horas - fora ponta'][i]+equip_dict['Horas - intermediário'][i])
            equip_dict['Consumo - ponta'].append(itens['Potência'][i]*calc_intervalo(categoria,itens['Início'][i],itens['Fim'][i],values)[2])
            equip_dict['Consumo - fora ponta'].append(itens['Potência'][i]*calc_intervalo(categoria,itens['Início'][i],itens['Fim'][i],values)[0])
            equip_dict['Consumo - intermediário'].append(itens['Potência'][i]*calc_intervalo(categoria,itens['Início'][i],itens['Fim'][i],values)[1])
            equip_dict['Consumo total'].append(equip_dict['Potência (kW)'][i]*equip_dict['Total de horas'][i])
        i+=1
    print(equip_dict)
    df_equip = pd.DataFrame(equip_dict)
    df_equip.to_excel(writer, sheet_name="Consumo por equipamento", startrow=1, header=False, index=False)
    workbook = writer.book
    worksheet = writer.sheets["Consumo por equipamento"]
    (max_row, max_col) = df_equip.shape
    column_settings = [{"header": column} for column in df_equip.columns]
    worksheet.add_table(0, 0, max_row, max_col - 1, {"columns": column_settings})
    worksheet.set_column(0, max_col - 1, 12)
    worksheet.autofit()
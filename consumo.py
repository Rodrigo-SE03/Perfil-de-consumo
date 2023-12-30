import pandas as pd
import calculo_tarifas,estilos

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
        consumo_dict = {'Horas':[],'Minutos':[],'Instante':[],'Potência - kW':[]}
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
                consumo_dict['Instante'].append(0) 
    
    elif categoria == 'Branca':
        consumo_dict = {'Horas':[],'Minutos':[],'Instante':[],'Potência FP - kW':[],'Potência P - kW':[],'Potência I - kW':[]}
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
                consumo_dict['Instante'].append(0)         
    
    else:
        consumo_dict = {'Horas':[],'Minutos':[],'Instante':[],'Potência FP - kW':[],'Potência P - kW':[]}
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
                consumo_dict['Instante'].append(0) 
    
    return consumo_dict
    

def criar_consumo(itens,writer,categoria,tarifas,values):
    consumo_dict = select_consumo(itens,categoria,values)

    custo = calculo_tarifas.select_tarifa(tarifas,categoria,consumo_dict)
    print(tarifas)

    df_consumo = pd.DataFrame(consumo_dict)
    df_consumo.to_excel(writer, sheet_name="Consumo geral", startrow=1, header=False, index=False)
    workbook = writer.book
    worksheet = writer.sheets["Consumo geral"]
    (max_row, max_col) = df_consumo.shape
    column_settings = [{"header": column} for column in df_consumo.columns]
    worksheet.add_table(0, 0, max_row, max_col - 1, {"columns": column_settings})
    worksheet.set_column(0, max_col - 1, 12)
    print(tarifas)
    

    i=0
    for t in tarifas:
        if isinstance(t,str):
            tarifas[i] = float(t.replace(",","."))
        i+=1
    estilos.consumo_geral_valores_dm(worksheet,workbook,categoria,tarifas,custo,int(values['-dias-']))

    hora_format = workbook.add_format({'num_format': 'hh:mm:ss'})
    i=1
    while i <= len(consumo_dict['Horas']):
        worksheet.write_formula(f'C{i+1}', f'=DATE(YEAR(TODAY()), MONTH(TODAY()), DAY(TODAY())) + TIME(A{i+1}, B{i+1}, 0)',hora_format)
        i+=1
    
    worksheet.autofit()
    worksheet.set_column("A:B",None, None,{"hidden":True})
    criar_grafico(worksheet,workbook,categoria)
    valores_equipamentos(itens,writer,categoria,values)

def criar_grafico(worksheet,workbook,categoria): #CRIAR OS GRÁFICOS DIFERENTES PARA TARIFA BRANCA E CONVENCIONAL
    chart = workbook.add_chart({'type':'column'})
    if categoria == 'Convencional':
        chart.add_series({'categories':"='Consumo geral'!$C$2:$C$1441",'name': "Potência",'values':"='Consumo geral'!$D$2:$D$1441"})
    elif categoria == 'Branca':
        chart.add_series({'categories':"='Consumo geral'!$C$2:$C$1441",'name': "Potência - Fora Ponta",'values':"='Consumo geral'!$D$2:$D$1441"})
        chart.add_series({'name':"Potência - Ponta",'values':"='Consumo geral'!$E$2:$E$1441"})
        chart.add_series({'name':"Potência - Intermediário",'values':"='Consumo geral'!$F$2:$F$1441"})
    else:
        chart.add_series({'categories':"='Consumo geral'!$C$2:$C$1441",'name': "Potência - Fora Ponta",'values':"='Consumo geral'!$D$2:$D$1441"})
        chart.add_series({'name':"Potência - Ponta",'values':"='Consumo geral'!$E$2:$E$1441"})
    
    chart.set_x_axis(
    {
        "interval_unit": 60,
        "num_format": "h"
    })
    chart.set_size({'width': 860, 'height': 300})
    chart.set_legend({'position': 'bottom'})
    if categoria == "Branca":
        worksheet.insert_chart('H9', chart)
    else:
        worksheet.insert_chart('G9', chart)

def valores_equipamentos(itens,writer,categoria,values):
    if categoria == 'Verde' or categoria == 'Azul':
        equip_dict = {"Equipamento":[],
                    "Potência (kW)":[],
                    "H - Ponta":[],
                    "H - Fora Ponta":[],
                    "Total - H":[],
                    "C - Ponta":[],
                    "C - Fora Ponta":[],
                    "Total - C":[]}
    elif categoria == 'Branca':
        equip_dict = {"Equipamento":[],
                    "Potência (kW)":[],
                    "H - Ponta":[],
                    "H - Intermediário":[],
                    "H - Fora Ponta":[],
                    "Total - H":[],
                    "C - Ponta":[],
                    "C - Intermediário":[],
                    "C - Fora Ponta":[],
                    "Total - C":[]}
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
            equip_dict['H - Ponta'].append(calc_intervalo(categoria,itens['Início'][i],itens['Fim'][i],values)[1])
            equip_dict['H - Fora Ponta'].append(calc_intervalo(categoria,itens['Início'][i],itens['Fim'][i],values)[0])
            equip_dict['Total - H'].append(equip_dict['H - Ponta'][i]+equip_dict['H - Fora Ponta'][i])
            equip_dict['C - Ponta'].append(itens['Potência'][i]*calc_intervalo(categoria,itens['Início'][i],itens['Fim'][i],values)[1])
            equip_dict['C - Fora Ponta'].append(itens['Potência'][i]*calc_intervalo(categoria,itens['Início'][i],itens['Fim'][i],values)[0])
            equip_dict['Total - C'].append(equip_dict['Potência (kW)'][i]*equip_dict['Total - H'][i])
        else:
            equip_dict['H - Ponta'].append(calc_intervalo(categoria,itens['Início'][i],itens['Fim'][i],values)[2])
            equip_dict['H - Intermediário'].append(calc_intervalo(categoria,itens['Início'][i],itens['Fim'][i],values)[1])
            equip_dict['H - Fora Ponta'].append(calc_intervalo(categoria,itens['Início'][i],itens['Fim'][i],values)[0])
            equip_dict['Total - H'].append(equip_dict['H - Ponta'][i]+equip_dict['H - Fora Ponta'][i]+equip_dict['H - Intermediário'][i])
            equip_dict['C - Ponta'].append(itens['Potência'][i]*calc_intervalo(categoria,itens['Início'][i],itens['Fim'][i],values)[2])
            equip_dict['C - Fora Ponta'].append(itens['Potência'][i]*calc_intervalo(categoria,itens['Início'][i],itens['Fim'][i],values)[0])
            equip_dict['C - Intermediário'].append(itens['Potência'][i]*calc_intervalo(categoria,itens['Início'][i],itens['Fim'][i],values)[1])
            equip_dict['Total - C'].append(equip_dict['Potência (kW)'][i]*equip_dict['Total - H'][i])
        i+=1
    print(equip_dict)
    df_equip = pd.DataFrame(equip_dict)
    df_equip.to_excel(writer, sheet_name="Consumo por equipamento", startrow=2, header=False, index=False)
    workbook = writer.book
    worksheet = writer.sheets["Consumo por equipamento"]
    (max_row, max_col) = df_equip.shape
    column_settings = [{"header": column} for column in df_equip.columns]
    worksheet.add_table(1, 0, max_row+1, max_col-1, {"columns": column_settings})
    worksheet.set_column(0, max_col - 1, 12)
    estilos.consumo_equip_style(worksheet,workbook,categoria)
    worksheet.autofit()
import pandas as pd
import calculo_tarifas

def get_hora(tempo):
    h = float(tempo.split(":")[0])
    m = float(tempo.split(":")[1])
    hora = h + (m/60)
    return hora

def select_consumo(itens,categoria,values):  #TENHO QUE MUDAR AQUI PRA ARRUMAR A QUESTÃO DE DEFINIR OS HORÁRIOS DE PONTA
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
                    if get_hora(f'{h}:{m}')>=get_hora(itens['Início'][i]) and get_hora(f'{h}:{m}')<get_hora(itens['Fim'][i]):   #MUDEI AQUI PRA FICAR IGUAL O EXEMPLO DO SLIDE
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
    df_consumo.to_excel(writer, sheet_name="Consumo", startrow=1, header=False, index=False)
    workbook = writer.book
    worksheet = writer.sheets["Consumo"]
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
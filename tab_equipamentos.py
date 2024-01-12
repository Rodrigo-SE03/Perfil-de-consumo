import PySimpleGUI as psg

def criarTabEquip(lista):
    horas = []
    i=0
    while i<24:
        horas.append(i)
        i+=1

    l_ponta = psg.Text('Início do horário de ponta')
    ponta = psg.Input(key='-h_ponta-',size=2,default_text='18')

    l_dias = psg.Text('                 Nº de dias de funcionamento')
    dias = psg.Input(key='-dias-',size=2,default_text='22')

    l1 = psg.Text('Nome do equipamento')
    equip = psg.Input(key='-equip-',size = 20)
    col1 = [[l1],[equip]]


    l2 = psg.Text('Potência (kW)')
    pot = psg.Input(key='-pot-',size=10)
    col2 = [[l2],[pot]]


    l3 = psg.Text('Horário de funcionamento')
    l31 = psg.Text('Início')
    h_inicio = psg.Input(key='-inicio-',size=6,default_text='00:00')
    l32 = psg.Text('Final')
    h_final = psg.Input(key='-final-',size=6,default_text='00:00')
    col3 =[[l3],[l31,h_inicio,l32,h_final]]

    l5 = psg.Text('Fator de Potência')
    fp = psg.Input(key='-fp-',default_text='1',size=4)
    fp_tipo = psg.Combo(['Indutivo','Capacitivo'],default_value="Indutivo",key='-fp_tipo-',size=8)
    col5 = [[l5],[fp,fp_tipo]]
    l6 = psg.Text('Quantidade')
    qtd = psg.Input(key='-qtd-',default_text='1',size=4)
    col6 = [[l6],[qtd]]
    

    b_add = psg.Button("Adicionar",key='-add-')
    sup_col = [[l_ponta,ponta,psg.Text('h'),l_dias,dias],[psg.Column(col1),psg.Column(col2),psg.Column(col3)],[psg.Column(col5),psg.Column(col6)],[b_add]]


    l4 = psg.Text('Lista de equipamentos')
    b_remove = psg.Button("Remover",key='-remove-')
    col4 = [[l4],[lista],[b_remove]]


    return [[psg.Column(sup_col),psg.Column(col4)]]

def saveEquip(event,values,itens,itens_concat,window):
    if values['-equip-'] == '' or values['-pot-'] == '' or values['-inicio-'] == '' or values['-final-'] == '' :
        psg.popup_auto_close("Preencha todos os campos para adicionar um item")
    elif values['-equip-'].rfind('-') != -1:
        psg.popup_auto_close("O nome do equipamento não pode conter o caractere -")
    elif values['-equip-'] not in itens['Equipamentos']:
        itens['Equipamentos'].append(values['-equip-'])
        itens['Potência'].append(values['-pot-'])
        itens['Fator de Potência'].append(values['-fp-'])
        itens['Tipo - FP'].append(values['-fp_tipo-'])
        itens['Início'].append(values['-inicio-'])
        itens['Fim'].append(values['-final-'])
        itens['Quantidade'].append(values['-qtd-'])
        i=0
        while i < len(itens['Equipamentos']):
            itens_concat.add(f"{itens['Equipamentos'][i]}- {itens['Potência'][i]} kW  FP: {itens['Fator de Potência'][i]}  {itens['Início'][i]} - {itens['Fim'][i]}")
            i+=1
        window['-list-'].update(itens_concat)

def updateEquip(itens,itens_concat,window,edit=False):
    i=0
    if edit == False:
        while i < len(itens['Equipamentos']):
            itens_concat.add(f"{itens['Equipamentos'][i]}- {itens['Potência'][i]} kW  FP: {itens['Fator de Potência'][i]}  {itens['Início'][i]} - {itens['Fim'][i]}")
            i+=1
    else:
        itens_concat.clear()
        while i < len(itens['Equipamentos']):
            itens_concat.add(f"{itens['Equipamentos'][i]}- {itens['Potência'][i]} kW  FP: {itens['Fator de Potência'][i]}  {itens['Início'][i]} - {itens['Fim'][i]}")
            i+=1
    window['-list-'].update(itens_concat)
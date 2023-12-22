import PySimpleGUI as psg
import pandas as pd
import equipamentos, consumo, tab_equipamentos, tab_tarifas
psg.set_options(font=("Arial Bold",14))
categoria = ''
itens = {'Equipamentos':[],
         'Potência':[],
         'Início':[],
         'Fim':[]
         }
itens_concat = set()

lista = psg.Listbox(itens_concat,size=(40,10),key='-list-')

tab1 = tab_equipamentos.criarTabEquip(lista=lista)

tab2 = tab_tarifas.criarTabTarifas()

b_save = psg.Button("Salvar",key='-save-')
b_load = psg.Button("Carregar",key='-load-')
layout = [[psg.TabGroup([[psg.Tab('Equipamentos',tab1)],[psg.Tab('Tarifas',tab2)]])],[b_save,b_load]]

window = psg.Window('Perfil de Consumo', layout)

def remove_item(val,itens):
    id = itens['Equipamentos'].index(val)
    itens['Equipamentos'].pop(id)
    itens['Potência'].pop(id)
    itens['Início'].pop(id)
    itens['Fim'].pop(id)
    return itens

while True:    
    event, values = window.read()  
    if event == '-add-':
        if values['-equip-'] in itens['Equipamentos']:
            print(itens)
            itens = remove_item(values['-equip-'],itens = itens)
            tab_equipamentos.updateEquip(itens=itens,itens_concat=itens_concat,window=window,edit=True)
        tab_equipamentos.saveEquip(event=event,values=values,itens=itens,itens_concat=itens_concat,window=window)

    if event == '-remove-':
        if len(lista.get()) == 0:
            psg.popup_auto_close("Selecione um item para remover")
        else:
            val = lista.get()[0]
            itens = remove_item(val.split('-')[0],itens = itens)
            print(itens_concat)
            itens_concat.remove(val)
            window['-list-'].update(itens_concat)

    if event == '-save-':
        if categoria == '':
            psg.popup_auto_close("Preencha os dados da tarifa")
        elif len(itens['Equipamentos']) == 0:
            psg.popup_auto_close("Insira no mínimo um equipamento")
        elif values['-h_ponta-'] == '' or values['-h_ponta-'] == None:
            psg.popup_auto_close("Horário de ponta inválido")
        else:
            xl_name = psg.popup_get_text('Digite o nome da nova planilha',title="Nome do arquivo")
            if xl_name == None or xl_name == '':
                continue
            xl_name = xl_name+".xlsx"
            writer = pd.ExcelWriter(xl_name,engine="xlsxwriter")
            equipamentos.criar_equip(itens,writer)
            consumo.criar_consumo(itens,writer,categoria,tarifas,values)
            writer.close()
    
    if event == '-load-':
        load_file = psg.popup_get_file('Selecione a planilha que deseja carregar',  title="Carregar arquivo")
        if load_file == None or load_file == '':
            continue
        try:
            load_df = pd.read_excel(load_file, sheet_name='Equipamentos')
        except: 
            psg.popup_auto_close("Arquivo inválido")
            continue
        itens = load_df.to_dict('list')
        itens_concat.clear()
        tab_equipamentos.updateEquip(itens=itens,itens_concat=itens_concat,window=window)

    if event == '-conv-' or event == '-branca-' or event == '-verde-' or event == '-azul-':
        tab_tarifas.change_tab(event=event,window=window)
    
    if event == '-reg-':
        tarifas_info = tab_tarifas.registrar(window=window,values=values)
        categoria = tarifas_info[0]
        tarifas = tarifas_info[1]
        print(categoria,tarifas)

    if event == psg.WIN_CLOSED: 
        break  



#Arrumar a questão da demanda nas tarifas verde e azul - cobrar de acordo com a demanda contratada
#Ultrapassagem

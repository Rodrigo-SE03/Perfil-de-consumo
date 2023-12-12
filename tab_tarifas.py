import PySimpleGUI as psg

def registrar(window,values):
    tarifas = []
    categoria = ''
    if window['-conv_col-'].visible == True:
        if values['-tarifa_conv-'] == '':
            psg.popup_auto_close("Preencha todos os campos")
        else:
            categoria = 'Convencional'
            tarifas.append(values['-tarifa_conv-'])

    elif window['-branca_col-'].visible == True:
        if values['-tarifaFP_branca-'] == '' or values['-tarifaP_branca-'] == '' or values['-tarifaI_branca-'] == '':
            psg.popup_auto_close("Preencha todos os campos")
        else:
            categoria = 'Branca'
            tarifas.append(values['-tarifaFP_branca-'])
            tarifas.append(values['-tarifaP_branca-'])
            tarifas.append(values['-tarifaI_branca-'])

    elif window['-verde_col-'].visible == True:
        if values['-tarifaFP_verde-'] == '' or values['-tarifaP_verde-'] == '' or values['-tarifaDEM_verde-'] == '':
            psg.popup_auto_close("Preencha todos os campos")
        else:
            categoria = 'Verde'
            tarifas.append(values['-tarifaFP_verde-'])
            tarifas.append(values['-tarifaP_verde-'])
            tarifas.append(values['-tarifaDEM_verde-'])

    elif window['-azul_col-'].visible == True:
        if values['-tarifaFP_azul-'] == '' or values['-tarifaP_azul-'] == '' or values['-tarifaDEMFP_azul-'] == '' or values['-tarifaDEMP_azul-'] == '':
            psg.popup_auto_close("Preencha todos os campos")
        else:
            categoria = 'Azul'
            tarifas.append(values['-tarifaFP_azul-'])
            tarifas.append(values['-tarifaP_azul-'])
            tarifas.append(values['-tarifaDEMFP_azul-'])
            tarifas.append(values['-tarifaDEMP_azul-'])

    return [categoria,tarifas]

def change_tab(event,window):
    if event == '-conv-':
        window['-verde_col-'].update(visible=False)
        window['-branca_col-'].update(visible=False)
        window['-azul_col-'].update(visible=False)
        window['-conv_col-'].update(visible=True)

    if event == '-branca-':
        window['-verde_col-'].update(visible=False)
        window['-branca_col-'].update(visible=True)
        window['-azul_col-'].update(visible=False)
        window['-conv_col-'].update(visible=False)

    if event == '-verde-':
        window['-verde_col-'].update(visible=True)
        window['-branca_col-'].update(visible=False)
        window['-azul_col-'].update(visible=False)
        window['-conv_col-'].update(visible=False)
    
    if event == '-azul-':
        window['-verde_col-'].update(visible=False)
        window['-branca_col-'].update(visible=False)
        window['-azul_col-'].update(visible=True)
        window['-conv_col-'].update(visible=False)

def criarTabTarifas():
    r1 = psg.Radio("Grupo B - Convencional","categoria",key = '-conv-',default=False,enable_events=True)
    r2 = psg.Radio("Grupo B - Branca","categoria",key = '-branca-',default=False,enable_events=True)
    r3 = psg.Radio("Grupo A - Verde","categoria",key = '-verde-',default=True,enable_events=True)
    r4 = psg.Radio("Grupo A - Azul","categoria",key = '-azul-',default=False,enable_events=True)

    l1_conv = psg.Text('Tarifa de consumo')
    tarifa_fp_conv = psg.Input(key='-tarifa_conv-',size = 10)
    col_conv = [[l1_conv,tarifa_fp_conv,psg.Text('R$/kWh')]]

    l1_branca = psg.Text('Tarifa de consumo Fora Ponta   ')
    tarifa_fp_branca = psg.Input(key='-tarifaFP_branca-',size = 10)
    l2_branca = psg.Text('Tarifa de consumo na Ponta       ')
    tarifa_p_branca = psg.Input(key='-tarifaP_branca-',size = 10)
    l3_branca = psg.Text('Tarifa de consumo Intermediária')
    tarifa_i_branca = psg.Input(key='-tarifaI_branca-',size = 10)
    col_branca = [[l1_branca,tarifa_fp_branca,psg.Text('R$/kWh')],[l2_branca,tarifa_p_branca,psg.Text('R$/kWh')],[l3_branca,tarifa_i_branca,psg.Text('R$/kWh')]]

    l1_verde = psg.Text('Tarifa de consumo Fora Ponta')
    tarifa_fp_verde = psg.Input(key='-tarifaFP_verde-',size = 10)
    l2_verde = psg.Text('Tarifa de consumo na Ponta   ')
    tarifa_p_verde = psg.Input(key='-tarifaP_verde-',size = 10)
    l3_verde = psg.Text('Tarifa de demanda')
    tarifa_dem_verde = psg.Input(key='-tarifaDEM_verde-',size = 10)
    l4_verde = psg.Text('Valor de demanda contratada')
    demC_verde = psg.Input(key='-tarifaDEMC_verde-',size = 10)
    col_verde = [[l1_verde,tarifa_fp_verde,psg.Text('R$/kWh'),l4_verde,demC_verde,psg.Text('kW')],[l2_verde,tarifa_p_verde,psg.Text('R$/kWh')],[l3_verde,tarifa_dem_verde,psg.Text('R$/kW')]]

    l1_azul = psg.Text('Tarifa de consumo Fora Ponta')
    tarifa_fp_azul = psg.Input(key='-tarifaFP_azul-',size = 10)
    l2_azul = psg.Text('Tarifa de consumo na Ponta   ')
    tarifa_p_azul = psg.Input(key='-tarifaP_azul-',size = 10)
    l3_azul = psg.Text('Tarifa de demanda Fora Ponta')
    tarifa_demFP_azul = psg.Input(key='-tarifaDEMFP_azul-',size = 10)
    l4_azul = psg.Text('Tarifa de demanda na Ponta   ')
    tarifa_demP_azul = psg.Input(key='-tarifaDEMP_azul-',size = 10)
    l5_azul = psg.Text('Valor de demanda contratada Fora Ponta')
    demCFP_azul = psg.Input(key='-tarifaDEMCFP_azul-',size = 10)
    l6_azul = psg.Text('Valor de demanda contratada na Ponta    ')
    demCP_azul = psg.Input(key='-tarifaDEMCP_azul-',size = 10)
    col_azul= [[l1_azul,tarifa_fp_azul,psg.Text('R$/kWh'),l5_azul,demCFP_azul,psg.Text('kW')],[l2_azul,tarifa_p_azul,psg.Text('R$/kWh'),l6_azul,demCP_azul,psg.Text('kW')],[l3_azul,tarifa_demFP_azul,psg.Text('R$/kW')],[l4_azul,tarifa_demP_azul,psg.Text('R$/kW')]]

    b_registrar = psg.Button("Registrar",key='-reg-')

    return [[r1,r2,r3,r4],[psg.Column(col_conv,visible=False,key='-conv_col-'),psg.Column(col_branca,visible=False,key='-branca_col-'),
                           psg.Column(col_verde,visible=True,key='-verde_col-'),psg.Column(col_azul,visible=False,key='-azul_col-')],[b_registrar]]
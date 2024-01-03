import openpyxl

def extrair_planilha(window):
    wb = openpyxl.load_workbook('Tarifas.xlsx')
    sheet = wb.active

    window['-tarifa_conv-'].update(sheet['D4'].value)

    window['-tarifaFP_branca-'].update(sheet['D8'].value)
    window['-tarifaI_branca-'].update(sheet['D9'].value)
    window['-tarifaP_branca-'].update(sheet['D10'].value)

    window['-tarifaFP_verde-'].update(sheet['D14'].value)
    window['-tarifaP_verde-'].update(sheet['D15'].value)
    window['-tarifaDEM_verde-'].update(sheet['D16'].value)
    
    window['-tarifaFP_azul-'].update(sheet['D20'].value)
    window['-tarifaP_azul-'].update(sheet['D21'].value)
    window['-tarifaDEMFP_azul-'].update(sheet['D22'].value)
    window['-tarifaDEMP_azul-'].update(sheet['D23'].value)

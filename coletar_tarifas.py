import openpyxl
from bs4 import BeautifulSoup
import requests
import PySimpleGUI as psg
from time import sleep

def extrair_planilha(window,planilha):
    wb = openpyxl.load_workbook(planilha)
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

def extrair_site(window):
    a_soup = BeautifulSoup(requests.get("https://go.equatorialenergia.com.br/valor-de-tarifas-e-servicos/#tarifas-grupo-a").content,"html.parser")
    b_soup = BeautifulSoup(requests.get("https://go.equatorialenergia.com.br/valor-de-tarifas-e-servicos/#residencial-normal").content,"html.parser")

    a_tarifas = []
    b_tarifas = []

    a_table = a_soup.find('table',width="577")
    try:
        a_table_body = a_table.find('tbody')
    except:
        sleep(1.5)
        psg.popup_auto_close("Funcionalidade indisponível no momento")
        return
    
    b_table = b_soup.find('table',width="900")
    b_table_body = b_table.find('tbody')

    rows = a_table_body.find_all('tr')
    for row in rows:
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        a_tarifas.append([ele for ele in cols if ele]) 

    rows = b_table_body.find_all('tr')
    for row in rows:
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        b_tarifas.append([ele for ele in cols if ele]) 
    
    print(b_tarifas)
    print(len(b_tarifas))

    window['-tarifa_conv-'].update(b_tarifas[2][1])

    window['-tarifaFP_branca-'].update(b_tarifas[5][3])
    window['-tarifaI_branca-'].update(b_tarifas[5][2])
    window['-tarifaP_branca-'].update(b_tarifas[5][1])
    
    window['-tarifaFP_verde-'].update(a_tarifas[17][2])
    window['-tarifaP_verde-'].update(a_tarifas[16][2])
    window['-tarifaDEM_verde-'].update(a_tarifas[15][2])
    
    window['-tarifaFP_azul-'].update(a_tarifas[12][2])
    window['-tarifaP_azul-'].update(a_tarifas[11][2])
    window['-tarifaDEMFP_azul-'].update(a_tarifas[10][2])
    window['-tarifaDEMP_azul-'].update(a_tarifas[9][2])
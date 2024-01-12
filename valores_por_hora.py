import math
import estilos
def tabela_por_hora(valores,categoria,writer,tarifas,dias):
    if categoria == "Verde":
        consumo_fp = valores[5]
        consumo_p = valores[6]
        demanda = valores[7]
        energia_reativa_fp = valores[8]
        energia_reativa_p = valores[9]
        indutivo = []
        capacitivo = []
        dem_kw = []
        rs = []

        t_cfp = tarifas[0]
        t_cp = tarifas[1]
        t_dem = tarifas[2]
        max_dem = valores[4]/t_dem

        demr = 0
        maior = 0
        cr = 0
        i=0
        while i<len(consumo_fp):
            consumo = consumo_fp[i] if consumo_fp[i] != 0 else consumo_p[i]
            reativo = energia_reativa_fp[i] if energia_reativa_fp[i] != 0 else energia_reativa_p[i]
            
            if consumo == 0: 
                indutivo.append(0)
                capacitivo.append(0)
                dem_kw.append(0)
                cr=0
            elif reativo<0: 
                capacitivo.append(consumo/math.sqrt(pow(consumo,2)+pow(reativo,2)))
                indutivo.append(0)
                dem_kw.append(demanda[i]*(0.92/capacitivo[i]))
            elif reativo>0:
                indutivo.append(consumo/math.sqrt(pow(consumo,2)+pow(reativo,2)))
                capacitivo.append(0)
                dem_kw.append(demanda[i]*(0.92/indutivo[i]))
            else:
                indutivo.append(1)
                capacitivo.append(1)
                dem_kw.append(demanda[i]*(0.92/indutivo[i]))

            if dem_kw[i] > maior:
                maior = dem_kw[i]
                if indutivo[i] != 0 or capacitivo[i] != 0:
                    if indutivo[i] >= 0.92 or capacitivo[i] >= 0.92: demr = 0
                    elif capacitivo[i] !=0 and capacitivo[i] < 0.92 and i>=6 and i<=17: demr = 0
                    else: demr = (maior-max_dem)*t_dem

            if indutivo[i] != 0 or capacitivo[i] != 0:
                if indutivo[i] >= 0.92 or capacitivo[i] >= 0.92:
                    cr=0
                elif capacitivo[i] !=0 and capacitivo[i] < 0.92 and i>=6 and i<=17: cr = 0
                else:
                    cr = math.fabs(consumo) * ((0.92/(indutivo[i] if indutivo[i] != 0 else capacitivo[i]))-1) * (t_cfp if consumo_fp[i] != 0 else t_cp)
            rs.append(cr)
            i+=1

        tabela_dict = {"Período":["0-1","1-2","2-3","3-4","4-5","5-6","6-7","7-8","8-9","9-10","10-11","11-12","12-13","13-14","14-15","15-16","16-17","17-18","18-19","19-20","20-21","21-22","22-23","23-24"],
                       "kW": demanda,
                       "Fora Ponta - kWh": consumo_fp,
                       "Ponta - kWh": consumo_p,
                       "Fora Ponta - kVArh": energia_reativa_fp,
                       "Ponta - kVArh": energia_reativa_p,
                       "Indutivo": indutivo,
                       "Capacitivo": capacitivo,
                       " kW":dem_kw,
                       "R$":rs
                       }
        consumo_mes = dias*sum(tabela_dict['R$'])
        estilos.tabela_reativos(tabela_dict,writer,demr,categoria,consumo_mes)
    else:
        consumo_fp = valores[6]
        consumo_p = valores[7]
        demanda_fp = valores[8]
        demanda_p = valores[9]
        energia_reativa_fp = valores[10]
        energia_reativa_p = valores[11]
        indutivo = []
        capacitivo = []
        dem_kw = []
        rs = []

        t_cfp = tarifas[0]
        t_cp = tarifas[1]
        t_dem_fp = tarifas[2]
        t_dem_p = tarifas[3]
        max_dem_fp = valores[4]
        max_dem_p = valores[5]
        demr_fp = 0
        demr_p = 0
        maior_fp = 0
        maior_p = 0
        cr = 0
        i=0
        while i<len(consumo_fp):
            consumo = consumo_fp[i] if consumo_fp[i] != 0 else consumo_p[i]
            reativo = energia_reativa_fp[i] if energia_reativa_fp[i] != 0 else energia_reativa_p[i]
            
            if consumo == 0: 
                indutivo.append(0)
                capacitivo.append(0)
                dem_kw.append(0)
                cr=0
            elif reativo<0: 
                capacitivo.append(consumo/math.sqrt(pow(consumo,2)+pow(reativo,2)))
                indutivo.append(0)
                dem_kw.append((0.92/capacitivo[i])*(demanda_fp[i] if demanda_fp[i]!=0 else demanda_p[i]))
            elif reativo>0:
                indutivo.append(consumo/math.sqrt(pow(consumo,2)+pow(reativo,2)))
                capacitivo.append(0)
                dem_kw.append((0.92/indutivo[i])*(demanda_fp[i] if demanda_fp[i]!=0 else demanda_p[i]))
            else:
                indutivo.append(1)
                capacitivo.append(1)
                dem_kw.append((0.92/indutivo[i])*(demanda_fp[i] if demanda_fp[i]!=0 else demanda_p[i]))

            if dem_kw[i] > maior_fp and demanda_fp[i] != 0:
                maior_fp = dem_kw[i]
                if indutivo[i] != 0 or capacitivo[i] != 0:
                    if indutivo[i] >= 0.92 or capacitivo[i] >= 0.92: demr_fp = 0
                    elif capacitivo[i] !=0 and capacitivo[i] < 0.92 and i>=6 and i<=17: demr_fp = 0
                    else: demr_fp = (maior_fp-max_dem_fp)*t_dem_fp

            if dem_kw[i] > maior_p and demanda_p[i] != 0:
                maior_p = dem_kw[i]
                if indutivo[i] != 0 or capacitivo[i] != 0:
                    if indutivo[i] > 0.92 or capacitivo[i] > 0.92: demr_p = 0
                    elif capacitivo[i] !=0 and capacitivo[i] < 0.92 and i>=6 and i<=17: demr_p = 0
                    else: demr_p = (maior_p-max_dem_p)*t_dem_p

            if indutivo[i] != 0 or capacitivo[i] != 0:
                if indutivo[i] > 0.92 or capacitivo[i] > 0.92:
                    cr=0
                elif capacitivo[i] !=0 and capacitivo[i] < 0.92 and i>=6 and i<=17: cr = 0
                else:
                    cr = math.fabs(consumo) * ((0.92/(indutivo[i] if indutivo[i] != 0 else capacitivo[i]))-1) * (t_cfp if consumo_fp[i] != 0 else t_cp)
            rs.append(cr)
            i+=1

        tabela_dict = {"Período":["0-1","1-2","2-3","3-4","4-5","5-6","6-7","7-8","8-9","9-10","10-11","11-12","12-13","13-14","14-15","15-16","16-17","17-18","18-19","19-20","20-21","21-22","22-23","23-24"],
                       "Fora Ponta - kW": demanda_fp,
                       "Ponta - kW": demanda_p,
                       "Fora Ponta - kWh": consumo_fp,
                       "Ponta - kWh": consumo_p,
                       "Fora Ponta - kVArh": energia_reativa_fp,
                       "Ponta - kVArh": energia_reativa_p,
                       "Indutivo": indutivo,
                       "Capacitivo": capacitivo,
                       " kW":dem_kw,
                       "R$":rs
                       }
        demr = [demr_fp,demr_p]
        consumo_mes = dias*sum(tabela_dict['R$'])
        estilos.tabela_reativos(tabela_dict,writer,demr,categoria,consumo_mes)
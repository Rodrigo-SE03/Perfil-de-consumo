import consumo

def select_tarifa(tarifas,categoria,consumo_dict):
    if categoria == "Convencional":
        return calc_convencional(tarifas,consumo_dict)
    elif categoria == "Branca":
        return calc_branca(tarifas,consumo_dict)
    elif categoria == "Verde":
        return calc_verde(tarifas,consumo_dict)
    elif categoria == "Azul":
        return calc_azul(tarifas,consumo_dict)

def calc_convencional(tarifas,consumo_dict):
    t = float(tarifas[0].replace(",",".")) if isinstance(tarifas[0],str) else float(tarifas[0])
    print(t)
    custo=0

    for pot in consumo_dict["Potência - kW"]:
        custo += t*float(pot)/60
    
    valores = [custo,'-']
    return valores

def calc_branca(tarifas,consumo_dict):
    t_fp = float(tarifas[0].replace(",",".")) if isinstance(tarifas[0],str) else float(tarifas[0])
    t_p = float(tarifas[1].replace(",",".")) if isinstance(tarifas[1],str) else float(tarifas[1])
    t_i = float(tarifas[2].replace(",",".")) if isinstance(tarifas[2],str) else float(tarifas[2])
    custo_fp=0
    custo_p=0
    custo_i=0

    for pot in consumo_dict["Potência FP - kW"]:
        custo_fp += t_fp*float(pot)/60
    for pot in consumo_dict["Potência P - kW"]:
        custo_p += t_p*float(pot)/60
    for pot in consumo_dict["Potência I - kW"]:
        custo_i += t_i*float(pot)/60
    
    custo = custo_i+custo_fp+custo_p
    valores = [custo_fp,custo_p,custo_i,custo,'-']
    return valores

def calc_verde(tarifas,consumo_dict):
    tc_fp = float(tarifas[0].replace(",",".")) if isinstance(tarifas[0],str) else float(tarifas[0])
    tc_p = float(tarifas[1].replace(",",".")) if isinstance(tarifas[1],str) else float(tarifas[1])
    td = float(tarifas[2].replace(",",".")) if isinstance(tarifas[2],str) else float(tarifas[2])

    consumo_fp=0
    c_fp_h = []

    consumo_p=0
    c_p_h = []

    demanda=0
    d_h = []
    mdh = 0

    energia_reativa_fp = []
    energia_reativa_p = []

    max_dem = 0

    er_fp=0
    er_p=0
    ch_fp=0
    ch_p=0
    i=0
    cont = 0

    while cont < len(consumo_dict["Potência FP - kW"]):
        consumo_fp += float(consumo_dict["Potência FP - kW"][cont])/60
        ch_fp +=float(consumo_dict["Potência FP - kW"][cont])/60
        er_fp +=float(consumo_dict["Potência Reativa FP - kVAr"][cont])/60

        consumo_p += float(consumo_dict["Potência P - kW"][cont])/60
        ch_p +=float(consumo_dict["Potência P - kW"][cont])/60
        er_p +=float(consumo_dict["Potência Reativa P - kVAr"][cont])/60

        i+=1

        pot = consumo_dict["Potência FP - kW"][cont] if consumo_dict["Potência FP - kW"][cont]>consumo_dict["Potência P - kW"][cont] else consumo_dict["Potência P - kW"][cont]
        demanda = float(pot)
        if demanda>max_dem:
            max_dem=demanda
        if demanda>mdh:
            mdh = demanda
        if i == 60:
            d_h.append(mdh)
            mdh=0
        demanda = 0

        if i == 60:
            c_fp_h.append(ch_fp)
            c_p_h.append(ch_p)
            energia_reativa_fp.append(er_fp)
            energia_reativa_p.append(er_p)
            er_p=0
            er_fp=0
            ch_fp=0
            ch_p=0
            i=0
        cont+=1

    custo_fp = consumo_fp*tc_fp
    custo_p = consumo_p*tc_p
    custo_demanda = max_dem*td
    valores = [consumo_fp,consumo_p,custo_fp,custo_p,custo_demanda,c_fp_h,c_p_h,d_h,energia_reativa_fp,energia_reativa_p]
    return valores

def calc_azul(tarifas,consumo_dict):
    tc_fp = float(tarifas[0].replace(",",".")) if isinstance(tarifas[0],str) else float(tarifas[0])
    tc_p = float(tarifas[1].replace(",",".")) if isinstance(tarifas[1],str) else float(tarifas[1])

    consumo_fp=0
    c_fp_h = []

    consumo_p=0
    c_p_h = []

    energia_reativa_fp = []
    energia_reativa_p = []

    demanda_p=0
    d_p_h = []
    mdph = 0
    max_dem_p = 0

    demanda_fp=0
    d_fp_h = []
    mdfph = 0
    max_dem_fp = 0

    er_fp=0
    er_p=0
    ch_fp=0
    ch_p=0
    cont=0
    i=0

    while cont < len(consumo_dict["Potência FP - kW"]):
        consumo_fp += float(consumo_dict["Potência FP - kW"][cont])/60
        ch_fp +=float(consumo_dict["Potência FP - kW"][cont])/60
        er_fp +=float(consumo_dict["Potência Reativa FP - kVAr"][cont])/60

        consumo_p += float(consumo_dict["Potência P - kW"][cont])/60
        ch_p +=float(consumo_dict["Potência P - kW"][cont])/60
        er_p +=float(consumo_dict["Potência Reativa P - kVAr"][cont])/60

        i+=1

        demanda_fp = float(consumo_dict["Potência FP - kW"][cont])
        if demanda_fp>max_dem_fp:
            max_dem_fp=demanda_fp
        if demanda_fp>mdfph:
            mdfph = demanda_fp
        if i == 60:
            d_fp_h.append(mdfph)
            mdfph=0
        demanda_fp = 0

        demanda_p = float(consumo_dict["Potência P - kW"][cont])
        if demanda_p>max_dem_p:
            max_dem_p=demanda_p
        if demanda_p>mdph:
            mdph = demanda_p
        if i == 60:
            d_p_h.append(mdph)
            mdph=0
        demanda_p = 0
        
        if i == 60:
            c_fp_h.append(ch_fp)
            c_p_h.append(ch_p)
            energia_reativa_fp.append(er_fp)
            energia_reativa_p.append(er_p)
            er_p=0
            er_fp=0
            ch_fp=0
            ch_p=0
            i=0
        cont+=1
    
    custo_fp = consumo_fp*tc_fp
    custo_p = consumo_p*tc_p
    valores = [consumo_fp,consumo_p,custo_fp,custo_p,max_dem_fp,max_dem_p,c_fp_h,c_p_h,d_fp_h,d_p_h,energia_reativa_fp,energia_reativa_p]
    return valores

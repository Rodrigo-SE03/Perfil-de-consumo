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
    t = float(tarifas[0].replace(",",".")) if isinstance(tarifas,str) else float(tarifas[0])
    print(t)
    custo=0

    for pot in consumo_dict["Potência - kW"]:
        custo += t*float(pot)/60
    
    valores = [custo,'-']
    return valores

def calc_branca(tarifas,consumo_dict):
    t_fp = float(tarifas[0].replace(",",".")) if isinstance(tarifas,str) else float(tarifas[0])
    t_p = float(tarifas[1].replace(",",".")) if isinstance(tarifas,str) else float(tarifas[1])
    t_i = float(tarifas[2].replace(",",".")) if isinstance(tarifas,str) else float(tarifas[2])
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
    tc_fp = float(tarifas[0].replace(",",".")) if isinstance(tarifas,str) else float(tarifas[0])
    tc_p = float(tarifas[1].replace(",",".")) if isinstance(tarifas,str) else float(tarifas[1])

    td = float(tarifas[2].replace(",",".")) if isinstance(tarifas,str) else float(tarifas[2])

    consumo_fp=0
    consumo_p=0
    demanda=0

    for pot in consumo_dict["Potência FP - kW"]:
        consumo_fp += tc_fp*float(pot)/60
    for pot in consumo_dict["Potência P - kW"]:
        consumo_p += tc_p*float(pot)/60
    
    i=0
    max_dem = 0
    for pot in consumo_dict['Potência FP - kW']:
        while i<=15:
            demanda = float(pot)
            i+=1
        i=0
        if demanda>max_dem:
            max_dem=demanda
        demanda = 0
    
    consumo = consumo_p+consumo_fp
    demanda = max_dem*td
    valores = [consumo_fp,consumo_p,consumo,demanda]
    return valores

def calc_azul(tarifas,consumo_dict):
    tc_fp = float(tarifas[0].replace(",",".")) if isinstance(tarifas,str) else float(tarifas[0])
    tc_p = float(tarifas[1].replace(",",".")) if isinstance(tarifas,str) else float(tarifas[1])

    td_fp = float(tarifas[2].replace(",",".")) if isinstance(tarifas,str) else float(tarifas[2])
    td_p = float(tarifas[3].replace(",",".")) if isinstance(tarifas,str) else float(tarifas[3])

    consumo_fp=0
    consumo_p=0
    demanda_p=0
    demanda_fp=0

    for pot in consumo_dict["Potência FP - kW"]:
        consumo_fp += tc_fp*float(pot)/60
    for pot in consumo_dict["Potência P - kW"]:
        consumo_p += tc_p*float(pot)/60
    
    i=0
    max_dem_fp = 0
    for pot in consumo_dict['Potência FP - kW']:
        while i<=15:
            demanda_fp = float(pot)
            i+=1
        i=0
        if demanda_fp>max_dem_fp:
            max_dem_fp=demanda_fp
        demanda_fp = 0

    i=0
    max_dem_p = 0
    for pot in consumo_dict['Potência P - kW']:
        while i<=15:
            demanda_p = float(pot)
            i+=1
        i=0
        if demanda_p>max_dem_p:
            max_dem_p=demanda_p
        demanda_p = 0
    
    consumo = consumo_p+consumo_fp
    demanda = max_dem_fp*td_fp+max_dem_p*td_p
    valores = [consumo_fp,consumo_p,consumo,max_dem_fp,max_dem_p,demanda]
    return valores

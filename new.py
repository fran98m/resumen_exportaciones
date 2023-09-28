import pandas as pd
import os
from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from config import correlativas

###########################################################################################
def totales(totales_df: pd.DataFrame) -> None:
    ##################################### Resumen Inicial##############################################
      
    ano_actual = 2203
    ano_anterior = 2202
    
    
    #1. Primer Bullet
    
    # Calcular años
    ano_actual = correlativas[9]
    ano_anterior = correlativas[8]
    
    # Usar correlativas para referenciar columnas
    # Calcular las sumas de exportación y el crecimiento
    expt_act_tot = totales_df[correlativas[9]].sum()
    expt_ant_tot = totales_df[correlativas[8]].sum()
    
    var_exp_tot = ((expt_act_tot - expt_ant_tot) / expt_ant_tot) * 100
    tagvar_tot = "crecimiento" if var_exp_tot > 0 else "decrecimiento"
    
    #2. Segudo Bullet para las Exportaciones NME
    no_mineras_df = totales_df[totales_df[correlativas[0]] == "No Mineras"]
    
    #Se hace exactamente lo mismo que en el primer bullet pero con las exportaciones no mineras
    expt_act_tot_no_min = no_mineras_df[correlativas[9]].sum()
    expt_ant_tot_no_min = no_mineras_df[correlativas[8]].sum()
    var_nm_tot = round(((expt_act_tot_no_min - expt_ant_tot_no_min) / expt_ant_tot_no_min) * 100,1)
    tagvar_nm_tot = "crecimiento" if var_nm_tot > 0 else "decrecimiento"


    #3. Conteo de empresas
    
    #María Paula aquí te pongo la parte del código,muchas gracias por tu ayuda :)
    #Te lo dejo definido para que no te toque cambiar el word, solo no te olvides que debe retornar un escalar 
    #Que obligatoriamente se llame conteo_emp, sino me avisas y se ajusta después. 
    
    conteo_emp = 6101
    
    #4. Va el disclaimer. El tipo de dato es una tupla entonces se necesita referenciar la pos 0 de la tupla. 
    disclaimer = ("Vale la pena tener en cuenta que los datos que arroja el DANE/DIAN mes a mes: "
                  "tienen dos meses de rezago, no incluyen las exportaciones desde Zona Franca "
                  "y no tienen en cuenta los servicios diferentes a editorial. Es decir, la cifra de "
                  "exportaciones no minero energéticas puede ser mayor a la reportada")

    return {
        "ano_actual": ano_actual,
        "ano_anterior": ano_anterior,
        "expt_act_tot": expt_act_tot,
        "tagvar_tot": tagvar_tot,
        "var_exp_tot": var_exp_tot,
        "expt_ant_tot": expt_ant_tot,
        "expt_act_tot_no_min": expt_act_tot_no_min,
        "tagvar_nm_tot": tagvar_nm_tot,
        "var_nm_tot": var_nm_tot,
        "conteo_emp": conteo_emp,
        "expt_ant_tot_no_min": expt_ant_tot_no_min	
    }

    
#Se va a hacer una función para el resto del análisis, con solo las no mineras para simplificar los cálculos. 
#Tecnicamente no retorna un str sino que retorna un doc de word pero bueno es lo que hay jajaja

def no_mineras(df: pd.DataFrame,vars_from_totales) -> str:
    ####################Variables de totales necesarias para esta funcion porque el docx se genera aca#############################
    ano_actual = vars_from_totales["ano_actual"]
    ano_anterior = vars_from_totales["ano_anterior"]
    expt_act_tot = vars_from_totales["expt_act_tot"]
    tagvar_tot = vars_from_totales["tagvar_tot"]
    var_exp_tot = vars_from_totales["var_exp_tot"]
    expt_ant_tot = vars_from_totales["expt_ant_tot"]
    expt_act_tot_no_min = vars_from_totales["expt_act_tot_no_min"]
    tagvar_nm_tot = vars_from_totales["tagvar_nm_tot"]
    var_nm_tot = vars_from_totales["var_nm_tot"]
    conteo_emp = vars_from_totales["conteo_emp"]
    expt_ant_tot_no_min = vars_from_totales["expt_ant_tot_no_min"]
    ###############################################################################################################################
  
    # Se genera otro df de no mineras en este punto ya se tienen en cuenta las correlativas
    no_mineras_df = df[df[correlativas[0]] == "No Mineras"]

    #Se agrupa por país y se organizan los datos de mayor a menor para obtener los 10 países que más exportaron 
    agrupado_por_pais=no_mineras_df.groupby(correlativas[6])[correlativas[9]].sum().sort_values(ascending=False).head(10)
    exportado_10_principales=agrupado_por_pais.sum()
    


  ################################################################################################  
    #0. Resumen Inicial

    total_exportacion_actual = no_mineras_df[correlativas[9]].sort_values(ascending=False).head(10).sum() #Escalar

    #Lo mismo para el anterior
    total_exportacion_anterior=no_mineras_df[correlativas[8]].sort_values(ascending=False).head(10).sum() #Escalar 
 
    #La variación entre ambos años
    crecimiento_paises_tot = ((total_exportacion_actual - total_exportacion_anterior) / total_exportacion_anterior) * 100
    #Determina si crece o decrece que es importante para el word. 
    #Finalmente se calcula el porcentaje de exportación de los 10 países que más exportaron sobre el total

 ###################################################################################################################   
    #1. Análisis por países
    #Resumen Inicial:
    td_export_10_paises = df.groupby("Pais Destino")[["2023 USD (Ene-Jul)", "2022 USD (Ene-Jul)"]].sum().reset_index()
    primeros_10_dest = td_export_10_paises.sort_values(by="2023 USD (Ene-Jul)", ascending=False).head(10)
    print(primeros_10_dest)
    total_dest_act=primeros_10_dest["2023 USD (Ene-Jul)"].sum()
    total_dest_ant=primeros_10_dest["2022 USD (Ene-Jul)"].sum()
    print(total_dest_act)
    variacion_destinos=((total_dest_act-total_dest_ant)/total_dest_ant)*100
    tag_var_dest="crecimiento" if variacion_destinos>0 else "decrecimiento"
    porcentaje_destinos=(total_dest_act/expt_act_tot_no_min)*100
    print(porcentaje_destinos)
    print(variacion_destinos)
 

    # Función auxiliar para obtener los tres principales exportadores de un país recibe solo el país que está determinado antes
    def tres_principales_exportadores(pais:str):
        companies = (no_mineras_df[no_mineras_df['Pais Destino'] == pais]
                .groupby(correlativas[5])[correlativas[9]]
                .sum()
                .nlargest(4)  # Seleccionar 4 para tener espacio para eliminar 'NO DEFINIDO'
                .drop('NO DEFINIDO', errors='ignore')
                .nlargest(3))
        return companies
    #Función auxiliar para obtener la varianza de estos países
    def calculate_country_variance(pais:str):
        exports_current_year = no_mineras_df[no_mineras_df['Pais Destino'] == pais][correlativas[9]].sum()
        exports_previous_year = no_mineras_df[no_mineras_df['Pais Destino'] == pais][correlativas[8]].sum()
        if exports_previous_year == 0:
            return 0 if exports_current_year == 0 else float('inf')
        return ((exports_current_year - exports_previous_year) / exports_previous_year) * 100
    #Luego se genera un diccionario donde se van a guardar los datos que se combinan de antes.Es decir los 10 principales países que se le pasan a la función para determinar los principales exportadores
    datos_principales_exportadores = {}
    #En este for se itera por país y se van guardando los datos en el diccionario
    for pais in agrupado_por_pais.index:
        principales_exportadores = tres_principales_exportadores(pais)
        datos_principales_exportadores[pais] = {
            "Principales exportadores": principales_exportadores,
            "Variación": calculate_country_variance(pais),
            "Tag": "crecimiento" if calculate_country_variance(pais) > 0 else "decrecimiento"
        }
    print (datos_principales_exportadores)      
###################################################################################################################
    #2. Análisis por empresas
    # Se agrupa por razón social y se suman las exportaciones de cada empresa para 2022 y para 2023 

    grouped_by_razon = no_mineras_df.groupby(correlativas[5])[[correlativas[8], correlativas[9]]].sum()
    #Se añaden al dataframe las columnas de variación y tendencia
    grouped_by_razon["Variacion_rs"] = (grouped_by_razon[correlativas[8]] - grouped_by_razon[correlativas[9]]) / grouped_by_razon[correlativas[8]]
    grouped_by_razon["Tendencia"] = grouped_by_razon["Variacion_rs"].apply(lambda x: "Crecimiento" if x > 0 else "Decrecimiento")
    
    # Aqui se determinan las 10 empresas que más exportaron en 2023 pero se excluye el no definido del análisis
    top_10_empresas = grouped_by_razon[grouped_by_razon.index != "NO DEFINIDO"].sort_values(correlativas[9], ascending=False).head(10)
    top_10_companies_names = top_10_empresas.index.tolist()    
    
    #Diccionario para guardar los datos 
    analisis_empresas = {}

    # Ciclo para hacer el análisis de cada empresa
    for company in top_10_companies_names:
        company_data = no_mineras_df[no_mineras_df[correlativas[5]] == company]
        # Variacion y tendencia
        variacion_empresas = top_10_empresas.loc[company, "Variacion_rs"]
        variance_percentage_emp = variacion_empresas*100
        tendencia_empresas = top_10_empresas.loc[company, "Tendencia"]
        # Top 3 Departamentos de Origen
        top_departamentos = company_data.groupby(correlativas[7])[correlativas[9]].sum().sort_values(ascending=False).head(3)

        # Top 3 Destinos de Exportación
        top_destinos = company_data.groupby(correlativas[6])[correlativas[9]].sum().sort_values(ascending=False).head(3)
        total_export_to_destinos = top_destinos.sum()
        top_destinos_percentage = round((top_destinos / total_export_to_destinos) * 100, 1)
        
        #Totales
        total_2023_exports_emp = top_10_empresas.loc[company, correlativas[9]]

        #Se genera un diccionario con los datos de cada empresa para meter al diccionario total. 
        analisis_empresas[company] = {
            "Porcentaje": variance_percentage_emp,
            "Tendencia": tendencia_empresas,
            "Top Departamentos": top_departamentos,
            "Top Destinos": top_destinos,
            "Top Destinos Participación": top_destinos_percentage,
            "Total 2023 Exports": total_2023_exports_emp  # Including the 2023 total exports
        }

    # Resumen inicial 
      #Variacion
    primer_paso_var_emp=no_mineras_df.groupby(correlativas[5])[correlativas[9]].sum().nlargest(10).index
    segundo_paso_var_emp = no_mineras_df[no_mineras_df[correlativas[5]].isin(primer_paso_var_emp)]
    tercer_paso_var_emp = segundo_paso_var_emp.groupby(correlativas[5])[[correlativas[9], correlativas[8]]].sum()
    tercer_paso_var_emp["Variacion"] = ((tercer_paso_var_emp[correlativas[9]] - tercer_paso_var_emp[correlativas[8]]) 
                                         / tercer_paso_var_emp[correlativas[8]]) * 100
    
    valor_total_top_10_act_emp=tercer_paso_var_emp[correlativas[9]].sum()
    valor_total_top_10_ant_emp=tercer_paso_var_emp[correlativas[8]].sum()
    varianza_empresas=((valor_total_top_10_act_emp-valor_total_top_10_ant_emp)/valor_total_top_10_ant_emp)*100 
    
    tag_var_emp="crecimiento" if varianza_empresas>0 else "decrecimiento"
    #Porcentaje
    percentage_export_top_10=valor_total_top_10_act_emp/expt_act_tot_no_min*100

    print(analisis_empresas)
    

##########################################################################################
#3. Analisis por producto
# Determinar los 3 principales departamentos de origen y cuánto fue enviado en USD desde esos orígenes para cada subsector

    top_10_subsectors_2023 = no_mineras_df.groupby(correlativas[3])[correlativas[9]].sum().nlargest(10)
    top_3_origins_by_subsector = {}
    usd_from_top_3_origins_by_subsector = {}

    for subsector in top_10_subsectors_2023.index:
        subsector_data = no_mineras_df[no_mineras_df[correlativas[3]] == subsector]
    
    # Determinar los 3 principales departamentos de origen para el subsector
        top_3_origins = subsector_data.groupby(correlativas[7])[correlativas[9]].sum().nlargest(3).index.tolist()
        top_3_origins_by_subsector[subsector] = top_3_origins
    
    # Calcular el valor exportado desde esos 3 departamentos
        usd_from_top_3_origins = subsector_data[subsector_data[correlativas[7]].isin(top_3_origins)].groupby(correlativas[7])[correlativas[9]].sum().to_dict()
        usd_from_top_3_origins_by_subsector[subsector] = usd_from_top_3_origins

    # Combinamos ambos resultados anteriores para que sea más fácil después el análisis. 
    analisis_subsectores = {}

    for subsector in top_10_subsectors_2023.index:
        valor_total_sub = top_10_subsectors_2023[subsector]
        analisis_subsectores[subsector] = {
            'Top 3 Origins': top_3_origins_by_subsector[subsector],
            'USD from Top 3 Origins': usd_from_top_3_origins_by_subsector[subsector],
            "Valor exportado Actual": valor_total_sub
    }
    y1_values = no_mineras_df.groupby(correlativas[3])[correlativas[8]].sum().to_dict()

    for subsector, data in analisis_subsectores.items():
        y2_value = data["Valor exportado Actual"]
        y1_value = y1_values.get(subsector, 0)  # Default to 0 if not present in 2022 data
        if y1_value != 0:
            variation_sub = ((y2_value - y1_value) / y1_value)*100
        else:
            # Handle the case where the product wasn't present in 2022
            variation = 1  # or np.nan
        data["Variacion_sub"] = variation_sub
        data["Tag"] = "crecimiento" if variation_sub >= 0 else "decrecimiento"    


    print(analisis_subsectores)
##########################################################################################
    #4. Análisis por departamento
    # Grouping by "Departamento Origen" to sum the export values for 2022 and 2023
    grouped_by_departamento = no_mineras_df.groupby(correlativas[7])[[correlativas[8],correlativas[9]]].sum()

    # Calculate the variance between 2023 and 2022
    grouped_by_departamento["Variacion_dep"] = grouped_by_departamento[correlativas[9]] - grouped_by_departamento[correlativas[8]]

    # Determine if there's growth or decrease
    grouped_by_departamento["Tendencia"] = grouped_by_departamento["Variacion_dep"].apply(lambda x: "crecimiento" if x > 0 else ("decrecimiento" if x < 0 else "no cambió"))

    # Calculate the percentage of variance
    grouped_by_departamento["Variance Percentage"] = round((grouped_by_departamento["Variacion_dep"] / grouped_by_departamento["2022 USD (Ene-Jul)"]) * 100, 1)

    # Sorting by the 2023 values to get the Top 5 departments
    top_5_departamentos = grouped_by_departamento.sort_values(correlativas[9], ascending=False).head(5)
    top_5_departamentos[[correlativas[9], 'Variacion_dep', 'Tendencia', 'Variance Percentage']]

    #Resumen
    combined_value_2023 = top_5_departamentos[correlativas[9]].sum()
    combined_value_2022 = top_5_departamentos[correlativas[8]].sum()
    combined_variation = ((combined_value_2023 - combined_value_2022) / combined_value_2022)
    combined_percentage_variation = (combined_variation * 100)
    
    total_exports = grouped_by_departamento[correlativas[9]].sum()
    percentage_of_total = ((combined_value_2023 / total_exports) * 100)

    # Adding these calculations to the top_5_departamentos DataFrame
    top_5_departamentos.loc["COMBINED"] = [combined_value_2022, combined_value_2023, combined_variation, combined_percentage_variation, percentage_of_total]

    print(top_5_departamentos[[correlativas[8], correlativas[9], 'Variacion_dep', 'Tendencia', 'Variance Percentage']])
    print(top_5_departamentos)

#############################################################################################
    #5. Venezuela
    # Filtrar los datos para obtener solo las exportaciones a Venezuela con el nombre correcto
    venezuela_data = no_mineras_df[no_mineras_df[correlativas[6]] == 'Venezuela']

    # 1. Calcular el valor total exportado a Venezuela en 2023 y determinar la variación frente al año anterior
    total_exported_2023_venezuela = venezuela_data[correlativas[9]].sum()
    total_exported_2022_venezuela = venezuela_data[correlativas[8]].sum()

    variation_venezuela = ((total_exported_2023_venezuela - total_exported_2022_venezuela) / total_exported_2022_venezuela) * 100
    growth_label_venezuela = "crecimiento" if variation_venezuela >= 0 else "decrecimiento"
    formatted_variation_venezuela = f"{variation_venezuela:.1f}% ({growth_label_venezuela})"

    # 2. Identificar los 5 sectores con mayores exportaciones a Venezuela en 2023 y calcular el valor exportado por cada sector
    top_5_sectors_venezuela = venezuela_data.groupby(correlativas[2])[correlativas[9]].sum().nlargest(5)

    # 3. Determinar las 5 empresas que más exportan a Venezuela en 2023 y calcular el valor exportado por cada empresa
    top_5_companies_venezuela = venezuela_data.groupby(correlativas[5])[correlativas[9]].sum().nlargest(5)

    # Calcular la variación para estas empresas y añadir una etiqueta de "crecimiento" o "decrecimiento"
    exports_2022_companies_venezuela = venezuela_data.groupby(correlativas[5])[correlativas[8]].sum().loc[top_5_companies_venezuela.index]
    variations_companies_venezuela = ((top_5_companies_venezuela - exports_2022_companies_venezuela) / exports_2022_companies_venezuela) * 100
    growth_labels_companies = variations_companies_venezuela.apply(lambda x: "crecimiento" if x >= 0 else "decrecimiento")
    formatted_variations_companies = (variations_companies_venezuela.round(1).map(str) + "% (" + growth_labels_companies + ")").to_dict()

    total_exported_2023_venezuela, formatted_variation_venezuela, top_5_sectors_venezuela, top_5_companies_venezuela, formatted_variations_companies
    #Generate a dataframe with the results called results_venezuela
    results_venezuela = pd.DataFrame({
        "Valor Total Exportado": [total_exported_2023_venezuela],
        "Variación": [formatted_variation_venezuela],
        "Top 5 Sectores": [top_5_sectors_venezuela],
        "Top 5 Empresas": [top_5_companies_venezuela],
        "Variación Empresas": [formatted_variations_companies]
    })

    print(results_venezuela)
    return {
    "tag_var_dest": tag_var_dest,
    "variacion_destinos": variacion_destinos,
    "porcentaje_destinos": porcentaje_destinos,
    "agrupado_por_pais": agrupado_por_pais,
    "datos_principales_exportadores": datos_principales_exportadores,
    "exportado_10_principales": exportado_10_principales,
    "percentage_export_top_10": percentage_export_top_10,
    "analisis_empresas": analisis_empresas,
    "analisis_subsectores": analisis_subsectores,
    "grouped_by_departamento": grouped_by_departamento,
    "top_5_departamentos": top_5_departamentos,
    "total_exports": total_exports,
    "percentage_of_total": percentage_of_total,
    "combined_percentage_variation": combined_percentage_variation,
    "results_venezuela": results_venezuela,
    "growth_label_venezuela": growth_label_venezuela,
    "variation_venezuela": variation_venezuela,
    "top_5_sectors_venezuela": top_5_sectors_venezuela,
    "formatted_variations_companies": formatted_variations_companies,
    "varianza_empresas": varianza_empresas,
    "tag_var_emp": tag_var_emp,
    }

path=r"D:\usuarios\Pvein2\OneDrive - PROCOLOMBIA\Escritorio\Francisco\Corrección Resumen Export (Doc) (S)\Base.csv"
df=pd.read_csv(path,sep=";",encoding="latin-1",decimal=",",thousands=".")

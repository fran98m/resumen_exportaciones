##############################################################################################
from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from new import totales, no_mineras, mes_ano
from config import correlativas
import os
import pandas as pd 

def generar_docx(vars_from_totales:dict,vars_from_no_mineras: dict,vars_from_mes_ano:dict):
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
    tag_var_dest=vars_from_no_mineras["tag_var_dest"]
    variacion_destinos=vars_from_no_mineras["variacion_destinos"]
    porcentaje_destinos=vars_from_no_mineras["porcentaje_destinos"]
    agrupado_por_pais=vars_from_no_mineras["agrupado_por_pais"]
    datos_principales_exportadores=vars_from_no_mineras["datos_principales_exportadores"]
    exportado_10_principales=vars_from_no_mineras["exportado_10_principales"]
    percentage_export_top_10=vars_from_no_mineras["percentage_export_top_10"]
    analisis_empresas=vars_from_no_mineras["analisis_empresas"]
    analisis_subsectores=vars_from_no_mineras["analisis_subsectores"]
    top_5_departamentos=vars_from_no_mineras["top_5_departamentos"]
    total_exports= vars_from_no_mineras["total_exports"]
    percentage_of_total= vars_from_no_mineras["percentage_of_total"]
    combined_percentage_variation=vars_from_no_mineras["combined_percentage_variation"]
    results_venezuela=vars_from_no_mineras["results_venezuela"]
    growth_label_venezuela=vars_from_no_mineras["growth_label_venezuela"]
    variation_venezuela=vars_from_no_mineras["variation_venezuela"]
    top_5_sectors_venezuela=vars_from_no_mineras["top_5_sectors_venezuela"]
    formatted_variations_companies=vars_from_no_mineras["formatted_variations_companies"]
    varianza_empresas=vars_from_no_mineras["varianza_empresas"]
    tag_var_emp=vars_from_no_mineras["tag_var_emp"]
    mes=vars_from_mes_ano["mes"]
    ano=vars_from_mes_ano["ano"]
    ano_ant=vars_from_mes_ano["ano_ant"]
##############################################Inicialización del documento############################################################################################
    doc = Document()
###################################################################################################################################################################
    #0. Resumen Inicial
    # Se genera el título de la sección del resumen inicial 
    doc.add_heading(f'EXPORTACIONES Enero - {vars_from_mes_ano["mes"]} -- {vars_from_mes_ano["ano"]} (DANE-DIAN)', 1)

    # Totales
    p = doc.add_paragraph()
    p.add_run('- ').bold = True
    p.add_run(f"Las cifras Enero - {vars_from_mes_ano['mes']} de {vars_from_mes_ano['ano']}, las exportaciones totales de Colombia fueron {format_to_millions(vars_from_totales['expt_act_tot'])}, con un {vars_from_totales['tagvar_tot']} del {vars_from_totales['var_exp_tot']:.1f}% frente al mismo periodo de {vars_from_mes_ano['ano_ant']}  {format_to_millions(vars_from_totales['expt_ant_tot'])}.")

    # No mineras
    p = doc.add_paragraph()
    p.add_run('- ').bold = True
    p.add_run(f"Entre Enero - {vars_from_mes_ano['ano']} de {vars_from_mes_ano['ano']}, las exportaciones no minero energéticas de Colombia fueron {format_to_millions(vars_from_totales['expt_act_tot_no_min'])}, con un {vars_from_totales['tagvar_nm_tot']} del {vars_from_totales['var_nm_tot']:.1f}% de las exportaciones del mismo periodo de {vars_from_mes_ano['ano_ant']} {format_to_millions(vars_from_totales['expt_ant_tot_no_min'])}.")

    # Conteo Empresas
    p = doc.add_paragraph()
    p.add_run('- ').bold = True
    p.add_run(f"Entre Enero - {vars_from_mes_ano['ano']} de {vars_from_mes_ano['ano']}, un total de {vars_from_totales['conteo_emp']} empresas exportaron productos no minero energéticos por montos superiores a USD 10,000 desde Colombia.")

    # Disclaimer
    p = doc.add_paragraph()
    p.add_run('- ').bold = True
    disclaimer = ("Vale la pena tener en cuenta que los datos que arroja el DANE-DIAN mes a mes: tienen dos meses de rezago, no incluyen las exportaciones desde Zona Franca y no tienen en cuenta servicios diferentes a Editorial (es decir que la cifra real no minero energética podría ser más alta).")
    p.add_run(disclaimer)
###################################################################################################################################################################

    #1. Top 10 destinos de la exportación no minero energética y las empresas.
    doc.add_heading(f'Top 10 destinos exportaciones no minero energéticas Enero-{vars_from_mes_ano["ano"]} {vars_from_totales["ano_actual"]} (DANE-DIAN)', 1)
    # Resumen inicial. 
    doc.add_paragraph(f"- Los 10 principales destinos de las exportaciones no minero energéticas del país suman total USD {format_to_millions(exportado_10_principales)} millones.")
    doc.add_paragraph(f"- Con un {tag_var_dest} del {variacion_destinos:.1f}% en nuestras exportaciones no minero energéticas frente al mismo período del año anterior.")
    doc.add_paragraph(f"- Estos mercados representan el {porcentaje_destinos:.1f}% de las exportaciones no minero energéticas de Colombia entre Enero - {vars_from_mes_ano['ano']} de {vars_from_mes_ano['ano']}.")
   # Countries details
    for index, (country, value) in enumerate(agrupado_por_pais.items()):
        # Get the top 3 companies for the country
        top_companies = datos_principales_exportadores[country]["Principales exportadores"]
        companies_text = ", ".join([f"{company} (USD {format_to_millions(value)} millones)" for company, value in top_companies.items()])
        
        variance = datos_principales_exportadores[country]["Variación"]
        change_text = f"crecimiento de {variance:.1f}%" if variance >= 0 else f"decrecimiento de {-variance:.1f}%"

            # Add paragraph for each country
        p = doc.add_paragraph()
        runner = p.add_run(f"{index + 1}. {country} : USD {format_to_millions(value)} millones, {change_text} frente a Enero – {vars_from_mes_ano['ano']} {vars_from_mes_ano['ano_ant']}; principales exportadores: {companies_text}.")
        font = runner.font
        font.size = Pt(11)
###################################################################################################################################################################    

# 2. Top 10 empresas exportadoras no minero energéticas Enero -mes actual, año actual

    doc.add_heading(f'Top 10 empresas exportadoras no minero energéticas Enero -{vars_from_mes_ano["ano"]} de {vars_from_mes_ano["ano"]}', level=1)
    doc.add_paragraph(f'- Las 10 principales empresas exportadoras no minero energéticas del país suman total USD {format_to_millions(exportado_10_principales)} millones.')
    doc.add_paragraph(f'- Se ve un: {tag_var_emp} de sus exportaciones en {varianza_empresas:.1f}% frente al mismo periodo del año {vars_from_mes_ano["ano_ant"]}.')
    doc.add_paragraph(f'- Concentran el {percentage_export_top_10:.1f}% de las exportaciones no minero energéticas de Colombia entre Enero – {vars_from_mes_ano["mes"]} {vars_from_mes_ano["ano"]}.')
    doc.add_paragraph()  # Linea nueva en blanco para formato. 
    # Company details
    for idx, company in enumerate(analisis_empresas, start=1):
        company_info = analisis_empresas[company]
        departamentos = ', '.join(company_info["Top Departamentos"].index)
        destinos = ', '.join(company_info["Top Destinos"].index)
        doc.add_paragraph(f'{idx}. {company} -> USD {format_to_millions(company_info["Total 2023 Exports"])} millones, {company_info["Tendencia"].lower()} del {company_info["Porcentaje"]:.1f}% frente a Enero – {vars_from_mes_ano["ano"]} {vars_from_mes_ano["ano_ant"]}; origen: {departamentos} ; destino: {destinos}.')

######################################################################################################################################################
    # 3. Top 10 productos exportados no minero energéticos Enero -mes actual, año actual
    doc.add_heading(f'Top 10 productos exportados no minero energéticos Enero - {vars_from_mes_ano["ano"]} 2023', 1)
    
    # Compute and add aggregate information
    total_export_value = sum([details["Valor exportado Actual"] for details in analisis_subsectores.values()])
    overall_variation = sum([details["Variacion_sub"] for details in analisis_subsectores.values()]) / len(analisis_subsectores)
    
    doc.add_paragraph(f"• Los 10 principales productos exportados no minero energéticos suman un total de USD {total_export_value/1000000:.1f} millones.")
    doc.add_paragraph(f"• Presentan un {'crecimiento' if overall_variation > 0 else 'decrecimiento'} del {overall_variation:.1f}% frente a Enero - {vars_from_mes_ano['ano']} de {vars_from_mes_ano['ano_ant']}.")
    doc.add_paragraph(f"• Concentran el {overall_variation / percentage_export_top_10 * 100:.1f}% de las exportaciones no minero energéticas de Colombia entre Enero -{vars_from_mes_ano['ano']} de {vars_from_mes_ano['ano']}.")
    doc.add_paragraph()  # Linea nueva en blanco para formato. 

    # Iterate through the data and add to the document
    for idx, (subsector, details) in enumerate(analisis_subsectores.items(), 1):
        # Get the subsector details
        valor_exportado_actual = details["Valor exportado Actual"]
        variation = details["Variacion_sub"]
        tag = details["Tag"]
        
        # Format the origins and their values
        origins_str = ', '.join([f"{origin} (USD {value/1000000:.1f} millones)" for origin, value in details['USD from Top 3 Origins'].items()])

        # Add the formatted string to the document
        formatted_string = f"{idx}. {subsector}. USD {valor_exportado_actual/1000000:.1f} millones, {tag} del {variation:.1f}% frente a Enero -{vars_from_mes_ano['ano']} {vars_from_mes_ano['ano_ant']}; origen principal: {origins_str}."
        doc.add_paragraph(formatted_string)
###################################################################################################################################################################
    #4. Análisis por depatamentos
    doc.add_heading(f'Top 5 departamentos no-mineroenergéticos Enero-{vars_from_mes_ano["ano"]}', 1)
    doc.add_paragraph(f"• Los cinco principales departamentos exportadores no minero energéticos suman un total de USD {top_5_departamentos.loc['COMBINED', correlativas[9]]/1000000:.1f} millones.")
    doc.add_paragraph(f"• Presentan un: {'crecimiento' if combined_percentage_variation > 0 else 'decrecimiento'} sus exportaciones en un {abs(combined_percentage_variation):.1f} % frente a Enero – {vars_from_mes_ano['mes']} de {vars_from_mes_ano['ano']} {ano_anterior}.")
    doc.add_paragraph(f"• Concentran el: {percentage_of_total:.1f} % de las exportaciones no minero energéticas de Colombia en Enero – {vars_from_mes_ano['ano']} {ano_actual}.")

    # Loop through the top 5 departments and add their details
    for idx, (depto, row) in enumerate(top_5_departamentos.drop('COMBINED').iterrows(), start=1):
        tendencia = "crecimiento" if row['Variance Percentage'] > 0 else "decrecimiento"
        doc.add_paragraph(f"{idx}. {depto}. USD {row[correlativas[9]]/1000000:.1f} millones, {tendencia} del {abs(row['Variance Percentage']):.1f}% frente a Enero - {vars_from_mes_ano['ano']} {ano_anterior}.") 
###################################################################################################################################################################
    #6. Analisis de Venezuela:

    doc.add_heading('Venezuela', level=1)

    # 1. Add the summary of export growth
    doc.add_paragraph(f"• Entre Enero – {vars_from_mes_ano['ano']} del presente año las exportaciones no mineras hacia Venezuela han {growth_label_venezuela} en {variation_venezuela:.1f}%.")

    # 2. Add the top 5 sectors with the highest exports to Venezuela
    sectors_str = ', '.join(top_5_sectors_venezuela.index)
    doc.add_paragraph(f"• Los sectores con mayores exportaciones al mercado son: {sectors_str}.")

    # 3. Add the top 5 companies with the highest exports to Venezuela
    companies_str_list = [f"{company} ({variation})" for company, variation in formatted_variations_companies.items()]
    companies_str = ', '.join(companies_str_list)
    doc.add_paragraph(f"• Las empresas con mayores exportaciones son: {companies_str}.")

    return doc

########################################################################################################################################

#Generación del word: 

def format_to_millions(value: float) -> str:
    """
    Format the provided value to millions with one decimal place, using points for thousands and comma for decimals.
    """
    value_in_millions = value / 10**6
    return '{:,.1f}'.format(value_in_millions)
    

from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from new import totales, no_mineras, mes_ano
from config import correlativas
import logging

logging.basicConfig(filename="document_generation.log", level=logging.ERROR)

def generar_docx(vars_from_totales: dict, vars_from_no_mineras: dict, vars_from_mes_ano: dict) -> Document:
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
    analisis_empresas=vars_from_no_mineras["analisis_empresas"]
    analisis_subsectores=vars_from_no_mineras["analisis_subsectores"]
    top_5_departamentos=vars_from_no_mineras["top_5_departamentos"]
    percentage_of_total= vars_from_no_mineras["percentage_of_total"]
    combined_percentage_variation=vars_from_no_mineras["combined_percentage_variation"]
    results_venezuela=vars_from_no_mineras["results_venezuela"]
    growth_label_venezuela=vars_from_no_mineras["growth_label_venezuela"]
    variation_venezuela=vars_from_no_mineras["variation_venezuela"]
    top_5_sectors_venezuela=vars_from_no_mineras["top_5_sectors_venezuela"]
    formatted_variations_companies=vars_from_no_mineras["formatted_variations_companies"]
    mes=vars_from_mes_ano["mes"]
    ano=vars_from_mes_ano["ano"]
    ano_ant=vars_from_mes_ano["ano_ant"]
    total_productos=vars_from_no_mineras["total_productos"]
    variacion_productos=vars_from_no_mineras["var_productos"]
    tag_var_productos=vars_from_no_mineras["tag_var_prod"]
    var_empresas_resumen=vars_from_no_mineras["var_empresas_resumen"]
    tag_var_emp=vars_from_no_mineras["tag_var_empresas"]
    porcentaje_top10_emp=vars_from_no_mineras["porcentaje_top10_emp"]
    valor_exp_top_10_emp=vars_from_no_mineras["top_10_grouped_act"]
    
    #Se crea el documento donde se va a guardar la información
    doc = Document()

    try:
        # 0. Resumen Inicial

        # Titulo Principal
        doc.add_heading('Resumen de Exportaciones', level=0)
        doc.add_heading(f'Enero - {vars_from_mes_ano["mes"]} de 2023', level=1)
        doc.add_heading()  # Linea nueva en blanco para formato.

        # Se genera el título de la sección del resumen inicial 
        doc.add_heading(f'EXPORTACIONES Enero - {vars_from_mes_ano["mes"]} de {vars_from_mes_ano["ano"]} (DANE-DIAN)', 1)

        # Totales
        p = doc.add_paragraph()
        p.add_run('- ').bold = True
        p.add_run("Las cifras Enero - ").bold = False
        p.add_run(f"{vars_from_mes_ano['mes']}").bold = True
        p.add_run(f" de {vars_from_mes_ano['ano']}, las exportaciones totales de Colombia fueron USD$ ").bold = False
        p.add_run(f"{format_to_millions(vars_from_totales['expt_act_tot'])} millones,").bold = True
        p.add_run(f" con un {vars_from_totales['tagvar_tot']} del {vars_from_totales['var_exp_tot']:.1f}% frente al mismo periodo de {vars_from_mes_ano['ano_ant']} USD$ ").bold = False
        p.add_run(f"{format_to_millions(vars_from_totales['expt_ant_tot'])} millones.").bold = True

        # No mineras
        p = doc.add_paragraph()
        p.add_run('- ').bold = True
        p.add_run(f"Entre Enero - ").bold = False
        p.add_run(f"{vars_from_mes_ano['mes']}").bold = True
        p.add_run(f" de {vars_from_mes_ano['ano']}, las exportaciones no minero energéticas de Colombia fueron USD$ ").bold = False
        p.add_run(f"{format_to_millions(vars_from_totales['expt_act_tot_no_min'])} millones,").bold = True
        p.add_run(f" con un {vars_from_totales['tagvar_nm_tot']} del {vars_from_totales['var_nm_tot']:.1f}% de las exportaciones del mismo periodo de {vars_from_mes_ano['ano_ant']} USD$ ").bold = False
        p.add_run(f"{format_to_millions(vars_from_totales['expt_ant_tot_no_min'])} millones.").bold = True

        # Conteo Empresas
        p = doc.add_paragraph()
        p.add_run('- ').bold = True
        p.add_run(f"Entre Enero - {vars_from_mes_ano['ano']} de {vars_from_mes_ano['ano']}, un total de ").bold = False
        p.add_run(f"{vars_from_totales['conteo_emp']}").bold = True
        p.add_run(" empresas exportaron productos no minero energéticos por montos superiores a USD 10,000 desde Colombia.").bold = False

        # Disclaimer
        p = doc.add_paragraph()
        p.add_run('- ').bold = True
        disclaimer = ("Vale la pena tener en cuenta que los datos que arroja el DANE-DIAN mes a mes: tienen dos meses de rezago, no incluyen las exportaciones desde Zona Franca y no tienen en cuenta servicios diferentes a Editorial (es decir que la cifra real no minero energética podría ser más alta).")
        p.add_run(disclaimer).bold = False


        # 1. Top 10 destinos de la exportación no minero energética y las empresas.
        doc.add_heading(f'Top 10 destinos exportaciones no minero energéticas Enero-{vars_from_mes_ano["mes"]} de {vars_from_totales["ano_actual"]} (DANE-DIAN)', 1)

        # Resumen inicial.
        p = doc.add_paragraph("- Los 10 principales destinos de las exportaciones no minero energéticas del país suman total USD ")
        p.add_run(f"{format_to_millions(exportado_10_principales)} millones.").bold = True
        p = doc.add_paragraph("- Con un ")
        p.add_run(f"{tag_var_dest} del {variacion_destinos:.1f}%").bold = True
        p.add_run(" en nuestras exportaciones no minero energéticas frente al mismo período del año anterior.")
        p = doc.add_paragraph("- Estos mercados representan el ")
        p.add_run(f"{porcentaje_destinos:.1f}%").bold = True
        p.add_run(f" de las exportaciones no minero energéticas de Colombia entre Enero - {vars_from_mes_ano['mes']} de {vars_from_mes_ano['ano']}.")

        # Detalles de los países
        for index, (country, value) in enumerate(agrupado_por_pais.items()):
            # Obtener las 3 principales empresas para el país
            top_companies = datos_principales_exportadores[country]["Principales exportadores"]
            companies_text = ", ".join([f"{company} (USD {format_to_millions(value)} millones)" for company, value in top_companies.items()])
            
            variance = datos_principales_exportadores[country]["Variación"]
            change_text = f"crecimiento de {variance:.1f}%" if variance >= 0 else f"decrecimiento de {-variance:.1f}%"

            # Añadir párrafo para cada país
            p = doc.add_paragraph()
            runner = p.add_run(f"{index + 1}. ")
            runner.bold = False
            runner = p.add_run(f"{country} : ")
            runner.bold = True
            runner = p.add_run(f"USD {format_to_millions(value)} millones, {change_text} frente a Enero – {vars_from_mes_ano['mes']} de {vars_from_mes_ano['ano_ant']}; principales exportadores: {companies_text}.")
            runner.bold = False
            font = runner.font
            font.size = Pt(11)

        # 2. Top 10 empresas exportadoras no minero energéticas

        doc.add_heading(f'Top 10 empresas exportadoras no minero energéticas Enero -{vars_from_mes_ano["mes"]} de {vars_from_mes_ano["ano"]}', level=1)

        # Resumen inicial
        p = doc.add_paragraph('- Las 10 principales empresas exportadoras no minero energéticas del país suman total USD ')
        p.add_run(f"{format_to_millions(valor_exp_top_10_emp)} millones.").bold = True
        p = doc.add_paragraph('- Se ve un: ')
        p.add_run(f"{tag_var_emp} de sus exportaciones en {var_empresas_resumen:.1f}%").bold = True
        p.add_run(f" frente al mismo periodo del año {vars_from_mes_ano['ano_ant']}.")
        p = doc.add_paragraph('- Concentran el ')
        p.add_run(f"{porcentaje_top10_emp:.1f}%").bold = True
        p.add_run(f" de las exportaciones no minero energéticas de Colombia entre Enero – {vars_from_mes_ano['mes']} {vars_from_mes_ano['ano']}.")
        doc.add_paragraph()  # Línea nueva en blanco para formato.

        # Detalles de las empresas
        for idx, company in enumerate(analisis_empresas, start=1):
            company_info = analisis_empresas[company]

            # Formatear los departamentos y destinos
            departamentos = ', '.join(company_info["Top Departamentos"].index)
            destinos = ', '.join(company_info["Top Destinos"].index)

            # Agregar al documento
            p = doc.add_paragraph(f'{idx}. ')
            runner = p.add_run(f"{company}")
            runner.bold = True
            runner = p.add_run(f" --> USD {format_to_millions(company_info['Total 2023 Exports'])} millones, {company_info['Tendencia'].lower()} del {company_info['Porcentaje']:.1f}% frente a Enero – {vars_from_mes_ano['mes']} {vars_from_mes_ano['ano_ant']}; origen: ")
            runner.bold = False
            runner = p.add_run(departamentos)
            runner.bold = True
            runner = p.add_run(' ; destino: ')
            runner.bold = False
            runner = p.add_run(destinos)
            runner.bold = True
            runner = p.add_run('.')
            runner.bold = False

        # 3. Top 10 productos exportados no minero energéticos
        # ... [código existente para esta sección]

        # 4. Análisis por departamentos
        # ... [código existente para esta sección]

        # 5. Analisis de Venezuela
        # ... [código existente para esta sección]

    except Exception as e:
        logging.error(f"Error generando el documento: {str(e)}")
        # En caso de error, informa al usuario
        return f"Ocurrió un error al generar el documento: {str(e)}"

    return doc

def format_to_millions(value: float) -> str:
    """
    Format the provided value to millions with one decimal place, using points for thousands and comma for decimals.
    """
    value_in_millions = value / 10**6
    return '{:,.1f}'.format(value_in_millions)
    
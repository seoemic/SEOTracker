import sys
import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

def create_apa_document(md_filepath, output_docx):
    doc = docx.Document()

    # Define APA Page Setup (1 inch margins all sides = 2.54 cm)
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)

    # Base Styles Setup
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(12)
    normal_style.font.color.rgb = RGBColor(0x33, 0x33, 0x33) # Dark charcoal
    normal_style.paragraph_format.line_spacing = 1.5
    normal_style.paragraph_format.space_after = Pt(6)
    normal_style.paragraph_format.space_before = Pt(0)

    # Helper functions for APA elements
    def add_title_page():
        # APA 7 Student / Professional Title Page
        p_space = doc.add_paragraph()
        p_space.paragraph_format.space_before = Pt(72) # Space down from top

        title_p = doc.add_paragraph()
        title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title_p.add_run("Documentación del Sistema SEOTracker\n(SEO Auditor)")
        title_run.font.name = 'Times New Roman'
        title_run.font.size = Pt(20)
        title_run.font.bold = True
        title_run.font.color.rgb = RGBColor(0x11, 0x22, 0x44)

        sub_p = doc.add_paragraph()
        sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sub_p.paragraph_format.space_after = Pt(48)
        sub_run = sub_p.add_run("Manual de Operación, Arquitectura Técnica, Roles, Instalación Local y Guía de Despliegue en VPS Linux")
        sub_run.font.name = 'Times New Roman'
        sub_run.font.size = Pt(13)
        sub_run.font.italic = True
        sub_run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

        info_p = doc.add_paragraph()
        info_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        info_p.paragraph_format.line_spacing = 1.5
        
        details = [
            ("Autor / Desarrollador:", "Fabian Pérez"),
            ("Organización / Departamento:", "Equipo de Desarrollo y Operaciones SEO"),
            ("Proyecto:", "Plataforma de Auditoría Automática SEOTracker"),
            ("Estándar de Documentación:", "Normas APA 7.ª Edición"),
            ("Fecha:", "Julio de 2026")
        ]
        
        for label, val in details:
            r1 = info_p.add_run(f"{label} ")
            r1.font.bold = True
            r1.font.size = Pt(11)
            r2 = info_p.add_run(f"{val}\n")
            r2.font.size = Pt(11)

        doc.add_page_break()

    def add_apa_heading(text, level):
        p = doc.add_paragraph()
        p.paragraph_format.keep_with_next = True
        
        if level == 1:
            # Level 1: Centered, Bold, Title Case
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(18)
            p.paragraph_format.space_after = Pt(8)
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(16)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
        elif level == 2:
            # Level 2: Flush Left, Bold, Title Case
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(6)
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(14)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0x11, 0x44, 0x77)
        elif level == 3:
            # Level 3: Flush Left, Bold Italic, Title Case
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12.5)
            run.font.bold = True
            run.font.italic = True
            run.font.color.rgb = RGBColor(0x22, 0x55, 0x88)
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            run.font.bold = True

    def add_apa_paragraph(text, bold_prefix=None, italic=False, bullet=False):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_after = Pt(4)
        
        if bullet:
            p.paragraph_format.left_indent = Inches(0.3)
            r_bullet = p.add_run("• ")
            r_bullet.font.bold = True
            r_bullet.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

        if bold_prefix:
            r_bold = p.add_run(bold_prefix)
            r_bold.font.name = 'Times New Roman'
            r_bold.font.size = Pt(12)
            r_bold.font.bold = True
            r_bold.font.color.rgb = RGBColor(0x11, 0x22, 0x44)

        r_text = p.add_run(text)
        r_text.font.name = 'Times New Roman'
        r_text.font.size = Pt(12)
        r_text.font.italic = italic
        return p

    def add_code_block(code_text, lang=""):
        # APA Code Block / Snippet styling
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)
        
        # Set cell background to light grayish blue #F4F6F9
        shading = parse_xml(r'<w:shd {} w:fill="F4F6F9"/>'.format(nsdecls('w')))
        cell._tc.get_or_add_tcPr().append(shading)

        # Set subtle border left
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = parse_xml(r'''
            <w:tcBorders {} >
                <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>
                <w:left w:val="single" w:sz="24" w:space="0" w:color="003366"/>
                <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>
                <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            </w:tcBorders>
        '''.format(nsdecls('w')))
        tcPr.append(tcBorders)

        p = cell.paragraphs[0]
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.left_indent = Inches(0.1)
        p.paragraph_format.right_indent = Inches(0.1)

        run = p.add_run(code_text.strip())
        run.font.name = 'Consolas'
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0x1A, 0x25, 0x36)

        # Space after code table
        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(4)

    def style_apa_table(table, num_str, title_str):
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Add APA Table Header (Table X in bold, Title in italics)
        p_num = doc.add_paragraph()
        p_num.paragraph_format.space_before = Pt(12)
        p_num.paragraph_format.space_after = Pt(2)
        p_num.paragraph_format.keep_with_next = True
        r_num = p_num.add_run(num_str)
        r_num.font.name = 'Times New Roman'
        r_num.font.bold = True
        r_num.font.size = Pt(11)

        p_title = doc.add_paragraph()
        p_title.paragraph_format.space_before = Pt(0)
        p_title.paragraph_format.space_after = Pt(6)
        p_title.paragraph_format.keep_with_next = True
        r_title = p_title.add_run(title_str)
        r_title.font.name = 'Times New Roman'
        r_title.font.italic = True
        r_title.font.size = Pt(11)

        # Apply APA Table Borders (top border, bottom of header, bottom of table)
        for r_idx, row in enumerate(table.rows):
            trPr = row._tr.get_or_add_trPr()
            trPr.append(parse_xml(r'<w:cantSplit {}/>'.format(nsdecls('w')))) # Prevent page split inside row
            
            for c_idx, cell in enumerate(row.cells):
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                tcPr = cell._tc.get_or_add_tcPr()
                
                # Header row shading
                if r_idx == 0:
                    shd = parse_xml(r'<w:shd {} w:fill="E9EEF4"/>'.format(nsdecls('w')))
                    tcPr.append(shd)

                # Format borders per APA (Horizontal lines only: top, header bottom, table bottom)
                top_val = "single" if r_idx == 0 else "none"
                top_sz = "12" if r_idx == 0 else "0"
                top_color = "003366" if r_idx == 0 else "auto"

                bottom_val = "single" if (r_idx == 0 or r_idx == len(table.rows)-1) else "none"
                bottom_sz = "12" if r_idx == len(table.rows)-1 else ("8" if r_idx == 0 else "0")
                bottom_color = "003366" if (r_idx == 0 or r_idx == len(table.rows)-1) else "auto"

                borders = parse_xml(r'''
                    <w:tcBorders {} >
                        <w:top w:val="{}" w:sz="{}" w:space="0" w:color="{}"/>
                        <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>
                        <w:bottom w:val="{}" w:sz="{}" w:space="0" w:color="{}"/>
                        <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>
                    </w:tcBorders>
                '''.format(nsdecls('w'), top_val, top_sz, top_color, bottom_val, bottom_sz, bottom_color))
                tcPr.append(borders)

                # Format text inside cells
                for p in cell.paragraphs:
                    p.paragraph_format.line_spacing = 1.15
                    p.paragraph_format.space_before = Pt(3)
                    p.paragraph_format.space_after = Pt(3)
                    for run in p.runs:
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(10.5)
                        if r_idx == 0:
                            run.font.bold = True
                            run.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    # -------------------------------------------------------------
    # BUILD DOCUMENT CONTENT ACCORDING TO APA 7
    # -------------------------------------------------------------

    # 1. Title Page
    add_title_page()

    # 2. Resumen / Abstract
    add_apa_heading("Resumen Ejecutivo", level=1)
    add_apa_paragraph(
        "El presente documento técnico proporciona la documentación oficial integral del sistema SEOTracker (SEO Auditor). "
        "Se detalla la arquitectura de software desarrollada sobre Python y Flask, el motor de extracción web en vivo con BeautifulSoup y requests, "
        "la integración automatizada con Google Sheets API y la persistencia en MySQL. Asimismo, se expone la matriz de roles y permisos del sistema, "
        "el manual de usuario paso a paso para cada uno de sus módulos operativos, la guía de instalación local y la metodología completa de despliegue en un entorno VPS privado con Nginx, Gunicorn y certificados SSL HTTPS mediante Let's Encrypt."
    )
    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 3. Sección 1: Visión General y Arquitectura
    add_apa_heading("1. Visión General y Arquitectura (Cómo Funciona)", level=1)
    add_apa_paragraph(
        "SEOTracker es una plataforma web orientada a la auditoría automatizada y monitoreo continuo de parámetros SEO on-page en portales web y plataformas de comercio electrónico. "
        "El sistema combina componentes de extracción de datos en tiempo real, procesamiento concurrente, análisis sintáctico HTML y generación automatizada de reportes gerenciales en formato hojas de cálculo Excel."
    )

    add_apa_heading("Flujo Operativo Extremo a Extremo", level=2)
    
    # Diagrama de flujo en formato APA Figure
    p_fig_num = doc.add_paragraph()
    p_fig_num.paragraph_format.space_before = Pt(10)
    p_fig_num.paragraph_format.space_after = Pt(2)
    r_fn = p_fig_num.add_run("Figura 1")
    r_fn.font.name = 'Times New Roman'
    r_fn.font.bold = True
    r_fn.font.size = Pt(11)

    p_fig_title = doc.add_paragraph()
    p_fig_title.paragraph_format.space_before = Pt(0)
    p_fig_title.paragraph_format.space_after = Pt(6)
    r_ft = p_fig_title.add_run("Diagrama de flujo de datos y arquitectura de procesamiento de SEOTracker")
    r_ft.font.name = 'Times New Roman'
    r_ft.font.italic = True
    r_ft.font.size = Pt(11)

    flowchart_ascii = (
        "+-------------------+       +-----------------------+       +------------------------+\n"
        "|  Google Sheets    | ----> |   API gspread         | ----> |  SEOTracker Core       |\n"
        "| (Contenido SEO)   |       | (Service Account JSON)|       | (Flask Engine)         |\n"
        "+-------------------+       +-----------------------+       +-----------+------------+\n"
        "                                                                        |\n"
        "                                                                        v (Scraping Concurrente)\n"
        "+-------------------+       +-----------------------+       +-----------+------------+\n"
        "| Notificación Email| <---- | Reportes Excel (.xlsx)| <---- |  Sitio Web en Vivo     |\n"
        "| (SMTP HTML)       |       | y Persistencia MySQL  |       | (BeautifulSoup4)       |\n"
        "+-------------------+       +-----------------------+       +------------------------+"
    )
    add_code_block(flowchart_ascii)

    add_apa_heading("Etapas del Proceso de Auditoría", level=2)
    add_apa_paragraph("Definición de Contenidos Esperados: ", bold_prefix="1. ", bullet=True)
    add_apa_paragraph("El equipo SEO mantiene actualizado un documento en Google Sheets para cada marca, definiendo los textos objetivo para etiquetas como Meta Titulo, Meta Descripción, H1, Descripción Corta y Descripción Larga.")
    
    add_apa_paragraph("Conexión con Google Drive API: ", bold_prefix="2. ", bullet=True)
    add_apa_paragraph("Mediante autenticación OAuth2 con cuenta de servicio (credentials.json), el módulo auditor descarga las hojas de cálculo sin requerir interacción manual.")

    add_apa_paragraph("Extracción Web Concurrente: ", bold_prefix="3. ", bullet=True)
    add_apa_paragraph("A través de ThreadPoolExecutor y solicitudes HTTP con la librería requests, el sistema rastrea simultáneamente las URLs. Parsea el código fuente utilizando BeautifulSoup4 y extrae los contenidos basándose en los selectores CSS asignados a la marca.")

    add_apa_paragraph("Soporte Responsive y No Responsive: ", bold_prefix="4. ", bullet=True)
    add_apa_paragraph("Para sitios declarados no responsive, el motor desencadena dos ciclos paralelos independientes utilizando User-Agents específicos para simular navegadores Mobile y Desktop.")

    add_apa_paragraph("Normalización y Comparación: ", bold_prefix="5. ", bullet=True)
    add_apa_paragraph("Limpia marcas HTML, resuelve caracteres especiales, elimina palabras excluidas parametrizadas y compara la exactitud del contenido rastreado frente al contenido del Google Sheet.")

    add_apa_paragraph("Generación de Reportes e Histórico: ", bold_prefix="6. ", bullet=True)
    add_apa_paragraph("Construye archivos Excel (.xlsx) con celdas resaltadas en color rojo (#FFB6C1) en caso de discrepancia. Guarda las métricas detalladas en la tabla MySQL historico, calculando errores arreglados y cambiados respecto al rastreo previo.")

    add_apa_paragraph("Notificación y Escalado por Correo: ", bold_prefix="7. ", bullet=True)
    add_apa_paragraph("Envía correos SMTP formateados en HTML con tablas comparativas. Si los errores persisten al alcanzar la Revisión #3 copia al Líder SEO, y en la Revisión #4+ copia al Líder de Operaciones.")

    add_apa_paragraph("Planificación Automática (Scheduler): ", bold_prefix="8. ", bullet=True)
    add_apa_paragraph("Un servicio en segundo plano (schedule) revisa la base de datos diariamente a las 02:00 AM y ejecuta la auditoría automática de cualquier marca cuya última revisión supere los 15 días de antigüedad.")

    # 4. Sección 2: Funcionalidades Principales
    add_apa_heading("2. Funcionalidades Principales (Qué Hace)", level=1)
    
    funcs = [
        ("Dashboard Ejecutivo Unificado: ", "Muestra tarjetas de estado por marca, fecha de la última auditoría, conteo de discrepancias, temporizador de próxima auditoría programada y estado de jobs en vivo."),
        ("Auditoría en Vivo a Demanda: ", "Permite iniciar el proceso completo de extracción y análisis de cualquier marca con un solo clic desde la interfaz."),
        ("Rastreo por Lista Manual: ", "Herramienta interactiva para auditar un listado arbitrario de URLs pegadas por el usuario. Visualiza resultados al instante en pantalla y permite el envío del reporte por email."),
        ("Gestión de Células Operativas: ", "Organiza el personal en células integradas por un Implementador, un Gestor de Cuenta y un Project Manager (PM)."),
        ("Configuración Parametrizada por Marca: ", "Asignación de selectores CSS específicos, país, entorno (producción/staging), exclusión de palabras y modo responsive/no responsive."),
        ("Módulo de Reportes & Analítica: ", "Tarjetas KPI globales, gráfico de dona para la distribución de errores, gráfico de evolución temporal con Chart.js y tabla de marcas activas."),
        ("Herramienta de Diagnóstico Integrado: ", "Endpoint accesible en /diagnostico/<id> para validar credenciales de Google, disponibilidad del archivo credentials.json y estructura de encabezados."),
        ("Notificaciones Inteligentes: ", "Creación automática de usuarios con envío de credenciales iniciales por correo y notificaciones HTML con archivos adjuntos.")
    ]
    for b_title, desc in funcs:
        add_apa_paragraph(desc, bold_prefix=b_title, bullet=True)

    # 5. Sección 3: Tiempos Aproximados de Rastreo
    add_apa_heading("3. Tiempos Aproximados de Rastreo", level=1)
    add_apa_paragraph(
        "El motor de rastreo aplica un retardo aleatorio de entre 1.5 y 4.0 segundos entre peticiones HTTP consecutivas "
        "y un tiempo de espera máximo (timeout) de 15 segundos por URL. Este comportamiento es intencional para "
        "evitar bloqueos o rate-limiting por parte de los servidores auditados."
    )
    add_apa_paragraph(
        "Fórmula de referencia: Tiempo total ≈ N_urls × (retraso_promedio + tiempo_servidor), "
        "donde retraso_promedio ≈ 2.75 s y tiempo_servidor ≈ 0.5–3 s (varía según el sitio).",
        italic=True
    )

    add_apa_heading("3.1. Rastreo por Lista Manual (Modo Lista)", level=2)
    add_apa_paragraph(
        "En el modo de Lista Manual el sistema procesa las URLs de forma secuencial (max_workers=1), "
        "priorizando la estabilidad de la conexión sobre la velocidad de procesamiento."
    )

    table3_data = [
        ["Situación", "Parámetro Técnico", "Tiempo Est. por URL", "Tiempo Total (50 URLs)"],
        ["Marca Responsive",    "1 ciclo · 1 worker secuencial",             "3 – 7 seg/URL",     "~3 – 6 min"],
        ["Marca No Responsive", "2 ciclos paralelos (Mobile + Desktop)",     "3 – 7 seg/URL × 2", "~6 – 12 min"],
    ]
    t3 = doc.add_table(rows=len(table3_data), cols=4)
    for r_i, row in enumerate(table3_data):
        for c_i, val in enumerate(row):
            t3.cell(r_i, c_i).text = val
    style_apa_table(t3, "Tabla 3", "Tiempos estimados de rastreo en modo Lista Manual según tipo de sitio")
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    add_apa_paragraph(
        "Nota: Para marcas No Responsive, los ciclos Mobile y Desktop se ejecutan en hilos paralelos entre sí, "
        "por lo que el tiempo total equivale al ciclo más lento —no a la suma de ambos."
    )
    add_apa_paragraph(
        "Al finalizar el rastreo de una marca No Responsive, el sistema muestra automáticamente "
        "dos tablas apiladas en pantalla (📱 Mobile y 🖥 Desktop) con pestañas para alternar entre "
        "'Solo Mobile', 'Solo Desktop' o 'Ambos'. El correo enviado incluye ambos archivos Excel adjuntos."
    )

    add_apa_heading("3.2. Auditoría Completa (desde Dashboard o Automática)", level=2)
    add_apa_paragraph(
        "La auditoría completa se ejecuta con mayor concurrencia (max_workers=2), permitiendo rastrear "
        "dos URLs en paralelo y reduciendo el tiempo total respecto al Modo Lista. Para marcas No Responsive, "
        "los ciclos Mobile y Desktop se ejecutan de forma secuencial (uno tras otro)."
    )

    table4_data = [
        ["Situación", "Parámetro Técnico", "Tiempo Est. por URL", "Tiempo Total (200 URLs)"],
        ["Marca Responsive",    "1 ciclo · 2 workers concurrentes",          "1.5 – 5 seg/URL",     "~5 – 17 min"],
        ["Marca No Responsive", "2 ciclos secuenciales (Mobile → Desktop)",  "1.5 – 5 seg/URL × 2", "~10 – 34 min"],
    ]
    t4 = doc.add_table(rows=len(table4_data), cols=4)
    for r_i, row in enumerate(table4_data):
        for c_i, val in enumerate(row):
            t4.cell(r_i, c_i).text = val
    style_apa_table(t4, "Tabla 4", "Tiempos estimados de rastreo en Auditoría Completa según tipo de sitio")
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    add_apa_heading("3.3. Resumen Comparativo Rápido", level=2)

    table5_data = [
        ["Modo",               "Tipo de Marca",  "URLs", "Tiempo Estimado"],
        ["Lista Manual",       "Responsive",     "10",   "~30 seg – 1.5 min"],
        ["Lista Manual",       "Responsive",     "50",   "~3 – 6 min"],
        ["Lista Manual",       "No Responsive",  "10",   "~1 – 2.5 min"],
        ["Lista Manual",       "No Responsive",  "50",   "~6 – 12 min"],
        ["Auditoría Completa", "Responsive",     "100",  "~3 – 9 min"],
        ["Auditoría Completa", "Responsive",     "200",  "~5 – 17 min"],
        ["Auditoría Completa", "No Responsive",  "100",  "~6 – 17 min"],
        ["Auditoría Completa", "No Responsive",  "200",  "~10 – 34 min"],
    ]
    t5 = doc.add_table(rows=len(table5_data), cols=4)
    for r_i, row in enumerate(table5_data):
        for c_i, val in enumerate(row):
            t5.cell(r_i, c_i).text = val
    style_apa_table(t5, "Tabla 5", "Resumen comparativo de tiempos de rastreo por modo y tipo de sitio")
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    add_apa_paragraph(
        "Advertencia: Los tiempos son estimaciones de referencia. El tiempo real varía según la velocidad "
        "de respuesta del servidor auditado, la carga del servidor SEOTracker y la velocidad de la "
        "conexión a internet del VPS donde está desplegado el sistema.",
        italic=True
    )

    # 6. Sección 4: Alcance y Limitaciones
    add_apa_heading("4. Alcance y Limitaciones (Qué NO Hace)", level=1)
    add_apa_paragraph(
        "Con el fin de delimitar el alcance operativo del sistema y evitar falsas expectativas tecnológicas, a continuación se establecen las capacidades incluidas y excluidas en el diseño del software."
    )

    # Tabla 1 APA
    table1_data = [
        ["Capacidad Incluida (Qué SI Hace)", "Limitación Tecnológica / Alcance Excluido (Qué NO Hace)"],
        ["Scraping e inspección HTML en servidor (SSR/HTML estático)", "NO ejecuta Javascript del lado del cliente (sin headless browsers como Puppeteer o Playwright). Páginas 100% Client-Side Rendered (CSR) sin SSR no extraen datos JS dinámicos."],
        ["Detección y reporte de discrepancias entre etiquetas", "NO modifica el código fuente ni realiza ediciones automáticas en el CMS (WordPress, Shopify, Magento, etc.)."],
        ["Lectura directa de hojas de cálculo de Google Sheets", "NO edita ni altera los datos almacenados en el archivo Google Sheets origen."],
        ["Auditoría de etiquetas SEO de contenido (Title, Meta Description, H1, etc.)", "NO realiza auditorías SEO técnicas avanzadas (Core Web Vitals, sitemaps XML, robots.txt, tags canonical o Schema markup)."],
        ["Notificaciones por correo y exportación formateada a Excel", "NO incluye pasarela de pagos ni facturación multitenant SaaS integrada."],
        ["Autenticación de usuarios por email y hash de contraseña", "NO cuenta con Autenticación de Dos Factores (2FA) ni Inicio de Sesión Social (Google OAuth login)."]
    ]
    
    t1 = doc.add_table(rows=len(table1_data), cols=2)
    for r_i, row in enumerate(table1_data):
        t1.cell(r_i, 0).text = row[0]
        t1.cell(r_i, 1).text = row[1]
    style_apa_table(t1, "Tabla 6", "Comparativa de alcance funcional y limitaciones tecnológicas de SEOTracker")
    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 7. Sección 5: Matriz de Roles y Permisos
    add_apa_heading("5. Matriz de Roles y Permisos", level=1)
    add_apa_paragraph(
        "El control de acceso basado en roles (RBAC) está implementado en la base de datos mediante el tipo ENUM('lider', 'pm', 'gestor_cuenta', 'implementador') "
        "y es validado a nivel de controlador en Flask a través de decoradores personalizados."
    )

    # Tabla 2 APA
    table2_data = [
        ["Módulo / Acción", "Líder (lider)", "Project Manager (pm)", "Gestor de Cuenta (gestor_cuenta)", "Implementador (implementador)"],
        ["Ver Dashboard & Reportes Globales", "Todos los registros", "Solo marcas de su célula", "Solo marcas de su célula", "Solo marcas de su célula"],
        ["Ejecutar Auditoría en Vivo", "Permitido", "Restringido", "Restringido", "Permitido"],
        ["Rastreo por Lista Manual", "Permitido", "Permitido", "Permitido", "Permitido"],
        ["Crear / Editar Marcas", "Permitido", "Permitido", "Restringido", "Restringido"],
        ["Eliminar Marcas", "Permitido", "Restringido", "Restringido", "Restringido"],
        ["Crear / Editar Células Operativas", "Permitido", "Permitido", "Restringido", "Restringido"],
        ["Eliminar Células Operativas", "Permitido", "Restringido", "Restringido", "Restringido"],
        ["Gestión de Usuarios (CRUD)", "Permitido", "Restringido", "Restringido", "Restringido"],
        ["Configuración Global (SMTP/Líderes)", "Permitido", "Restringido", "Restringido", "Restringido"],
        ["Ejecutar Diagnóstico de Marca", "Permitido", "Restringido", "Restringido", "Restringido"]
    ]
    t2 = doc.add_table(rows=len(table2_data), cols=5)
    for r_i, row in enumerate(table2_data):
        for c_i, val in enumerate(row):
            t2.cell(r_i, c_i).text = val
    style_apa_table(t2, "Tabla 7", "Matriz de acceso y permisos por rol dentro de la aplicación")
    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 8. Sección 6: Manual de Usuario Paso a Paso
    add_apa_heading("6. Manual de Usuario Paso a Paso", level=1)
    
    add_apa_heading("6.1. Inicio de Sesión y Gestión de Perfil", level=2)
    add_apa_paragraph("1. Ingrese a la URL del sistema e introduzca su dirección de correo electrónico y contraseña asignada.")
    add_apa_paragraph("2. Para modificar sus datos personales o cambiar su clave, haga clic en el nombre de usuario situado en la barra superior derecha y seleccione la opción 'Mi Perfil'.")

    add_apa_heading("6.2. Navegación en el Dashboard Principal", level=2)
    add_apa_paragraph("• Cada tarjeta representa una marca registrada.")
    add_apa_paragraph("• Haga clic en 'Auditar Ahora' para forzar una auditoría inmediata en segundo plano.")
    add_apa_paragraph("• Utilice el enlace 'Ver Histórico' para revisar el archivo de auditorías anteriores y descargar los Excel correspondientes.")

    add_apa_heading("6.3. Uso de la Herramienta de Rastreo por Lista", level=2)
    add_apa_paragraph("1. Seleccione 'Rastreo por Lista' en el menú principal.")
    add_apa_paragraph("2. Elija la marca deseada o busque la coincidencia mediante el dominio.")
    add_apa_paragraph("3. Pegue el listado de URLs que desea inspeccionar (una por línea) y presione 'Iniciar Rastreo'.")
    add_apa_paragraph(
        "4. Al finalizar, el sistema mostrará los resultados en pantalla:\n"
        "   • Marca Responsive: una tabla única con los resultados.\n"
        "   • Marca No Responsive: dos tablas apiladas (Mobile y Desktop) visibles simultáneamente, "
        "con pestañas 'Solo Mobile', 'Solo Desktop' y 'Ambos' (activa por defecto)."
    )
    add_apa_paragraph(
        "5. Seleccione los destinatarios de la célula y haga clic en 'Enviar Reporte por Correo'. "
        "Para marcas No Responsive, el correo incluirá automáticamente ambos archivos Excel adjuntos en un único mensaje."
    )

    add_apa_heading("6.4. Configuración de Células y Marcas (Administradores y PMs)", level=2)
    add_apa_paragraph("• Células: En la sección 'Células', cree un nuevo equipo asignando los correos del Implementador, Gestor y PM. El sistema dará de alta de forma automática a los usuarios inexistentes.")
    add_apa_paragraph("• Marcas: En la sección 'Marcas', registre el enlace del Google Sheet y defina las clases CSS correspondientes a las etiquetas H1, Descripción Corta y Descripción Larga.")

    add_apa_heading("6.5. Análisis e Interpretación de Reportes", level=2)
    add_apa_paragraph("• La vista 'Reportes' presenta gráficas dinámicas de la evolución de errores en el tiempo.")
    add_apa_paragraph("• Filtre por marca específica para medir la efectividad de las correcciones implementadas entre revisiones.")

    # 9. Sección 7: Guía de Instalación Local
    add_apa_heading("7. Guía de Instalación Local", level=1)
    add_apa_paragraph("Para desplegar el entorno de desarrollo en un equipo local, siga la secuencia de comandos descrita a continuación:")

    add_apa_heading("Paso 1: Clonar el Repositorio e Instalar Dependencias", level=3)
    add_code_block(
        "git clone <URL_REPOSITORIO>\n"
        "cd \"Trabajo Python\"\n"
        "python -m venv venv\n"
        "# En Windows:\n"
        ".\\venv\\Scripts\\Activate.ps1\n"
        "# En Linux/Mac:\n"
        "source venv/bin/activate\n"
        "pip install -r requirements.txt",
        lang="bash"
    )

    add_apa_heading("Paso 2: Inicializar la Base de Datos MySQL", level=3)
    add_code_block(
        "python setup_db.py\n"
        "# Ingrese usuario MySQL (ej. root) y contraseña cuando sea solicitado.\n"
        "# Se creará la base de datos 'seo_auditor' y el usuario admin@seoauditor.com / admin123",
        lang="bash"
    )

    add_apa_heading("Paso 3: Configurar Credenciales y Ejecutar Servidor", level=3)
    add_apa_paragraph("Ubique el archivo credentials.json obtenido desde Google Cloud Console en la raíz del proyecto e inicie la aplicación:")
    add_code_block(
        "python app.py\n"
        "# Servidor activo en http://localhost:5000",
        lang="bash"
    )

    # 10. Sección 8: Despliegue en VPS Online
    add_apa_heading("8. Guía de Despliegue en VPS Online (Paso a Paso)", level=1)
    add_apa_paragraph(
        "A continuación se especifica el procedimiento estandarizado para desplegar SEOTracker en un Servidor Privado Virtual (VPS) "
        "con sistema operativo Ubuntu 22.04 LTS o 24.04 LTS, utilizando Nginx como Proxy Inverso, Gunicorn como servidor de aplicaciones WSGI, "
        "MySQL Server y certificados SSL emitidos por Let's Encrypt."
    )

    add_apa_heading("8.1. Requisitos del Servidor", level=2)
    add_apa_paragraph("• Sistema Operativo: Ubuntu 22.04 LTS o 24.04 LTS limpio.")
    add_apa_paragraph("• Hardware Mínimo: 1 CPU, 1 GB RAM, 25 GB SSD.")
    add_apa_paragraph("• Registro DNS: Dominio o subdominio apuntando mediante registro A a la IP pública del VPS (ej. seo.tudominio.com).")

    add_apa_heading("8.2. Instalación de Paquetes del Sistema", level=2)
    add_code_block(
        "sudo apt update && sudo apt upgrade -y\n"
        "sudo apt install -y python3-pip python3-venv python3-dev mysql-server nginx git curl certbot python3-certbot-nginx",
        lang="bash"
    )

    add_apa_heading("8.3. Configuración de Base de Datos MySQL", level=2)
    add_code_block(
        "sudo mysql\n"
        "CREATE DATABASE seo_auditor DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;\n"
        "CREATE USER 'seo_user'@'localhost' IDENTIFIED BY 'TuPasswordSeguro2026!';\n"
        "GRANT ALL PRIVILEGES ON seo_auditor.* TO 'seo_user'@'localhost';\n"
        "FLUSH PRIVILEGES;\n"
        "EXIT;",
        lang="sql"
    )

    add_apa_heading("8.4. Despliegue del Código y Entorno Virtual", level=2)
    add_code_block(
        "sudo mkdir -p /var/www/seotracker\n"
        "sudo chown -R $USER:$USER /var/www/seotracker\n"
        "cd /var/www/seotracker\n"
        "git clone <URL_DEL_REPOSITORIO> .\n"
        "python3 -m venv venv\n"
        "source venv/bin/activate\n"
        "pip install --upgrade pip\n"
        "pip install -r requirements.txt\n"
        "pip install gunicorn",
        lang="bash"
    )

    add_apa_heading("8.5. Configuración de Variables de Entorno (.env)", level=2)
    add_apa_paragraph("Cree el archivo /var/www/seotracker/.env e ingrese los parámetros de producción:")
    add_code_block(
        "DB_HOST=127.0.0.1\n"
        "DB_USER=seo_user\n"
        "DB_PASS=TuPasswordSeguro2026!\n"
        "DB_NAME=seo_auditor\n"
        "SMTP_EMAIL=fabianperez@estrategiamic.com\n"
        "SMTP_PASS=tu_app_password_gmail\n"
        "SMTP_SERVER=smtp.gmail.com\n"
        "SMTP_PORT=587\n"
        "SECRET_KEY=clave_ultra_secreta_para_flask_2026",
        lang="ini"
    )

    add_apa_heading("8.6. Servicio Systemd para Gunicorn", level=2)
    add_apa_paragraph("Cree el archivo de servicio de sistema /etc/systemd/system/seotracker.service:")
    add_code_block(
        "[Unit]\n"
        "Description=Servidor SEOTracker Gunicorn\n"
        "After=network.target mysql.service\n\n"
        "[Service]\n"
        "User=root\n"
        "Group=www-data\n"
        "WorkingDirectory=/var/www/seotracker\n"
        "Environment=\"PATH=/var/www/seotracker/venv/bin\"\n"
        "EnvironmentFile=/var/www/seotracker/.env\n"
        "ExecStart=/var/www/seotracker/venv/bin/gunicorn --workers 3 --bind 127.0.0.1:5000 app:app\n\n"
        "[Install]\n"
        "WantedBy=multi-user.target",
        lang="ini"
    )
    add_code_block(
        "sudo systemctl daemon-reload\n"
        "sudo systemctl start seotracker\n"
        "sudo systemctl enable seotracker",
        lang="bash"
    )

    add_apa_heading("8.7. Configuración de Proxy Inverso Nginx y SSL", level=2)
    add_apa_paragraph("Cree el bloque de sitio /etc/nginx/sites-available/seotracker:")
    add_code_block(
        "server {\n"
        "    listen 80;\n"
        "    server_name seo.tudominio.com;\n\n"
        "    client_max_body_size 50M;\n\n"
        "    location / {\n"
        "        proxy_pass http://127.0.0.1:5000;\n"
        "        proxy_set_header Host $host;\n"
        "        proxy_set_header X-Real-IP $remote_addr;\n"
        "        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;\n"
        "        proxy_set_header X-Forwarded-Proto $scheme;\n"
        "    }\n\n"
        "    location /static/ {\n"
        "        alias /var/www/seotracker/static/;\n"
        "        expires 30d;\n"
        "    }\n"
        "}",
        lang="nginx"
    )
    add_code_block(
        "sudo ln -s /etc/nginx/sites-available/seotracker /etc/nginx/sites-enabled/\n"
        "sudo nginx -t\n"
        "sudo systemctl restart nginx\n"
        "sudo certbot --nginx -d seo.tudominio.com",
        lang="bash"
    )

    add_apa_heading("8.8. Monitoreo y Mantenimiento", level=2)
    add_code_block(
        "# Ver logs del servicio Gunicorn:\n"
        "sudo journalctl -u seotracker -f\n\n"
        "# Ver logs de Nginx:\n"
        "sudo tail -f /var/log/nginx/error.log\n\n"
        "# Actualizar aplicación y reiniciar:\n"
        "cd /var/www/seotracker && git pull && sudo systemctl restart seotracker",
        lang="bash"
    )

    # 11. Sección 9: Historial de Cambios
    add_apa_heading("9. Historial de Cambios", level=1)
    add_apa_paragraph(
        "Esta sección registra las actualizaciones y mejoras aplicadas al sistema SEOTracker, "
        "ordenadas de forma cronológica descendente."
    )

    changelog_entries = [
        (
            "v1.4 — Agosto 2026",
            "Soporte para Headers Personalizados y Rol de Empresa",
            [
                "Base de Datos: Se añadió 'headers_personalizados' (TEXT), 'fecha_headers' (DATETIME), y 'alerta_headers_enviada' (BOOLEAN) a la tabla 'marcas'. Se añadió 'rol_empresa' a la tabla 'usuarios'.",
                "Frontend e Interfaces: Se actualizaron los formularios y endpoints para capturar, guardar y mostrar el rol_empresa de cada usuario y los headers_personalizados de cada marca.",
                "Archivos modificados: setup_db.py · auditor.py · app.py · templates/marca_form.html · templates/usuario_form.html · templates/usuarios.html"
            ]
        ),
        (
            "v1.3 — 22 de julio de 2026",
            "Resultados Duales en Modo Lista para Marcas No Responsive",
            [
                "Pantalla: Al completar el rastreo de una marca configurada como No Responsive, el sistema muestra "
                "automáticamente ambas tablas (Mobile y Desktop) apiladas verticalmente. Se añadió una tercera "
                "pestaña 'Ambos' (activa por defecto), además de 'Solo Mobile' y 'Solo Desktop'.",
                "Correo: El botón Enviar Correo genera y adjunta dos archivos Excel (Mobile y Desktop) en un único "
                "correo consolidado cuando la marca es No Responsive.",
                "Archivos modificados: app.py · auditor.py · templates/rastreo_lista.html"
            ]
        ),
        (
            "v1.2 — 22 de julio de 2026",
            "Permisos de Creación de Marcas para el Rol Project Manager",
            [
                "Los usuarios con rol Project Manager (pm) ahora pueden crear y editar marcas desde la "
                "sección Marcas, en igualdad de condiciones con el rol Líder.",
                "Archivos modificados: app.py · templates/marcas.html · templates/dashboard.html"
            ]
        ),
        (
            "v1.1 — 22 de julio de 2026",
            "Documentación Inicial del Proyecto",
            [
                "Creación del documento DOCUMENTACION_PROYECTO.md con arquitectura, manual de usuario, "
                "guías de instalación local y despliegue en VPS.",
                "Generación del documento en formato Word con Normas APA 7.ª edición (DOCUMENTACION_PROYECTO_APA.docx)."
            ]
        ),
    ]

    for version, title, items in changelog_entries:
        p_ver = doc.add_paragraph()
        p_ver.paragraph_format.space_before = Pt(10)
        p_ver.paragraph_format.space_after = Pt(2)
        r_ver = p_ver.add_run(version)
        r_ver.font.name = 'Times New Roman'
        r_ver.font.bold = True
        r_ver.font.size = Pt(12)
        r_ver.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

        p_title = doc.add_paragraph()
        p_title.paragraph_format.space_before = Pt(0)
        p_title.paragraph_format.space_after = Pt(4)
        r_title = p_title.add_run(title)
        r_title.font.name = 'Times New Roman'
        r_title.font.bold = True
        r_title.font.italic = True
        r_title.font.size = Pt(11)

        for item in items:
            add_apa_paragraph(item, bullet=True)

    # Save document
    doc.save(output_docx)
    print(f"Documento generado exitosamente en: {output_docx}")

if __name__ == '__main__':
    md_file = r"c:\Users\fabianperez\Desktop\SEOTracker\DOCUMENTACION_PROYECTO.md"
    out_file = r"c:\Users\fabianperez\Desktop\SEOTracker\DOCUMENTACION_PROYECTO_APA.docx"
    create_apa_document(md_file, out_file)

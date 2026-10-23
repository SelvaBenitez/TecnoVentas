import openpyxl
from tabulate import tabulate
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

ENCABEZADOS = ["ID", "Producto", "Categoría", "Precio ($)", "Ventas (Unid)"]
ALINEACION = ("center", "left", "center", "right", "right")


def cargar_datos_excel(nombre_archivo="ventas_tecnologia_v2.xlsx"):
    """Carga los datos del Excel omitiendo la primera fila de títulos."""
    try:
        wb = openpyxl.load_workbook(nombre_archivo, data_only=True)
        hoja = wb.active
        datos = []
        for fila in hoja.iter_rows(min_row=2, values_only=True):
            if fila[0] is not None:
                datos.append(list(fila))
        return datos
    except FileNotFoundError:
        print(f"\nNo se encontró el archivo '{nombre_archivo}'.")
        return []


# OPCIÓN 1
def ver_catalogo(datos):
    print(f"\n--- CATÁLOGO COMPLETO ({len(datos)} PRODUCTOS) ---")
    print(tabulate(datos, headers=ENCABEZADOS, tablefmt="fancy_grid", colalign=ALINEACION))


# OPCIÓN 2
def filtrar_por_precio(datos, limite=100):
    datos_filtrados = [fila for fila in datos if fila[3] > limite]
    print(f"\n--- REPORTE DE VENTAS EN CONSOLA (PRECIO > ${limite}) ---")
    print(f"Total de productos filtrados: {len(datos_filtrados)}")
    print(tabulate(datos_filtrados, headers=ENCABEZADOS, tablefmt="fancy_grid", colalign=ALINEACION))


# OPCIÓN 3
def buscar_por_id(datos):
    print("\n--- BÚSQUEDA DE PRODUCTO POR ID ---")
    id_buscar = input("Ingrese el ID del producto que desea buscar: ").strip()

    producto_encontrado = []
    for fila in datos:
        if str(fila[0]).strip().lower() == id_buscar.lower():
            producto_encontrado.append(fila)
            break

    if producto_encontrado:
        print(f"\nProducto encontrado con ID: {id_buscar}")
        print(tabulate(producto_encontrado, headers=ENCABEZADOS, tablefmt="fancy_grid", colalign=ALINEACION))
    else:
        print(f"\n[ADVERTENCIA] No se encontró ningún producto con el ID '{id_buscar}'.")


# OPCIÓN 4
def exportar_informe_pdf(datos, nombre_pdf="reporte_ejecutivo_ventas.pdf"):
    if not datos:
        print("\n[ERROR] No hay datos cargados en memoria.")
        return

    # Reto 2
    total_unidades = sum(fila[4] for fila in datos)
    precio_promedio = sum(fila[3] for fila in datos) / len(datos)

    print("\n--- MÉTRICAS GENERALES EN CONSOLA ---")
    print(f"Total de productos: {len(datos)}")
    print(f"Total de unidades vendidas: {total_unidades}")
    print(f"Precio promedio por producto: ${precio_promedio:.2f}")

    print("\nGenerando reporte .pdf ...")
    doc = SimpleDocTemplate(nombre_pdf, pagesize=letter)
    estilos = getSampleStyleSheet()

    estilo_titulo = ParagraphStyle('EstiloTitulo', parent=estilos['Title'], fontName='Helvetica-Bold', fontSize=18, textColor=colors.HexColor("#1A365D"), spaceAfter=10)
    estilo_subtitulo = ParagraphStyle('EstiloSubtitulo', parent=estilos['Heading2'], fontName='Helvetica-Bold', fontSize=12, textColor=colors.HexColor("#1A365D"), spaceAfter=6)
    estilo_cuerpo = ParagraphStyle('EstiloCuerpo', parent=estilos['Normal'], fontSize=9)
    estilo_cuerpo_c = ParagraphStyle('EstiloCuerpoC', parent=estilo_cuerpo, alignment=1)
    estilo_encabezado = ParagraphStyle('EstiloEncabezado', parent=estilos['Normal'], fontName='Helvetica-Bold', fontSize=10, textColor=colors.whitesmoke, alignment=1)
    estilo_resumen = ParagraphStyle('EstiloResumen', parent=estilo_cuerpo, fontName='Helvetica-Bold')
    estilo_resumen_derecha = ParagraphStyle('EstiloResumenR', parent=estilo_resumen, alignment=2)

    story = [
        Paragraph("Informe Ejecutivo de Ventas Tecno", estilo_titulo),
        Spacer(1, 8)
    ]

    # RETO 2
    tabla_metricas_data = [
        [Paragraph("<b>Métrica General</b>", estilo_encabezado), Paragraph("<b>Valor Calculado</b>", estilo_encabezado)],
        [Paragraph("Total de Productos Procesados", estilo_cuerpo), Paragraph(f"<b>{len(datos)}</b>", estilo_cuerpo_c)],
        [Paragraph("Total de Unidades Vendidas", estilo_cuerpo), Paragraph(f"<b>{total_unidades}</b>", estilo_cuerpo_c)],
        [Paragraph("Precio Promedio por Producto", estilo_cuerpo), Paragraph(f"<b>${precio_promedio:.2f}</b>", estilo_cuerpo_c)]
    ]
    tabla_metricas = Table(tabla_metricas_data, colWidths=[200, 150])
    tabla_metricas.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1A365D")),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor("#F7FAFC")),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))

    story.append(tabla_metricas)
    story.append(Spacer(1, 15))
    story.append(Paragraph("Detalle General de Productos", estilo_subtitulo))

    encabezados_pdf = [Paragraph(f"<b>{h}</b>", estilo_encabezado) for h in ENCABEZADOS]

    filas_pdf = []
    for r in datos:
        filas_pdf.append([
            Paragraph(str(r[0]), ParagraphStyle('C', parent=estilo_cuerpo, alignment=1)),
            Paragraph(str(r[1]), ParagraphStyle('L', parent=estilo_cuerpo, alignment=0)),
            Paragraph(str(r[2]), ParagraphStyle('C', parent=estilo_cuerpo, alignment=1)),
            Paragraph(f"${r[3]:.2f}", ParagraphStyle('R', parent=estilo_cuerpo, alignment=2)),
            Paragraph(str(r[4]), ParagraphStyle('R', parent=estilo_cuerpo, alignment=2))
        ])

    # RETO 2
    fila_resumen = [
        Paragraph("<b>TOTALES / PROMEDIO</b>", estilo_resumen),
        Paragraph("", estilo_cuerpo),
        Paragraph("", estilo_cuerpo),
        Paragraph(f"<b>${precio_promedio:.2f}</b>", estilo_resumen_derecha),
        Paragraph(f"<b>{total_unidades}</b>", estilo_resumen_derecha)
    ]
    filas_pdf.append(fila_resumen)

    anchos_cols = [40, 160, 100, 80, 80]
    tabla_pdf = Table([encabezados_pdf] + filas_pdf, colWidths=anchos_cols, repeatRows=1)
    tabla_pdf.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1A365D")),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
        ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor("#F7FAFC")]),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#E2E8F0")),
    ]))

    story.append(tabla_pdf)
    doc.build(story)
    print(f"\nArchivo '{nombre_pdf}' generado correctamente.")
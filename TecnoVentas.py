import openpyxl
from tabulate import tabulate

'''
Importaciones para realizar el reporte
'''
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

wb=openpyxl.load_workbook("ventas_tecnologia_v2.xlsx", data_only=True)
hoja=wb.active
print(f"Hoja cargada exitosamente: {hoja.title}")

datos = []
for fila in hoja.iter_rows(min_row=2, values_only=True):
    datos.append(list(fila))

# Reto 1   
datos_filtrados = []
for fila in datos:
    if fila[3] >100:
        datos_filtrados.append(fila)

print(f"Total de productos filtrados (> $100): {len(datos_filtrados)}")
print("\n--- REPORTE DE VENTAS EN CONSOLA (PRECIO > $100) ---")


# print(f'Total de registros leídos en memoria: {len(datos)}')


print("Generando reporte .pdf ...")
encabezados = ["ID", "Producto", "Categoría", "Precio ($)", "Ventas (Unid)"]
alineacion = ("center", "left", "center", "right", "right")
print("\n--- REPORTE DE VENTAS EN CONSOLA ---")
print("\n--- REPORTE DE VENTAS EN CONSOLA (PRECIO > $100) ---")
print(tabulate(datos_filtrados, headers=encabezados, tablefmt="fancy_grid", colalign=alineacion))#reto 1
# print(tabulate(datos, headers=encabezados, tablefmt="fancy_grid", colalign=alineacion))

# reto 2
total_unidades = sum(fila[4] for fila in datos)
precio_promedio = sum(fila[3] for fila in datos) / len(datos)

print("\n--- MÉTRICAS GENERALES ---")
print(f"Total de productos: {len(datos)}")
print(f"Total de unidades vendidas: {total_unidades}")
print(f"Precio promedio por producto: ${precio_promedio:.2f}")


#Formato impresión pdf
#crear hoja
doc = SimpleDocTemplate("reporte_ejecutivo_ventas.pdf", pagesize=letter)
estilos = getSampleStyleSheet()

#estilo textos
estilo_titulo = ParagraphStyle(
    'EstiloTitulo', parent=estilos['Title'],
    fontName='Helvetica-Bold', fontSize=18,
    textColor=colors.HexColor("#1A365D"), spaceAfter=10
)
estilo_cuerpo = ParagraphStyle('EstiloCuerpo', parent=estilos['Normal'], fontSize=9)
estilo_encabezado = ParagraphStyle(
    'EstiloEncabezado', parent=estilos['Normal'],
    fontName='Helvetica-Bold', fontSize=10,
    textColor=colors.whitesmoke, alignment=1
)

#Lista story 
story = [
    Paragraph("Informe Ejecutivo de Ventas Tecno", estilo_titulo),
    Paragraph(f"<b>Total de Productos Procesados:</b> {len(datos)}", estilo_cuerpo),
    Spacer(1, 12)
]

#Convertir los encabezados y celdas a objetos Paragraph
encabezados_pdf = [Paragraph(f"<b>{h}</b>", estilo_encabezado) for h in encabezados]

filas_pdf = []

for r in datos_filtrados:
    filas_pdf.append([
        Paragraph(str(r[0]), ParagraphStyle('C', parent=estilo_cuerpo, alignment=1)),
        Paragraph(str(r[1]), ParagraphStyle('L', parent=estilo_cuerpo, alignment=0)),
        Paragraph(str(r[2]), ParagraphStyle('C', parent=estilo_cuerpo, alignment=1)),
        Paragraph(f"${r[3]:.2f}", ParagraphStyle('R', parent=estilo_cuerpo, alignment=2)),
        Paragraph(str(r[4]), ParagraphStyle('R', parent=estilo_cuerpo, alignment=2))
    ])

# for r in datos:
#     filas_pdf.append([
#         Paragraph(str(r[0]), ParagraphStyle('C', parent=estilo_cuerpo, alignment=1)),
#         Paragraph(str(r[1]), ParagraphStyle('L', parent=estilo_cuerpo, alignment=0)),
#         Paragraph(str(r[2]), ParagraphStyle('C', parent=estilo_cuerpo, alignment=1)),
#         Paragraph(f"${r[3]:.2f}", ParagraphStyle('R', parent=estilo_cuerpo, alignment=2)),
#         Paragraph(str(r[4]), ParagraphStyle('R', parent=estilo_cuerpo, alignment=2))
#     ])

#Configurar anchos de columna y crear la tabla
anchos_cols = [40, 160, 100, 80, 80]
tabla_pdf = Table([encabezados_pdf] + filas_pdf, colWidths=anchos_cols, repeatRows=1)

#Dar estilo visual a la tabla
tabla_pdf.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1A365D")),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7FAFC")])
]))

#Agregar la tabla al carrito y compilar el PDF
story.append(tabla_pdf)
doc.build(story)

print("\n[ÉXITO] Archivo 'reporte_ejecutivo_ventas.pdf' generado correctamente.")

#reto2
estilo_resumen = ParagraphStyle('EstiloResumen', parent=estilo_cuerpo, fontName='Helvetica-Bold')
estilo_resumen_derecha = ParagraphStyle('EstiloResumenR', parent=estilo_resumen, alignment=2)

fila_resumen = [
    Paragraph("<b>TOTALES / PROMEDIO</b>", estilo_resumen),
    Paragraph("", estilo_cuerpo),
    Paragraph("", estilo_cuerpo),
    Paragraph(f"<b>${precio_promedio:.2f}</b>", estilo_resumen_derecha),
    Paragraph(f"<b>{total_unidades}</b>", estilo_resumen_derecha)
]

filas_pdf.append(fila_resumen)

#reto 3
print("\n--- BÚSQUEDA DE PRODUCTO POR ID ---")
id_buscar = input("Ingrese el ID del producto que desea buscar: ").strip()

producto_encontrado = []
for fila in datos:
    if str(fila[0]).strip() == id_buscar:
        producto_encontrado.append(fila)
        break 

if producto_encontrado:
    print(f"\n: Producto con ID= {id_buscar}:")
    print(tabulate(producto_encontrado, headers=encabezados, tablefmt="fancy_grid", colalign=alineacion))
else:
    print(f"\n[ADVERTENCIA] No se encontró ningún producto con el ID '{id_buscar}'.")
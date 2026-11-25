import os
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


def generar_excel_entrada(registro, carpeta_salida='exports'):
  """Genera un archivo .xlsx con los datos de una nota de entrada."""
  os.makedirs(carpeta_salida, exist_ok=True)
  nombre_archivo = f"Nota_Entrada_{registro['folio']}_{registro['id']}.xlsx"
  ruta_archivo = os.path.join(carpeta_salida, nombre_archivo)

  wb = Workbook()
  ws = wb.active
  ws.title = 'NOTA DE ENTRADA'

  font_titulo = Font(name='Arial', size=13, bold=True, color='FFFFFF')
  font_subtitulo = Font(name='Arial', size=10, bold=True, color='1E293B')
  font_label = Font(name='Arial', size=10, bold=True)
  font_normal = Font(name='Arial', size=10)

  fill_encabezado = PatternFill(
      start_color='1E3A8A', end_color='1E3A8A', fill_type='solid'
  )
  fill_tabla_header = PatternFill(
      start_color='E2E8F0', end_color='E2E8F0', fill_type='solid'
  )

  thin_side = Side(border_style='thin', color='94A3B8')
  borde_celda = Border(
      left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
  )

  # Encabezado institucional
  ws.merge_cells('A1:F1')
  ws['A1'] = 'COMISIÓN ESTATAL DE SUELO Y VIVIENDA (COESVI)'
  ws['A1'].font = font_titulo
  ws['A1'].fill = fill_encabezado
  ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
  ws.row_dimensions[1].height = 26

  ws.merge_cells('A2:F2')
  ws['A2'] = 'PARQUE DE MATERIALES — NOTA DE ENTRADA (FOR-CSV-DAD-PQM-01)'
  ws['A2'].font = font_subtitulo
  ws['A2'].alignment = Alignment(horizontal='center', vertical='center')
  ws.row_dimensions[2].height = 20

  # Datos generales
  datos_generales = [
      ('FOLIO:', registro['folio'], 'FECHA:', registro['fecha']),
      (
          'PROVEEDOR:',
          registro['proveedor'] or 'Producción Interna',
          'FACTURA No.:',
          registro['factura_no'] or '-',
      ),
      ('REMISIÓN No.:', registro['remision_no'] or '-', '', ''),
  ]

  fila_actual = 4
  for lbl1, val1, lbl2, val2 in datos_generales:
    ws[f'A{fila_actual}'] = lbl1
    ws[f'A{fila_actual}'].font = font_label if lbl1 else font_normal
    ws.merge_cells(f'B{fila_actual}:C{fila_actual}')
    ws[f'B{fila_actual}'] = val1
    ws[f'B{fila_actual}'].font = font_normal

    ws[f'D{fila_actual}'] = lbl2
    ws[f'D{fila_actual}'].font = font_label if lbl2 else font_normal
    ws.merge_cells(f'E{fila_actual}:F{fila_actual}')
    ws[f'E{fila_actual}'] = val2
    ws[f'E{fila_actual}'].font = font_normal

    for col in ['A', 'B', 'C', 'D', 'E', 'F']:
      ws[f'{col}{fila_actual}'].border = borde_celda
    fila_actual += 1

  fila_actual += 1

  # Encabezado de detalle de material
  headers = [
      'CANTIDAD',
      'UNIDAD',
      'DESCRIPCIÓN',
      'FOLIO ENTRADA',
      'FECHA',
      'IMPORTE ($)',
  ]
  for idx, h in enumerate(headers, start=1):
    celda = ws.cell(row=fila_actual, column=idx, value=h)
    celda.font = font_label
    celda.fill = fill_tabla_header
    celda.border = borde_celda
    celda.alignment = Alignment(horizontal='center', vertical='center')

  fila_detalle = fila_actual + 1
  importe = float(registro['importe'] or 0.0)
  valores_detalle = [
      float(registro['cantidad']),
      registro['unidad'],
      registro['descripcion'],
      registro['folio'],
      registro['fecha'],
      importe,
  ]
  for idx, val in enumerate(valores_detalle, start=1):
    celda = ws.cell(row=fila_detalle, column=idx, value=val)
    celda.font = font_normal
    celda.border = borde_celda
    if idx in (1, 6):
      celda.alignment = Alignment(horizontal='right', vertical='center')
      celda.number_format = '#,##0.00'
    else:
      celda.alignment = Alignment(horizontal='center', vertical='center')

  anchos = {'A': 16, 'B': 16, 'C': 34, 'D': 18, 'E': 16, 'F': 16}
  for letra, ancho in anchos.items():
    ws.column_dimensions[letra].width = ancho

  wb.save(ruta_archivo)
  return ruta_archivo


def generar_excel_salida(registro, carpeta_salida='exports'):
  """Genera un archivo .xlsx con los datos de una nota de salida."""
  os.makedirs(carpeta_salida, exist_ok=True)
  nombre_archivo = f"Nota_Salida_{registro['folio']}_{registro['id']}.xlsx"
  ruta_archivo = os.path.join(carpeta_salida, nombre_archivo)

  wb = Workbook()
  ws = wb.active
  ws.title = 'NOTA DE SALIDA'

  font_titulo = Font(name='Arial', size=13, bold=True, color='FFFFFF')
  font_subtitulo = Font(name='Arial', size=10, bold=True, color='1E293B')
  font_label = Font(name='Arial', size=10, bold=True)
  font_normal = Font(name='Arial', size=10)

  fill_encabezado = PatternFill(
      start_color='92400E', end_color='92400E', fill_type='solid'
  )
  fill_tabla_header = PatternFill(
      start_color='FEF3C7', end_color='FEF3C7', fill_type='solid'
  )

  thin_side = Side(border_style='thin', color='94A3B8')
  borde_celda = Border(
      left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
  )

  # Encabezado institucional
  ws.merge_cells('A1:F1')
  ws['A1'] = 'COMISIÓN ESTATAL DE SUELO Y VIVIENDA (COESVI)'
  ws['A1'].font = font_titulo
  ws['A1'].fill = fill_encabezado
  ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
  ws.row_dimensions[1].height = 26

  ws.merge_cells('A2:F2')
  ws['A2'] = 'PARQUE DE MATERIALES — NOTA DE SALIDA (FOR-CSV-DAD-PQM-02)'
  ws['A2'].font = font_subtitulo
  ws['A2'].alignment = Alignment(horizontal='center', vertical='center')
  ws.row_dimensions[2].height = 20

  # Datos generales
  datos_generales = [
      ('FOLIO:', registro['folio'], 'FECHA:', registro['fecha']),
      (
          'ENTREGADO A:',
          registro['beneficiario'],
          'RECIBE EN:',
          registro['recibe_en'] or '-',
      ),
  ]

  fila_actual = 4
  for lbl1, val1, lbl2, val2 in datos_generales:
    ws[f'A{fila_actual}'] = lbl1
    ws[f'A{fila_actual}'].font = font_label
    ws.merge_cells(f'B{fila_actual}:C{fila_actual}')
    ws[f'B{fila_actual}'] = val1
    ws[f'B{fila_actual}'].font = font_normal

    ws[f'D{fila_actual}'] = lbl2
    ws[f'D{fila_actual}'].font = font_label
    ws.merge_cells(f'E{fila_actual}:F{fila_actual}')
    ws[f'E{fila_actual}'] = val2
    ws[f'E{fila_actual}'].font = font_normal

    for col in ['A', 'B', 'C', 'D', 'E', 'F']:
      ws[f'{col}{fila_actual}'].border = borde_celda
    fila_actual += 1

  fila_actual += 1

  # Encabezado de detalle de material
  headers = [
      'CANTIDAD',
      'UNIDAD',
      'DESCRIPCIÓN',
      'FOLIO SALIDA',
      'FECHA',
      'RECIBE EN',
  ]
  for idx, h in enumerate(headers, start=1):
    celda = ws.cell(row=fila_actual, column=idx, value=h)
    celda.font = font_label
    celda.fill = fill_tabla_header
    celda.border = borde_celda
    celda.alignment = Alignment(horizontal='center', vertical='center')

  fila_detalle = fila_actual + 1
  valores_detalle = [
      float(registro['cantidad']),
      registro['unidad'],
      registro['descripcion'],
      registro['folio'],
      registro['fecha'],
      registro['recibe_en'] or '-',
  ]
  for idx, val in enumerate(valores_detalle, start=1):
    celda = ws.cell(row=fila_detalle, column=idx, value=val)
    celda.font = font_normal
    celda.border = borde_celda
    if idx == 1:
      celda.alignment = Alignment(horizontal='right', vertical='center')
      celda.number_format = '#,##0.00'
    else:
      celda.alignment = Alignment(horizontal='center', vertical='center')

  anchos = {'A': 16, 'B': 18, 'C': 34, 'D': 18, 'E': 16, 'F': 26}
  for letra, ancho in anchos.items():
    ws.column_dimensions[letra].width = ancho

  wb.save(ruta_archivo)
  return ruta_archivo
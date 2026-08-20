"""Aplica formato profesional al archivo resumen_ventas.xlsx."""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

RUTA = "output/resumen_ventas.xlsx"

# Paleta corporativa (navy + acento)
NAVY = "1F3864"
BLANCO = "FFFFFF"
GRIS_CLARO = "F2F2F2"

fuente_header = Font(name="Arial", size=11, bold=True, color=BLANCO)
fuente_cuerpo = Font(name="Arial", size=10)
fill_header = PatternFill("solid", fgColor=NAVY)
fill_zebra = PatternFill("solid", fgColor=GRIS_CLARO)
centro = Alignment(horizontal="center", vertical="center")
izq = Alignment(horizontal="left", vertical="center")
der = Alignment(horizontal="right", vertical="center")
borde = Border(*[Side(style="thin", color="BFBFBF")] * 4)

wb = load_workbook(RUTA)

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    max_row = ws.max_row
    max_col = ws.max_column

    # Encabezados
    for col in range(1, max_col + 1):
        c = ws.cell(row=1, column=col)
        c.font = fuente_header
        c.fill = fill_header
        c.alignment = centro
        c.border = borde

    # Cuerpo
    for row in range(2, max_row + 1):
        for col in range(1, max_col + 1):
            c = ws.cell(row=row, column=col)
            c.font = fuente_cuerpo
            c.border = borde
            header = ws.cell(row=1, column=col).value
            if header and ("Total" in str(header)):
                c.number_format = '"$"#,##0'
                c.alignment = der
            elif header == "Mes":
                c.alignment = centro
            else:
                c.alignment = izq
            if row % 2 == 0:
                c.fill = fill_zebra

    # Anchos de columna
    for col in range(1, max_col + 1):
        letra = get_column_letter(col)
        max_len = max(
            (len(str(ws.cell(row=r, column=col).value)) for r in range(1, max_row + 1)
             if ws.cell(row=r, column=col).value is not None),
            default=10,
        )
        ws.column_dimensions[letra].width = max(14, min(max_len + 4, 30))

    # Freeze header
    ws.freeze_panes = "A2"

wb.save(RUTA)
print(f"[OK] Formato aplicado a {RUTA}")

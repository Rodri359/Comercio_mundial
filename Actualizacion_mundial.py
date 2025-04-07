import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Alignment

def clear_cell_format(ws, start_row, start_col, nrows, ncols):
    """
    Limpia el formato en el rango especificado dejando solo los valores.
    start_row y start_col son índices 1-basados.
    """
    for row in ws.iter_rows(min_row=start_row, 
                            max_row=start_row + nrows - 1, 
                            min_col=start_col, 
                            max_col=start_col + ncols - 1):
        for cell in row:
            cell.font = Font()  # fuente por defecto
            cell.fill = PatternFill()  # sin relleno
            cell.border = Border()  # sin bordes
            cell.alignment = Alignment()  # alineación por defecto
            cell.number_format = 'General'
            
# Configuración de rutas
source_dir = r'C:/Users/rodri/Downloads/Datos_Extraidos'
template_path = r'C:/Users/rodri/Downloads/estadisticas_macro_shared/estadisticas_macro_shared/Plantillas/Mercado mundial -  Plantilla A.xlsx'
output_dir = r'C:/Users/rodri/Downloads/estadisticas_macro_shared/estadisticas_macro_shared/Resultado'

if not os.path.exists(template_path):
    raise FileNotFoundError(f"La plantilla no fue encontrada en: {template_path}")

os.makedirs(output_dir, exist_ok=True)
files = [f for f in os.listdir(source_dir) if f.endswith('.xlsx')]

for file in files:
    source_file = os.path.join(source_dir, file)
    
    datos_archivo = pd.read_excel(source_file, sheet_name=None, skiprows=1)
    
    base_name = os.path.splitext(file)[0]
    output_file = os.path.join(output_dir, f"{base_name}_actualizado.xlsx")
    
    book = load_workbook(template_path)
    writer = pd.ExcelWriter(output_file, engine='openpyxl')
    writer._book = book
    writer._sheets = {ws.title: ws for ws in book.worksheets}
    
    for sheet in list(book.sheetnames):
        if sheet != "(Paises)" and sheet in datos_archivo:
            df = datos_archivo[sheet]
            df.to_excel(writer, sheet_name=sheet, index=False, startrow=11, startcol=1)
            ws = book[sheet]
            nrows, ncols = df.shape
            clear_cell_format(ws, start_row=11, start_col=2, nrows=nrows, ncols=ncols)
    
    # Si existe la hoja "(Paises)", usarla como plantilla para hojas extra (por ejemplo, países)
    if "(Paises)" in book.sheetnames:
        for sheet in datos_archivo.keys():
            if sheet not in book.sheetnames:
                new_sheet = book.copy_worksheet(book["(Paises)"])
                new_sheet.title = sheet  
                writer._sheets[new_sheet.title] = new_sheet
                df = datos_archivo[sheet]
                df.to_excel(writer, sheet_name=new_sheet.title, index=False, startrow=11, startcol=1)
                clear_cell_format(new_sheet, start_row=11, start_col=2, nrows=df.shape[0], ncols=df.shape[1])
        paises_sheet = book["(Paises)"]
        book.remove(paises_sheet)
    
    writer.close()
    print(f"Procesado: {output_file}")
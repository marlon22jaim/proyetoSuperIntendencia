import os
import zipfile
import shutil
from datetime import datetime
import re
import pandas as pd

# Obtener lista de archivos zip
zip_files = [f for f in os.listdir('.') if f.endswith('.zip')]

# Expresión regular para match YYYY_MM_DD
date_regex = r'(\d{4})_(\d{2})_(\d{2})'

# Ordenar con expresión regular para extraer fecha
zip_files.sort(key=lambda x: datetime.strptime(re.search(date_regex, x).group(), '%Y_%m_%d'), reverse=True)

# Tomar el archivo más reciente
latest_zip = zip_files[0]

# Descomprimir el zip
with zipfile.ZipFile(latest_zip, 'r') as zip_ref:
    zip_ref.extractall()

# Obtener el nombre de la carpeta descomprimida
extracted_dir = latest_zip.replace('.zip', '')

# Buscar los archivos Excel en la carpeta
excel_files = [f for f in os.listdir(extracted_dir) if f.endswith('.xlsx')]

# Ruta de salida para los archivos generados
output_path = r'D:\PUBLICO\Proyecto HCI\Centro_Monitoreo\Bases de Datos\BVC_Boletines_Diario'

for excel_file in excel_files:
    # Leer archivo Excel con múltiples hojas en un diccionario de DataFrames
    dfs = pd.read_excel(os.path.join(extracted_dir, excel_file), sheet_name=None)

    # Recorrer las hojas y DataFrames correspondientes
    for sheet_name, cols in zip(
        ['RV-Cap. Bursátil', 'RV-Ventas en Corto', 'RF-Mercado Primario'],
        [dfs['RV-Cap. Bursátil'].iloc[3:, 2:6], dfs['RV-Ventas en Corto'].iloc[3:, 2:8], dfs['RF-Mercado Primario'].iloc[3:, 2:13]]
    ):
        # Asignar nombres de columnas
        cols.columns = dfs[sheet_name].iloc[3, 2:cols.shape[1]+2]

        # Eliminar la segunda fila
        cols = cols.iloc[1:]

        # Generar nombre archivo de salida
        now = datetime.now()
        date_time = now.strftime("%Y-%m-%d-%H-%M-%S")
        date_regex = r'(\d{4}-\d{2}-\d{2})'
        date = re.search(date_regex, excel_file).group(1)
        out_name = f'{sheet_name} {date}-{now.strftime("%H-%M-%S")}.xlsx'

        # Crear ExcelWriter con ruta completa
        writer = pd.ExcelWriter(os.path.join(output_path, out_name), engine='xlsxwriter')

        # Escribir DataFrame en hoja 'Hoja1' del libro nuevo sin índice
        cols.to_excel(writer, sheet_name='Hoja1', index=False)

        # Guardar archivo Excel
        writer.close()

# Eliminar la carpeta descomprimida
shutil.rmtree(extracted_dir)

print('Archivos Excel extraídos y carpetas eliminadas')

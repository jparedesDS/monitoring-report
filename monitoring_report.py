# Imports
import os
import time
import pandas as pd
import xlsxwriter
import jpype
import asposecells
jpype.startJVM()
from asposecells.api import Workbook
from openpyxl.reader.excel import load_workbook
from tools.mapping_mr import *
from tools.apply_style_mr import *


start_time = time.time()
# Ruta del archivo Excel existente
path_pending = 'C:\\Users\\alejandro.berzal\\Desktop\\DATA SCIENCE\\monitoring_report\\data_import\\pending.xlsx'
path_under_review = 'C:\\Users\\alejandro.berzal\\Desktop\\DATA SCIENCE\\monitoring_report\\data_import\\under_review.xlsx'
path_to_upload = 'C:\\Users\\alejandro.berzal\\Desktop\\DATA SCIENCE\\monitoring_report\\data_import\\to_upload.xlsx'
path_to_graphics = 'C:\\Users\\alejandro.berzal\\Desktop\\DATA SCIENCE\\monitoring_report\\data_import\\total.xlsx'
# Leer el DataFrame desde una hoja existente
df = pd.read_excel(path_pending)
df2 = pd.read_excel(path_under_review)
df3 = pd.read_excel(path_to_upload)
df5 = pd.read_excel(path_to_graphics)
df6 = pd.read_excel(path_to_graphics)

# TRATAMIENTO DEL DATAFRAME "PENDING / df"
today_date = pd.to_datetime('today', format='%d-%m-%Y', dayfirst=True)  # Capturamos la fecha actual del día
today_date_str = today_date.strftime('%d-%m-%Y') # Formateamos la fecha_actual a strf para la lectura y guardado de archivos
# Transformamos todas las columnas de fechas to_datetime
df['Fecha'] = pd.to_datetime(df['Fecha'], format="%d-%m-%Y", dayfirst=True)
df['Fecha Pedido'] = pd.to_datetime(df['Fecha Pedido'], format="%d-%m-%Y", dayfirst=True)
df['Fecha Prevista'] = pd.to_datetime(df['Fecha Prevista'], format="%d-%m-%Y", dayfirst=True)
df.insert(12, "Notas", df['Estado']) # Insertar nueva columna 'Notas' en el dataframe
df['Notas'] = df['Fecha'] # Añadimos en la columna 'Notas' la fecha del pedido
# Sumar 15 días a la columna 'Notas' cuando la columna contiene 'Rechazado, Com. Menores, Com. Mayores, Comentado'
df.loc[df['Notas'] == df['Fecha'], 'Notas'] += pd.to_timedelta(15, unit='D')
df['Notas'] = "Enviar antes del " + df['Notas'].dt.strftime('%d-%m-%Y') # Transformamos la fecha a formato 'DIA-MES-AÑO'
df.insert(15, "Días Devolución", (today_date - df['Fecha']).dt.days) # Insertar nueva columna 'Días devolución' y se resta utilizando la fecha actual(today)
# Añadimos la columna 'Fecha Contractual' dividida en semanas
df.insert(18, 'Fecha Contractual', ((df['Fecha Prevista'] - df['Fecha Pedido']).dt.days // 7))
df['Fecha Contractual'] = "Aprobación + " + df['Fecha Contractual'].astype(str) + ' Semanas'
df.insert(16, "Fecha AP VDDL", df['Nº Pedido']) # Insertamos la columna 'Fecha AP VDDL'
process_vddl(df) # Aplicar el mapping para cambiar el tipo de estado en la columna 'Fecha AP VDDL'
apply_responsable(df)
identificar_cliente_por_PO(df) # Aplicar el mapping para cambiar el tipo de 'Cliente'
# Insertamos la columna 'Días VDDL'
df['Fecha AP VDDL'] = pd.to_datetime(df['Fecha AP VDDL'], format="mixed", dayfirst=True)
df.insert(17, "Días VDDL", (today_date - df['Fecha AP VDDL']).dt.days)
# Transformamos todas las fechas
df['Fecha'] = df['Fecha'].dt.date
df['Fecha Prevista'] = df['Fecha Prevista'].dt.date
df['Fecha Pedido'] = df['Fecha Pedido'].dt.date
df['Fecha AP VDDL'] = df['Fecha AP VDDL'].dt.date
print(df)

# TRATAMIENTO DEL DATAFRAME "UNDER REVIEW / df2"
# Transformamos todas las columnas de fechas to_datetime
df2['Fecha'] = pd.to_datetime(df2['Fecha'], format="%d-%m-%Y", dayfirst=True)
df2['Fecha Pedido'] = pd.to_datetime(df2['Fecha Pedido'], format="%d-%m-%Y", dayfirst=True)
df2['Fecha Prevista'] = pd.to_datetime(df2['Fecha Prevista'], format="%d-%m-%Y", dayfirst=True)
df2.insert(14, "Días Devolución", (today_date - df2['Fecha']).dt.days) # Insertar nueva columna 'Días Devolución' y restamos a la fecha actual para que nos de el total de días
# Añadimos la columna 'Fecha Contractual' dividida en semanas
df2.insert(15, 'Fecha Contractual', ((df2['Fecha Prevista'] - df2['Fecha Pedido']).dt.days // 7))
df2['Fecha Contractual'] = "Aprobación + " + df2['Fecha Contractual'].astype(str) + ' Semanas'
df2.insert(16, "Fecha AP VDDL", df2['Nº Pedido']) # Insertamos la columna 'Fecha AP VDDL'
process_vddl(df2) # Aplicar el mapeo para cambiar el tipo de estado en la columna 'Fecha AP VDDL'
apply_responsable(df2)
identificar_cliente_por_PO(df2) # Aplicar el mapping para cambiar el tipo de 'Cliente'
# Insertamos la columna 'Días VDDL'
df2['Fecha AP VDDL'] = pd.to_datetime(df2['Fecha AP VDDL'], format="mixed", dayfirst=True)
df2.insert(17, "Días VDDL", (today_date - df2['Fecha AP VDDL']).dt.days)
# Transformamos todas las fechas al formato 'DIA-MES-AÑO' sin la hora
df2['Fecha'] = df2['Fecha'].dt.date
df2['Fecha Prevista'] = df2['Fecha Prevista'].dt.date
df2['Fecha Pedido'] = df2['Fecha Pedido'].dt.date
df2['Fecha AP VDDL'] = df2['Fecha AP VDDL'].dt.date
print(df2)

# TRATAMIENTO DEL DATAFRAME "SIN ENVIAR /df3"
# Transformamos todas las columnas de fechas to_datetime
df3['Fecha'] = pd.to_datetime(df3['Fecha'], format="%d-%m-%Y", dayfirst=True)
df3['Fecha Pedido'] = pd.to_datetime(df3['Fecha Pedido'], format="%d-%m-%Y", dayfirst=True)
df3['Fecha Prevista'] = pd.to_datetime(df3['Fecha Prevista'], format="%d-%m-%Y", dayfirst=True)
df3['Estado'] = df3['Estado'].fillna('Sin Enviar') # Completamos la columna 'Estado' con 'Sin Enviar'
df3.insert(14, "Días Devolución", (today_date - df3['Fecha']).dt.days) # Insertar nueva columna 'Días Devolución' y restamos a la fecha actual para que nos de el total de días
# Añadimos la columna 'Fecha Contractual'
df3.insert(15, 'Fecha Contractual', ((df3['Fecha Prevista'] - df3['Fecha Pedido']).dt.days // 7))
df3['Fecha Contractual'] = "Aprobación + " + df3['Fecha Contractual'].astype(str) + ' Semanas'
apply_responsable(df3)
identificar_cliente_por_PO(df3) # Aplicar el mapping para cambiar el tipo de 'Cliente'
# Transformamos todas las fechas al formato 'DIA-MES-AÑO' sin la hora
df3['Fecha'] = df3['Fecha'].dt.date
df3['Fecha Prevista'] = df3['Fecha Prevista'].dt.date
df3['Fecha Pedido'] = df3['Fecha Pedido'].dt.date
critics_no = df3[df3['Crítico'] == 'No'] # Filtrar los documentos que tienen 'No' en la columna 'Crítico'
critics_si = df3[df3['Crítico'] == 'Sí'] # Filtrar los documentos que tienen 'Sí' en la columna 'Crítico'
print(df3)

# TRATAMIENTO DEL DATAFRAME "GRÁFICOS / df5"
df5['Estado'] = df5['Estado'].fillna('Sin Enviar') # Completamos la columna 'Estado' con 'Sin Enviar'
# Contar la frecuencia de cada estado por 'Nº Pedido'
df5 = df5.groupby(['Nº Pedido', 'Estado']).size().unstack(fill_value=0).reset_index()
# Eliminar la columna 'Eliminado' si existe
df5 = df5.drop(columns=['Eliminado'])
# Lista de columnas necesarias
columnas_necesarias = ['Aprobado', 'Com. Mayores', 'Com. Menores', 'Enviado', 'Rechazado', 'Sin Enviar']
# Asegurarnos de que cada columna necesaria existe y añadirla con 0 si no existe
for columna in columnas_necesarias:
    if columna not in df5.columns:
        df5[columna] = 0
df5['Total'] = df5.iloc[:, 1:8].sum(axis=1)
suma_total = df5['Total']
suma_total_general = df5['Aprobado']
# Calcular el porcentaje total
porcentaje_total = (suma_total_general / suma_total) * 100
df5['% Completado'] = porcentaje_total
df5 = df5.reindex(columns=['Nº Pedido', '% Completado', 'Aprobado', 'Com. Mayores', 'Com. Menores', 'Enviado', 'Rechazado', 'Sin Enviar', 'Total'])
print(df5)
print("Generando porcentaje total de los pedidos...")

# TRATAMIENTO DEL DATAFRAME "GRÁFICOS / df6"
df6['Fecha'] = pd.to_datetime(df6['Fecha'], format="%d-%m-%Y", dayfirst=True)
df6['Fecha Pedido'] = pd.to_datetime(df6['Fecha Pedido'], format="%d-%m-%Y", dayfirst=True)
df6['Fecha Prevista'] = pd.to_datetime(df6['Fecha Prevista'], format="%d-%m-%Y", dayfirst=True)
df6['Estado'] = df6['Estado'].fillna('Sin Enviar') # Completamos la columna 'Estado' con 'Sin Enviar'
# Eliminar la columna 'Eliminado' si existe
df6 = df6[df6['Estado'] != 'Eliminado']
df6 = df6[df6['Estado'] != 'Aprobado']
df6.insert(14, "Días Devolución", (today_date - df6['Fecha']).dt.days) # Insertar nueva columna 'Días Devolución' y restamos a la fecha actual para que nos de el total de días
# Añadimos la columna 'Fecha Contractual' dividida en semanas
df6.insert(15, 'Fecha Contractual', ((df6['Fecha Prevista'] - df6['Fecha Pedido']).dt.days // 7))
df6['Fecha Contractual'] = "Aprobación + " + df6['Fecha Contractual'].astype(str) + ' Semanas'
df6.insert(16, "Fecha AP VDDL", df6['Nº Pedido']) # Insertamos la columna 'Fecha AP VDDL'
process_vddl(df6) # Aplicar el mapeo para cambiar el tipo de estado en la columna 'Fecha AP VDDL'
apply_responsable(df6)
identificar_cliente_por_PO(df6) # Aplicar el mapping para cambiar el tipo de 'Cliente'
# Insertamos la columna 'Días VDDL'
df6['Fecha AP VDDL'] = pd.to_datetime(df6['Fecha AP VDDL'], format="mixed", dayfirst=True)
df6.insert(17, "Días VDDL", (today_date - df6['Fecha AP VDDL']).dt.days)
# Transformamos todas las fechas al formato 'DIA-MES-AÑO' sin la hora
df6['Fecha'] = df6['Fecha'].dt.date
df6['Fecha Prevista'] = df6['Fecha Prevista'].dt.date
df6['Fecha Pedido'] = df6['Fecha Pedido'].dt.date
df6['Fecha AP VDDL'] = df6['Fecha AP VDDL'].dt.date
print(df6)

# Leer la hoja "DATA" usando pandas para análisis
df_cal_pla = pd.read_excel(path_to_graphics)
# Filtrar por 'Tipo Doc.'
df_cal_pla = df_cal_pla[df_cal_pla['Tipo Doc.'] == 'Cálculo y plano']
# Rellenar valores nulos en 'Estado' con 'Sin Enviar'
df_cal_pla['Estado'] = df_cal_pla['Estado'].fillna('Sin Enviar')
# Reemplazar NaN en 'Nº Doc. Cliente' con 'Sin TR. Number'
df_cal_pla['Nº Doc. Cliente'] = df_cal_pla['Nº Doc. Cliente'].fillna('Sin TR. Number')
# Concatenar 'Nº Doc. Cliente' en caso de duplicados por 'Nº Pedido'
df_cal_pla['Nº Doc. Cliente'] = df_cal_pla.groupby('Nº Pedido')['Nº Doc. Cliente'].transform(lambda x: ';'.join(x.astype(str).unique()))
# Eliminar duplicados manteniendo solo la primera ocurrencia de cada 'Nº Pedido'
df_cal_pla = df_cal_pla.drop_duplicates(subset=['Nº Pedido', 'Estado'])
# Contar la frecuencia de cada estado por 'Nº Pedido'
df_cal_pla = df_cal_pla.groupby(['Nº Pedido', 'Nº Doc. Cliente', 'Estado']).size().unstack(fill_value=0).reset_index()
# Calcular el total y el porcentaje completado
df_cal_pla['Total'] = df_cal_pla.iloc[:, 2:].sum(axis=1)
suma_total = df_cal_pla['Total']
suma_total_general = df_cal_pla.get('Aprobado', 0)
porcentaje_total = (suma_total_general / suma_total) * 100
df_cal_pla['% Completado'] = porcentaje_total
# Calcular 'Aprobados' y 'Sin_Enviar' usando get para evitar errores si las columnas no existen
aprobados = df_cal_pla.get('Aprobado', 0) + df_cal_pla.get('Enviado', 0)
sin_enviar = df_cal_pla.get('Com. Menores', 0) + df_cal_pla.get('Sin Enviar', 0) + df_cal_pla.get('Com. Mayores', 0) + df_cal_pla.get('Rechazado', 0)
df_cal_pla['Enviados'] = aprobados
df_cal_pla['Pendiente'] = sin_enviar
# Reordenar columnas y filtrar por 'Sin_Enviar' > 0
df_cal_pla = df_cal_pla[['Nº Pedido', 'Enviados', 'Pendiente', 'Nº Doc. Cliente']]
df_cal_pla = df_cal_pla[df_cal_pla['Pendiente'] > 0]
print(df_cal_pla)

# Leer la hoja "DATA" usando pandas para análisis
df_planos = pd.read_excel(path_to_graphics)
# Filtrar por 'Tipo Doc.'
df_planos = df_planos[df_planos['Tipo Doc.'] == 'Planos']
# Rellenar valores nulos en 'Estado' con 'Sin Enviar'
df_planos['Estado'] = df_planos['Estado'].fillna('Sin Enviar')
# Reemplazar NaN en 'Nº Doc. Cliente' con 'Sin TR. Number'
df_planos['Nº Doc. Cliente'] = df_planos['Nº Doc. Cliente'].fillna('Sin TR. Number')
# Concatenar 'Nº Doc. Cliente' en caso de duplicados por 'Nº Pedido'
df_planos['Nº Doc. Cliente'] = df_planos.groupby('Nº Pedido')['Nº Doc. Cliente'].transform(lambda x: ';'.join(x.astype(str).unique()))
# Eliminar duplicados manteniendo solo la primera ocurrencia de cada 'Nº Pedido'
df_planos = df_planos.drop_duplicates(subset=['Nº Pedido', 'Estado'])
# Contar la frecuencia de cada estado por 'Nº Pedido'
df_planos = df_planos.groupby(['Nº Pedido', 'Nº Doc. Cliente', 'Estado']).size().unstack(fill_value=0).reset_index()
# Calcular el total y el porcentaje completado
df_planos['Total'] = df_planos.iloc[:, 2:].sum(axis=1)
suma_total = df_planos['Total']
suma_total_general = df_planos.get('Aprobado', 0)
porcentaje_total = (suma_total_general / suma_total) * 100
df_planos['% Completado'] = porcentaje_total
# Calcular 'Aprobados' y 'Sin_Enviar' usando get para evitar errores si las columnas no existen
aprobados = df_planos.get('Aprobado', 0) + df_planos.get('Enviado', 0)
sin_enviar = df_planos.get('Com. Menores', 0) + df_planos.get('Sin Enviar', 0) + df_planos.get('Com. Mayores', 0) + df_planos.get('Rechazado', 0)
df_planos['Enviados'] = aprobados
df_planos['Pendiente'] = sin_enviar
# Reordenar columnas y filtrar por 'Sin_Enviar' > 0
df_planos = df_planos[['Nº Pedido', 'Enviados', 'Pendiente', 'Nº Doc. Cliente']]
df_planos = df_planos[df_planos['Pendiente'] > 0]
print(df_planos)
print("Generando gráficos en la hoja de excel...")

# Reorganizamos las columnas
df = df.reindex(columns=['Nº Pedido', 'Resp.', 'Nº PO','Cliente', 'Material', 'Nº Doc. Cliente', 'Nº Doc. EIPSA', 'Título', 'Tipo Doc.', 'Crítico', 'Estado', 'Notas','Nº Revisión', 'Fecha', 'Días Devolución', 'Fecha Pedido', 'Fecha Prevista', 'Fecha Contractual', 'Fecha AP VDDL', 'Días VDDL', 'Historial Rev.', 'Seguimiento'])
df2 = df2.reindex(columns=['Nº Pedido', 'Resp.', 'Nº PO', 'Cliente', 'Material', 'Nº Doc. Cliente', 'Nº Doc. EIPSA', 'Título', 'Tipo Doc.' ,'Crítico', 'Estado', 'Nº Revisión', 'Fecha', 'Fecha Pedido', 'Días Devolución', 'Fecha Prevista',  'Fecha Contractual', 'Fecha AP VDDL', 'Días VDDL', 'Historial Rev.', 'Seguimiento'])
df3 = critics_no.reindex(columns=['Nº Pedido', 'Resp.', 'Nº PO', 'Cliente', 'Material', 'Nº Doc. Cliente', 'Nº Doc. EIPSA', 'Título', 'Tipo Doc.' , 'Crítico', 'Estado', 'Fecha Pedido', 'Fecha Prevista', 'Fecha Contractual'])
df4 = critics_si.reindex(columns=['Nº Pedido', 'Resp.', 'Nº PO', 'Cliente', 'Material', 'Nº Doc. Cliente', 'Nº Doc. EIPSA', 'Título', 'Tipo Doc.' , 'Crítico', 'Estado', 'Fecha Pedido', 'Fecha Prevista', 'Fecha Contractual'])
df6 = df6.reindex(columns=['Nº Pedido', 'Resp.', 'Nº PO', 'Cliente', 'Material', 'Nº Doc. Cliente', 'Nº Doc. EIPSA', 'Título', 'Tipo Doc.' ,'Crítico', 'Estado', 'Nº Revisión', 'Fecha', 'Fecha Pedido', 'Días Devolución', 'Fecha Prevista',  'Fecha Contractual', 'Historial Rev.', 'Seguimiento'])
df_cal_pla.to_excel('.\\data\\df_cal_pla_' + str(today_date_str) + '.xlsx', index=False)
df_planos.to_excel('.\\data\\df_planos_' + str(today_date_str) + '.xlsx', index=False)
print("¡Generando columnas...!")

# Crear un archivo Excel y escribir ambos DataFrames
output_path = f'./data/merged_data_{today_date_str}.xlsx'
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    # Escribir df_cal_pla en la hoja de cálculo
    df_cal_pla.to_excel(writer, sheet_name='ESTADO GLOBAL', index=False)
    # Obtener el objeto workbook y worksheet
    workbook = writer.book
    worksheet = writer.sheets['ESTADO GLOBAL']
    # Escribir df_planos a partir de la columna N
    startrow = 0
    startcol = 11  # Columna N es la columna 13 (índice 12 en 0-based indexing)
    # Escribir encabezados de df_planos a partir de la columna N
    for c_idx, value in enumerate(df_planos.columns.values):
        worksheet.write(startrow, startcol + c_idx, value)
    # Escribir datos de df_planos a partir de la columna N
    for r_idx, row in enumerate(df_planos.values):
        for c_idx, value in enumerate(row):
            worksheet.write(startrow + r_idx + 1, startcol + c_idx, value)

graph_merge = f'./data/merged_data_{today_date_str}.xlsx'
df_graph_merge = pd.read_excel(graph_merge)

# Seleccionamos las columnas que van a ser coloreadas según el 'ESTADO' que tiene la documentación
with pd.ExcelWriter(r'C:\\Users\\alejandro.berzal\\Desktop\\DATA SCIENCE\\monitoring_report\\data\\monitoring_report_' + str(today_date_str) + '.xlsx', engine='xlsxwriter') as writer:
    # Aplicar estilos a cada hoja de excel
    style_sheet6 = df6.style.apply(
        highlight_row_content, value="Rechazado", color='#FFA19A', subset=["Estado"], axis=1).apply(
        highlight_row_content, value="Com. Menores", color='#FFE5AD', subset=["Estado"], axis=1).apply(
        highlight_row_content, value="Com. Mayores", color='#DBB054', subset=["Estado"], axis=1).apply(
        highlight_row_content, value="Comentado", color='#F79646', subset=["Estado"], axis=1).apply(
        highlight_row_content, value="Enviado", color='#B1E1B9', subset=["Estado"], axis=1).apply(
        highlight_row_content, value="Sin Enviar", color='#FFFFAB', subset=["Estado"], axis=1).apply(
        highlight_row_content, value="Aprobado", color='#D4DCF4', subset=["Estado"], axis=1)
    style_sheet6.to_excel(writer, sheet_name='DOC. TOTAL',index=False)  # Escribir el DataFrame con estilos en la hoja 'pending'
    style_sheet = df.style.apply(
        highlight_row_content, value="Rechazado", color='#FFA19A', subset=["Estado", "Notas"], axis=1).apply(
        highlight_row_content, value="Com. Menores", color='#FFE5AD', subset=["Estado", "Notas"], axis=1).apply(
        highlight_row_content, value="Com. Mayores", color='#DBB054', subset=["Estado", "Notas"], axis=1).apply(
        highlight_row_content, value="Comentado", color='#F79646', subset=["Estado", "Notas"], axis=1)
    style_sheet.to_excel(writer, sheet_name='DOC. COMENTARIOS', index=False) # Escribir el DataFrame con estilos en la hoja 'pending'
    style_sheet_2 = df2.style.apply(highlight_row_content, value="Enviado", color='#B1E1B9', subset=["Estado"], axis=1) # Aplicar estilos al DataFrame 'df_under_review'
    style_sheet_2.to_excel(writer, sheet_name='DOC. ENVIADA', index=False) # Escribir el DataFrame con estilos en la hoja 'df_under_review'
    style_sheet_3 = df3.style.apply(highlight_row_content, value='Sin Enviar', color='#FFFFAB', subset=["Estado"], axis=1) # Aplicar estilos al DataFrame 'df_to_upload'
    style_sheet_3.to_excel(writer, sheet_name='DOC. SIN ENVIAR', index=False) # Escribir el DataFrame con estilos en la hoja 'to_upload'
    style_sheet_4 = df4.style.apply(highlight_row_content, value='Sin Enviar', color='#FFFFAB', subset=["Estado"], axis=1).apply(highlight_row_content, value=" ", color='#FFFFAB', subset=["Estado"], axis=1) # Aplicar estilos al DataFrame 'df_to_upload'
    style_sheet_4.to_excel(writer, sheet_name='CRÍTICOS', index=False) # Escribir el DataFrame con estilos en la hoja 'to_upload'
    df_graph_merge.to_excel(writer, sheet_name='GRÁFICO CRÍTICOS', index=False)
    df5.to_excel(writer, sheet_name='ESTADO GLOBAL', index=False) # Escribir el DataFrame con estilos en la hoja 'pending'
print("¡Estilo, formato y color aplicado correctamente a todas las hojas del excel!")


# Cargar archivo de Excel con las tres hojas de datos
wb = Workbook('C:\\Users\\alejandro.berzal\\Desktop\\DATA SCIENCE\\monitoring_report\\data\\monitoring_report_' + today_date_str + '.xlsx')
# Obtener la referencia de las hojas/sheets de trabajo deseadas
sheets = {"DOC. TOTAL": wb.getWorksheets().get("DOC. TOTAL"),
          "DOC. COMENTARIOS": wb.getWorksheets().get("DOC. COMENTARIOS"),
          "DOC. ENVIADA": wb.getWorksheets().get("DOC. ENVIADA"),
          "DOC. SIN ENVIAR": wb.getWorksheets().get("DOC. SIN ENVIAR"),
          "CRÍTICOS": wb.getWorksheets().get("CRÍTICOS"),
          "GRÁFICO CRÍTICOS": wb.getWorksheets().get("GRÁFICO CRÍTICOS"),
          "ESTADO GLOBAL":wb.getWorksheets().get("ESTADO GLOBAL")}

# Ajuste automático de todas las columnas en cada hoja
for sheet_name, sheet in sheets.items():
    if sheet:
        auto_fit_columns(sheet)
wb.save("Monitoring_Report_" + str(today_date_str) + ".xlsx") # Guardar libro de trabajo
print("¡Columnas y celdas ajustadas para una mejor visualización!")

# Utilizamos la función para aplicar todos los estilos y coloreado del excel
apply_excel_styles(today_date_str)

# Eliminamos la sheet evaluation warning que nos genera ASPOSECELLS
df_final = load_workbook("Monitoring_Report_" + str(today_date_str) + ".xlsx")
# Verificar si la hoja "Evaluation Warning" existe
if "Evaluation Warning" in df_final.sheetnames:
    del df_final["Evaluation Warning"] # Eliminar la hoja "Evaluation Warning"

# Guardar los cambios en el archivo
df_final.save("Monitoring_Report_" + str(today_date_str) + ".xlsx")
print("¡Exito! ¡Archivo Excel guardado en la carpeta C:\\Users\\alejandro.berzal\\Desktop\\DATA SCIENCE\\monitoring_report...!")
print("Duración del proceso: {} seconds".format(time.time() - start_time))
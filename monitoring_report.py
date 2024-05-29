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
df5['Total'] = df5.iloc[:, 1:8].sum(axis=1)
suma_total = df5['Total']
suma_total_general = df5['Aprobado']
# Calcular el porcentaje total
porcentaje_total = (suma_total_general / suma_total) * 100
df5['% Completado'] = porcentaje_total
print(df5)
print("Generando porcentaje total de los pedidos...")

# Reorganizamos las columnas
df = df.reindex(columns=['Nº Pedido', 'Resp.', 'Nº PO','Cliente', 'Material', 'Nº Doc. Cliente', 'Nº Doc. EIPSA', 'Título', 'Tipo Doc.', 'Crítico', 'Estado', 'Notas','Nº Revisión', 'Fecha', 'Días Devolución', 'Fecha Pedido', 'Fecha Prevista', 'Fecha Contractual', 'Fecha AP VDDL', 'Días VDDL', 'Historial Rev.', 'Seguimiento'])
df2 = df2.reindex(columns=['Nº Pedido', 'Resp.', 'Nº PO', 'Cliente', 'Material', 'Nº Doc. Cliente', 'Nº Doc. EIPSA', 'Título', 'Tipo Doc.' ,'Crítico', 'Estado', 'Nº Revisión', 'Fecha', 'Días Devolución', 'Fecha Pedido', 'Fecha Prevista', 'Fecha Contractual', 'Fecha AP VDDL', 'Días VDDL', 'Historial Rev.', 'Seguimiento'])
df3 = critics_no.reindex(columns=['Nº Pedido', 'Resp.', 'Nº PO', 'Cliente', 'Material', 'Nº Doc. Cliente', 'Nº Doc. EIPSA', 'Título', 'Tipo Doc.' , 'Crítico', 'Estado', 'Fecha Pedido', 'Fecha Prevista', 'Fecha Contractual'])
df4 = critics_si.reindex(columns=['Nº Pedido', 'Resp.', 'Nº PO', 'Cliente', 'Material', 'Nº Doc. Cliente', 'Nº Doc. EIPSA', 'Título', 'Tipo Doc.' , 'Crítico', 'Estado', 'Fecha Pedido', 'Fecha Prevista', 'Fecha Contractual'])
# Se genera archivo excel para su almacenamiento
#df.to_excel('.\\data\\pending_' + str(today_date_str) + '.xlsx', index=False)
#df2.to_excel('.\\data\\under_review_' + str(today_date_str) + '.xlsx', index=False)
#df3.to_excel('.\\data\\to_upload_' + str(today_date_str) + '.xlsx', index=False)
#df4.to_excel('.\\data\\to_upload_criticos_' + str(today_date_str) + '.xlsx', index=False)
#df5.to_excel('.\\data\\df_graphics_' + str(today_date_str) + '.xlsx', index=False)
print("¡Generando columnas...!")

# Seleccionamos las columnas que van a ser coloreadas según el 'ESTADO' que tiene la documentación
with pd.ExcelWriter('C:\\Users\\alejandro.berzal\\Desktop\\DATA SCIENCE\\monitoring_report\\data\\monitoring_report_' + str(today_date_str) + '.xlsx', engine='xlsxwriter') as writer:
    # Aplicar estilos a cada hoja de excel
    style_sheet = df.style.apply(
        highlight_row_content, value="Rechazado", color='#FFA19A', subset=["Estado", "Notas"], axis=1).apply(
        highlight_row_content, value="Com. Menores", color='#FFE5AD', subset=["Estado", "Notas"], axis=1).apply(
        highlight_row_content, value="Com. Mayores", color='#DBB054', subset=["Estado", "Notas"], axis=1).apply(
        highlight_row_content, value="Comentado", color='#F79646', subset=["Estado", "Notas"], axis=1)
    style_sheet.to_excel(writer, sheet_name='Documentación con comentarios', index=False) # Escribir el DataFrame con estilos en la hoja 'pending'
    style_sheet_2 = df2.style.apply(highlight_row_content, value="Enviado", color='#B1E1B9', subset=["Estado"], axis=1) # Aplicar estilos al DataFrame 'df_under_review'
    style_sheet_2.to_excel(writer, sheet_name='Enviada para aprobación', index=False) # Escribir el DataFrame con estilos en la hoja 'df_under_review'
    style_sheet_3 = df3.style.apply(highlight_row_content, value='Sin Enviar', color='#FFFFAB', subset=["Estado"], axis=1) # Aplicar estilos al DataFrame 'df_to_upload'
    style_sheet_3.to_excel(writer, sheet_name='Documentación sin enviar', index=False) # Escribir el DataFrame con estilos en la hoja 'to_upload'
    style_sheet_4 = df4.style.apply(highlight_row_content, value='Sin Enviar', color='#FFFFAB', subset=["Estado"], axis=1).apply(highlight_row_content, value=" ", color='#FFFFAB', subset=["Estado"], axis=1) # Aplicar estilos al DataFrame 'df_to_upload'
    style_sheet_4.to_excel(writer, sheet_name='CRÍTICOS', index=False) # Escribir el DataFrame con estilos en la hoja 'to_upload'
    df5.to_excel(writer, sheet_name='DATA', index=False) # Escribir el DataFrame con estilos en la hoja 'pending'
print("¡Estilo, formato y color aplicado correctamente a todas las hojas del excel!")


# Cargar archivo de Excel con las tres hojas de datos
wb = Workbook('C:\\Users\\alejandro.berzal\\Desktop\\DATA SCIENCE\\monitoring_report\\data\\monitoring_report_' + today_date_str + '.xlsx')
# Obtener la referencia de las hojas/sheets de trabajo deseadas
sheets = {"Documentación con comentarios": wb.getWorksheets().get("Documentación con comentarios"),
          "Enviada para aprobación": wb.getWorksheets().get("Enviada para aprobación"),
          "Documentación sin enviar": wb.getWorksheets().get("Documentación sin enviar"),
          "CRÍTICOS": wb.getWorksheets().get("CRÍTICOS"),
          "DATA":wb.getWorksheets().get("DATA")}

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
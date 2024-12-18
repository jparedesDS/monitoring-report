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
import warnings
from tqdm import tqdm


start_time = time.time()
warnings.filterwarnings('ignore')
# Ruta del archivo excel con todos lo datos del ERP
path_total = 'C:\\Users\\alejandro.berzal\\Desktop\\DATA SCIENCE\\monitoring_report\\data_import\\data_erp.xlsx'

# Tratamiento dataset general
erp_data = pd.read_excel(path_total)
df_total = pd.read_excel(path_total)
erp_data['Estado'] = erp_data['Estado'].fillna('Sin Enviar') # Completamos la columna 'Estado' con 'Sin Enviar'
df_total['Estado'] = df_total['Estado'].fillna('Sin Enviar') # Completamos la columna 'Estado' con 'Sin Enviar'
#df_total['Estado'] = df_total['Estado'].replace('Aprobado', 'APROBADO') # Ponemos en mayusculas
# Transformamos todas las columnas de fechas to_datetime
today_date = pd.to_datetime('today', format='%d-%m-%Y', dayfirst=True)  # Capturamos la fecha actual del día
today_date_str = today_date.strftime('%d-%m-%Y') # Formateamos la fecha_actual a strf para la lectura y guardado de archivos
erp_data['Fecha'] = pd.to_datetime(erp_data['Fecha'], format="%d-%m-%Y", dayfirst=True)
erp_data['Fecha Pedido'] = pd.to_datetime(erp_data['Fecha Pedido'], format="%d-%m-%Y", dayfirst=True)
erp_data['Fecha Prevista'] = pd.to_datetime(erp_data['Fecha Prevista'], format="%d-%m-%Y", dayfirst=True)
df_total['Fecha'] = pd.to_datetime(df_total['Fecha'], format="%d-%m-%Y", dayfirst=True)
df_total['Fecha Pedido'] = pd.to_datetime(df_total['Fecha Pedido'], format="%d-%m-%Y", dayfirst=True)
df_total['Fecha Prevista'] = pd.to_datetime(df_total['Fecha Prevista'], format="%d-%m-%Y", dayfirst=True)

# Tratamiento del dataframe "Pending" // df_comentados
df_menores = erp_data[erp_data['Estado'] == 'Com. Menores']
df_mayores = erp_data[erp_data['Estado'] == 'Com. Mayores']
df_rechazado = erp_data[erp_data['Estado'] == 'Rechazado']
df_comentados = pd.concat([df_menores, df_mayores, df_rechazado]) # Unimos los tres DataFrames
# Transformamos todas las columnas de fechas to_datetime
df_comentados.insert(12, "Notas", df_comentados['Estado']) # Insertar nueva columna 'Notas' en el dataframe
df_comentados['Notas'] = df_comentados['Fecha'] # Añadimos en la columna 'Notas' la fecha del pedido
# Sumar 15 días a la columna 'Notas' cuando la columna contiene 'Rechazado, Com. Menores, Com. Mayores, Comentado'
df_comentados.loc[df_comentados['Notas'] == df_comentados['Fecha'], 'Notas'] += pd.to_timedelta(15, unit='D')
df_comentados['Notas'] = "Enviar antes del " + df_comentados['Notas'].dt.strftime('%d-%m-%Y') # Transformamos la fecha a formato 'DIA-MES-AÑO'
df_comentados.insert(15, "Días Devolución", (today_date - df_comentados['Fecha']).dt.days) # Insertar nueva columna 'Días devolución' y se resta utilizando la fecha actual(today)
# Añadimos la columna 'Fecha Contractual' dividida en semanas
df_comentados.insert(18, 'Fecha Contractual', ((df_comentados['Fecha Prevista'] - df_comentados['Fecha Pedido']).dt.days // 7))
df_comentados['Fecha Contractual'] = "Aprobación + " + df_comentados['Fecha Contractual'].astype(str) + ' Semanas'
df_comentados.insert(16, "Fecha AP VDDL", df_comentados['Nº Pedido']) # Insertamos la columna 'Fecha AP VDDL'
process_vddl(df_comentados) # Aplicar el mapping para cambiar el tipo de estado en la columna 'Fecha AP VDDL'
apply_responsable(df_comentados)
identificar_cliente_por_PO(df_comentados) # Aplicar el mapping para cambiar el tipo de 'Cliente'
apply_reclamaciones(df_comentados) # Aplicar el mapping para indicar cuantas reclamaciones lleva el documento.
# Insertamos la columna 'Días VDDL'
df_comentados['Fecha AP VDDL'] = pd.to_datetime(df_comentados['Fecha AP VDDL'], format="mixed", dayfirst=True)
df_comentados.insert(17, "Días VDDL", (today_date - df_comentados['Fecha AP VDDL']).dt.days)
# Cambiar el nombre de la columna 'viejo_nombre' a 'nuevo_nombre'
df_comentados = df_comentados.rename(columns={'Fecha': 'Fecha Dev. Doc.'})
df_comentados = df_comentados.rename(columns={'Fecha Prevista': 'Fecha FIN'})
df_comentados = df_comentados.rename(columns={'Fecha Pedido': 'Fecha INICIAL'})
print(df_comentados)

# Tratamiento del dataframe "Under review" // df_envio
df_envio = erp_data[erp_data['Estado'] == 'Enviado']
df_envio.insert(14, "Días Devolución", (today_date - df_envio['Fecha']).dt.days) # Insertar nueva columna 'Días Devolución' y restamos a la fecha actual para que nos de el total de días
# Añadimos la columna 'Fecha Contractual' dividida en semanas
df_envio.insert(15, 'Fecha Contractual', ((df_envio['Fecha Prevista'] - df_envio['Fecha Pedido']).dt.days // 7))
df_envio['Fecha Contractual'] = "Aprobación + " + df_envio['Fecha Contractual'].astype(str) + ' Semanas'
df_envio.insert(16, "Fecha AP VDDL", df_envio['Nº Pedido']) # Insertamos la columna 'Fecha AP VDDL'
process_vddl(df_envio) # Aplicar el mapeo para cambiar el tipo de estado en la columna 'Fecha AP VDDL'
apply_responsable(df_envio)
identificar_cliente_por_PO(df_envio) # Aplicar el mapping para cambiar el tipo de 'Cliente'
apply_reclamaciones(df_envio) # Aplicar el mapping para indicar cuantas reclamaciones lleva el documento.
# Insertamos la columna 'Días VDDL'
df_envio['Fecha AP VDDL'] = pd.to_datetime(df_envio['Fecha AP VDDL'], format="mixed", dayfirst=True)
df_envio.insert(17, "Días VDDL", (today_date - df_envio['Fecha AP VDDL']).dt.days)
df_envio = df_envio.rename(columns={'Fecha': 'Fecha Env. Doc.'})
df_envio = df_envio.rename(columns={'Fecha Prevista': 'Fecha FIN'})
df_envio = df_envio.rename(columns={'Fecha Pedido': 'Fecha INICIAL'})
print(df_envio)

# TRATAMIENTO DEL DATAFRAME "SIN ENVIAR // df_sin_envio"
df_sin_envio = erp_data[erp_data['Estado'] == 'Sin Enviar']
# Añadimos la columna 'Fecha Contractual'
df_sin_envio.insert(14, 'Fecha Contractual', ((df_sin_envio['Fecha Prevista'] - df_sin_envio['Fecha Pedido']).dt.days // 7))
df_sin_envio['Fecha Contractual'] = "Aprobación + " + df_sin_envio['Fecha Contractual'].astype(str) + ' Semanas'
df_sin_envio.insert(15, "Días Devolución", (today_date - df_sin_envio['Fecha Pedido']).dt.days) # Insertar nueva columna 'Días Devolución' y restamos a la fecha actual para que nos de el total de días
df_sin_envio = df_sin_envio.rename(columns={'Fecha Prevista': 'Fecha FIN'})
df_sin_envio = df_sin_envio.rename(columns={'Fecha Pedido': 'Fecha INICIAL'})
apply_responsable(df_sin_envio)
identificar_cliente_por_PO(df_sin_envio) # Aplicar el mapping para cambiar el tipo de 'Cliente'
print(df_sin_envio)

# TRATAMIENTO DEL DATAFRAME "CRÍTICOS" Crítico
df_criticos = erp_data[erp_data['Crítico'] == 'Sí']
df_criticos = df_criticos[df_criticos['Estado'] != 'Eliminado']
df_criticos = df_criticos[df_criticos['Estado'] != 'Aprobado']
df_criticos = df_criticos[df_criticos['Estado'] != 'Enviado']
df_criticos.insert(14, "Días Devolución", (today_date - df_criticos['Fecha']).dt.days) # Insertar nueva columna 'Días Devolución' y restamos a la fecha actual para que nos de el total de días
# Añadimos la columna 'Fecha Contractual'
df_criticos.insert(15, 'Fecha Contractual', ((df_criticos['Fecha Prevista'] - df_criticos['Fecha Pedido']).dt.days // 7))
df_criticos['Fecha Contractual'] = "Aprobación + " + df_criticos['Fecha Contractual'].astype(str) + ' Semanas'
apply_responsable(df_criticos)
identificar_cliente_por_PO(df_criticos) # Aplicar el mapping para cambiar el tipo de 'Cliente'
apply_reclamaciones(df_criticos) # Aplicar el mapping para indicar cuantas reclamaciones lleva el documento.
df_criticos = df_criticos.rename(columns={'Fecha': 'Fecha Doc.'})
df_criticos = df_criticos.rename(columns={'Fecha Prevista': 'Fecha FIN'})
df_criticos = df_criticos.rename(columns={'Fecha Pedido': 'Fecha INICIAL'})
critics_si = df_criticos[df_criticos['Crítico'] == 'Sí'] # Filtrar los documentos que tienen 'Sí' en la columna 'Crítico'
print(df_criticos)

# TRATAMIENTO DEL DATAFRAME "GRÁFICOS / df5"
# Contar la frecuencia de cada estado por 'Nº Pedido'
erp_data = erp_data.groupby(['Nº Pedido', 'Estado']).size().unstack(fill_value=0).reset_index()
# Eliminar la columna 'Eliminado' si existe
erp_data = erp_data.drop(columns=['Eliminado'])
# Lista de columnas necesarias
columnas_necesarias = ['Aprobado', 'Com. Mayores', 'Com. Menores', 'Enviado', 'Rechazado', 'Sin Enviar']
# Asegurarnos de que cada columna necesaria existe y añadirla con 0 si no existe
for columna in columnas_necesarias:
    if columna not in erp_data.columns:
        erp_data[columna] = 0
erp_data['Total'] = erp_data.iloc[:, 1:8].sum(axis=1)
suma_total = erp_data['Total']
suma_total_general = erp_data['Aprobado']
# Calcular el porcentaje total
porcentaje_total = (suma_total_general / suma_total) * 100
erp_data['% Completado'] = porcentaje_total
erp_data = erp_data.reindex(columns=['Nº Pedido', '% Completado', 'Aprobado', 'Com. Mayores', 'Com. Menores', 'Enviado', 'Rechazado', 'Sin Enviar', 'Total'])
# Ordenar los datos por una columna específica en orden descendente (de Z a A)
columna_para_ordenar = 'Nº Pedido'  # Reemplaza con el nombre de tu columna
erp_data = erp_data.sort_values(by=columna_para_ordenar, ascending=False)
erp_data['% Completado'] = erp_data['% Completado'].fillna(0) # Completamos la columna '% Completado' con '0'
erp_data = erp_data[erp_data['% Completado'] != 100] # Eliminamos los pedidos que se encuentren 100% completos
erp_data = erp_data.round(2) # Que muestre máximo 2 decimales
print(erp_data)
print("Generando porcentaje total de los pedidos...")

# TRATAMIENTO DEL DATAFRAME "TODOS LOS DOCUMENTOS"
# Eliminar la columna 'Eliminado' si existe
df_total = df_total[df_total['Estado'] != 'Eliminado']
#df_total = df_total[df_total['Estado'] != 'Final'] # Se puede añadir todos los aprobados al total eliminando esta opción
df_total.insert(14, "Días Devolución", (today_date - df_total['Fecha']).dt.days) # Insertar nueva columna 'Días Devolución' y restamos a la fecha actual para que nos de el total de días
# Añadimos la columna 'Fecha Contractual' dividida en semanas
df_total.insert(15, 'Fecha Contractual', ((df_total['Fecha Prevista'] - df_total['Fecha Pedido']).dt.days // 7))
df_total['Fecha Contractual'] = "Aprobación + " + df_total['Fecha Contractual'].astype(str) + ' Semanas'
df_total.insert(16, "Fecha AP VDDL", df_total['Nº Pedido']) # Insertamos la columna 'Fecha AP VDDL'
process_vddl(df_total) # Aplicar el mapeo para cambiar el tipo de estado en la columna 'Fecha AP VDDL'
apply_responsable(df_total)
identificar_cliente_por_PO(df_total) # Aplicar el mapping para cambiar el tipo de 'Cliente'
apply_reclamaciones(df_total) # Aplicar el mapping para indicar cuantas reclamaciones lleva el documento.
# Insertamos la columna 'Días VDDL'
df_total['Fecha AP VDDL'] = pd.to_datetime(df_total['Fecha AP VDDL'], format="mixed", dayfirst=True)
df_total.insert(17, "Días VDDL", (today_date - df_total['Fecha AP VDDL']).dt.days)
df_total = df_total.rename(columns={'Fecha': 'Fecha Doc.'})
df_total = df_total.rename(columns={'Fecha Prevista': 'Fecha FIN'})
df_total = df_total.rename(columns={'Fecha Pedido': 'Fecha INICIAL'})
print(df_total)

# Reorganizamos las columnas
df_comentados = df_comentados.reindex(columns=['Nº Pedido', 'Resp.', 'Nº PO','Cliente', 'Material', 'Nº Doc. Cliente', 'Nº Doc. EIPSA', 'Título', 'Tipo Doc.', 'Crítico', 'Estado', 'Notas','Nº Revisión', 'Fecha Dev. Doc.', 'Días Devolución', 'Fecha INICIAL', 'Fecha FIN', 'Reclamaciones', 'Seguimiento', 'Historial Rev.']) # 'Fecha AP VDDL', 'Días VDDL',
df_envio = df_envio.reindex(columns=['Nº Pedido', 'Resp.', 'Nº PO', 'Cliente', 'Material', 'Nº Doc. Cliente', 'Nº Doc. EIPSA', 'Título', 'Tipo Doc.' , 'Crítico', 'Estado', 'Nº Revisión', 'Fecha Env. Doc.', 'Días Devolución', 'Fecha INICIAL', 'Fecha FIN', 'Reclamaciones', 'Seguimiento', 'Historial Rev.']) # 'Fecha AP VDDL', 'Días VDDL',
df_sin_envio = df_sin_envio.reindex(columns=['Nº Pedido', 'Resp.', 'Nº PO', 'Cliente', 'Material', 'Nº Doc. Cliente', 'Nº Doc. EIPSA', 'Título', 'Tipo Doc.' , 'Crítico', 'Estado', 'Fecha INICIAL', 'Fecha FIN', 'Seguimiento'])
df_criticos = critics_si.reindex(columns=['Nº Pedido', 'Resp.', 'Nº PO', 'Cliente', 'Material', 'Nº Doc. Cliente', 'Nº Doc. EIPSA', 'Título', 'Tipo Doc.' , 'Crítico', 'Estado','Nº Revisión', 'Fecha Doc.', 'Días Devolución', 'Fecha INICIAL', 'Fecha FIN', 'Reclamaciones', 'Seguimiento', 'Historial Rev.'])
df_total = df_total.reindex(columns=['Nº Pedido', 'Resp.', 'Nº PO', 'Cliente', 'Material', 'Nº Doc. Cliente', 'Nº Doc. EIPSA', 'Título', 'Tipo Doc.' , 'Crítico', 'Estado', 'Nº Revisión', 'Fecha Doc.', 'Fecha INICIAL', 'Fecha FIN', 'Reclamaciones', 'Seguimiento', 'Historial Rev.'])
print("¡Generando columnas...!")

# Seleccionamos las columnas que van a ser coloreadas según el 'ESTADO' que tiene la documentación
with pd.ExcelWriter(r'C:\\Users\\alejandro.berzal\\Desktop\\DATA SCIENCE\\monitoring_report\\data\\monitoring_report_' + str(today_date_str) + '.xlsx', engine='xlsxwriter') as writer:
    # Aplicar estilos a cada hoja de excel
    style_sheet6 = df_total.style.apply(
        highlight_row_content, value="Rechazado", color='#FFA19A', subset=["Estado"], axis=1).apply(
        highlight_row_content, value="Com. Menores", color='#FFE5AD', subset=["Estado"], axis=1).apply(
        highlight_row_content, value="Com. Mayores", color='#DBB054', subset=["Estado"], axis=1).apply(
        highlight_row_content, value="Comentado", color='#F79646', subset=["Estado"], axis=1).apply(
        highlight_row_content, value="Enviado", color='#B1E1B9', subset=["Estado"], axis=1).apply(
        highlight_row_content, value="Sin Enviar", color='#FFFFAB', subset=["Estado"], axis=1).apply(
        highlight_row_content, value="Aprobado", color='#00D25F', subset=["Estado"], axis=1)
    style_sheet6.to_excel(writer, sheet_name='ALL DOC.',index=False)  # Grabar el DataFrame con estilos en la hoja 'pending'
    style_sheet_2 = df_envio.style.apply(highlight_row_content, value="Enviado", color='#B1E1B9', subset=["Estado"], axis=1) # Aplicar estilos al DataFrame 'df_under_review'
    style_sheet_2.to_excel(writer, sheet_name='ENVIADOS', index=False) # Grabar el DataFrame con estilos en la hoja 'df_under_review'
    style_sheet_3 = df_sin_envio.style.apply(highlight_row_content, value='Sin Enviar', color='#FFFFAB', subset=["Estado"], axis=1) # Aplicar estilos al DataFrame 'df_to_upload'
    style_sheet_3.to_excel(writer, sheet_name='SIN ENVIAR', index=False) # Grabar el DataFrame con estilos en la hoja 'to_upload'
    style_sheet = df_comentados.style.apply(
        highlight_row_content, value="Rechazado", color='#FFA19A', subset=["Estado", "Notas"], axis=1).apply(
        highlight_row_content, value="Com. Menores", color='#FFE5AD', subset=["Estado", "Notas"], axis=1).apply(
        highlight_row_content, value="Com. Mayores", color='#DBB054', subset=["Estado", "Notas"], axis=1).apply(
        highlight_row_content, value="Comentado", color='#F79646', subset=["Estado", "Notas"], axis=1)
    style_sheet.to_excel(writer, sheet_name='COMENTADOS', index=False) # Grabar el DataFrame con estilos en la hoja 'pending'
    style_sheet_4 = df_criticos.style.apply(highlight_row_content, value="Rechazado", color='#FFA19A', subset=["Estado"], axis=1).apply(
        highlight_row_content, value="Com. Menores", color='#FFE5AD', subset=["Estado"], axis=1).apply(
        highlight_row_content, value="Com. Mayores", color='#DBB054', subset=["Estado"], axis=1).apply(
        highlight_row_content, value="Comentado", color='#F79646', subset=["Estado"], axis=1).apply(
        highlight_row_content, value="Sin Enviar", color='#FFFFAB', subset=["Estado"], axis=1) # Aplicar estilos al DataFrame 'df_to_upload'
    style_sheet_4.to_excel(writer, sheet_name='CRÍTICOS', index=False) # Grabar el DataFrame con estilos en la hoja 'to_upload'
    erp_data.to_excel(writer, sheet_name='STATUS GLOBAL', index=False) # Grabar el DataFrame con estilos en la hoja 'pending'
print("¡Estilo, formato y color aplicado correctamente a todas las hojas del excel!")

# Cargar archivo de Excel con las tres hojas de datos
wb = Workbook('C:\\Users\\alejandro.berzal\\Desktop\\DATA SCIENCE\\monitoring_report\\data\\monitoring_report_' + today_date_str + '.xlsx')
# Obtener la referencia de las hojas/sheets de trabajo deseadas
sheets = {"ALL DOC.": wb.getWorksheets().get("ALL DOC."),
          "COMENTADOS": wb.getWorksheets().get("COMENTADOS"),
          "ENVIADOS": wb.getWorksheets().get("ENVIADOS"),
          "SIN ENVIAR": wb.getWorksheets().get("SIN ENVIAR"),
          "CRÍTICOS": wb.getWorksheets().get("CRÍTICOS"),
          "STATUS GLOBAL":wb.getWorksheets().get("STATUS GLOBAL")}

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
df_final.save("Z:\\JOSE\\01 MONITORING REPORT\\Monitoring_Report_" + str(today_date_str) + ".xlsx")
print("¡Exito! ¡Archivo Excel guardado en la carpeta Z:\\JOSE\\01 MONITORING REPORT\\monitoring_report...!")
print("Duración del proceso: {} seconds".format(time.time() - start_time))
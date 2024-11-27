# Monitoring Report for ERP

Este script realiza el procesamiento de datos extraídos de un ERP y los organiza en diferentes dataframes según su estado. Además, genera un archivo Excel estilizado con varias hojas que representan diferentes vistas del estado de los documentos.

## Requisitos Previos

### Librerías Necesarias
Asegúrate de instalar las siguientes librerías antes de ejecutar el script:

```bash
pip install pandas xlsxwriter openpyxl tqdm aspose-cells
```

## Estructura del Proyecto
El script utiliza herramientas y mapeos personalizados que deben estar incluidos en la carpeta tools. 
Estas herramientas son:

- mapping_mr.py
- apply_style_mr.py

## Funcionalidades del Script
#### Principales Funciones
1. Importación y Limpieza de Datos
- Carga un archivo Excel con información del ERP.
- Realiza limpieza y formateo, como llenar valores nulos y convertir fechas.
2. Procesamiento por Estado
- Divide los datos en diferentes grupos (Enviado, Sin Enviar, Comentado, etc.).
- Calcula métricas como días de devolución, semanas contractuales, y notas adicionales.
3. Estilización y Exportación
- Aplica estilos personalizados a los dataframes.
- Exporta los datos a un archivo Excel con hojas separadas.
  
## Organización de Hojas en Excel
El archivo final incluye las siguientes hojas:

- ALL DOC.: Todos los documentos con estilos según su estado.
- ENVIADOS: Documentos en estado "Enviado".
- SIN ENVIAR: Documentos en estado "Sin Enviar".
- COMENTADOS: Documentos con comentarios ("Com. Menores", "Com. Mayores", etc.).
- STATUS: Gráfica de seguimiento general.

## Estructura del Código
#### Imports y Configuración Inicial
El script importa las librerías necesarias y configura la ruta del archivo de datos:
```
import os
import time
import pandas as pd
import xlsxwriter
from tools.mapping_mr import *
from tools.apply_style_mr import *
```
## Carga de Datos y Transformaciones
#### Carga los datos desde un archivo Excel y realiza las siguientes transformaciones:

- Relleno de valores nulos.
- Conversión de fechas a datetime.
- Cálculo de columnas adicionales como Días Devolución y Fecha Contractual.
## Procesamiento por Estado
Los datos se dividen en los siguientes grupos:

- Comentados: Estados como "Com. Menores", "Com. Mayores" o "Rechazado".
- Enviados: Documentos marcados como "Enviado".
- Sin Enviar: Documentos sin enviar.
- Aprobado: Documentación finalizada.
## Generación del Archivo Final
Se crea un archivo Excel estilizado donde cada hoja representa un grupo de datos procesados.

```
with pd.ExcelWriter('monitoring_report_' + str(today_date_str) + '.xlsx', engine='xlsxwriter') as writer:
    style_sheet6.to_excel(writer, sheet_name='ALL DOC.', index=False)
    style_sheet_2.to_excel(writer, sheet_name='ENVIADOS', index=False)
    style_sheet_3.to_excel(writer, sheet_name='SIN ENVIAR', index=False)
```

## Cómo Ejecutar el Script
1. Coloca el archivo de datos (data_erp.xlsx) en la ruta especificada.
2. Ejecuta el script con Python:
```
python monitoring_report.py
```
3. Encuentra el archivo generado en la carpeta data.
   
## Notas Adicionales
- Estilización: Los estilos de las celdas se aplican según el estado del documento, utilizando colores definidos.
- Personalización: Puedes ajustar las columnas que se procesan y los colores para adaptarlos a tus necesidades.

## Resultado Final
El archivo Excel resultante incluye todas las métricas y estilos necesarios para un seguimiento detallado de los documentos.
```
Autor: Jose Paredes
```

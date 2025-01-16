# Monitoring Report for ERP

This script processes data extracted from an ERP and organises them in different dataframes according to their status. In addition, it generates a stylised Excel file with several sheets representing different views of the status of the documents.

## Prerequisites

### Necessary Libraries
Make sure to install the following libraries before running the script:

```bash
pip install pandas xlsxwriter openpyxl tqdm aspose-cells
```

## Project Structure
The script uses custom tools and mappings that must be included in the tools folder.
These tools are:

- mapping_mr.py
- apply_style_mr.py

## Script functionalities
#### Main functions
1. Data Import and Data Cleansing
- Loads an Excel file with information from the ERP.
- Performs cleaning and formatting, such as filling null values and converting dates.
2. Processing by Status
- Divides data into different groups (Sent, Not Sent, Commented, etc.).
- Calculates metrics such as return days, contract weeks, and additional notes.
3. Styling and Exporting
- Apply custom styles to dataframes.
- Export data to an Excel file with separate sheets.
  
## Organisation of Excel Sheets
El archivo final incluye las siguientes hojas:

- ALL DOC.: All documents styled according to their status.
- ENVIADOS: Documents in "Sent" status.
- SIN ENVIAR: Documents in "Unsent" status.
- COMENTADOS: Documents with comments ("Minor Com.", "Major Com.", etc.).
- STATUS: General tracking chart.

## Structure of the Code
#### Imports and Initial Configuration
The script imports the necessary libraries and configures the path to the data file:
```
import os
import time
import pandas as pd
import xlsxwriter
from tools.mapping_mr import *
from tools.apply_style_mr import *
```
## Data Loading and Transformations
#### Load the data from an Excel file and perform the following transformations:

- Filling of null values.
- Conversion of dates to datetime.
- Calculation of additional columns such as Return Days and Contract Date.
  
## Processing by State
The data are divided into the following groups:

- Annotated: Statuses such as "Minor Com.", "Major Com." or "Rejected".
- Sent: Documents marked as "Sent".
- Unsent: Documents not sent.
- Approved: Documentation finalised.
## Final File Generation
A stylised Excel file is created where each sheet represents a set of processed data.

```
with pd.ExcelWriter('monitoring_report_' + str(today_date_str) + '.xlsx', engine='xlsxwriter') as writer:
    style_sheet6.to_excel(writer, sheet_name='ALL DOC.', index=False)
    style_sheet_2.to_excel(writer, sheet_name='ENVIADOS', index=False)
    style_sheet_3.to_excel(writer, sheet_name='SIN ENVIAR', index=False)
```

## How to Run the Script
1. Place the data file (data_erp.xlsx) in the specified path.
2. Run the script with Python:
```
python monitoring_report.py
```
3. Find the generated file in the data folder.
   
## Additional Notes
- Styling: Cell styles are applied according to the state of the document, using defined colours.
- Customisation: You can adjust the columns that are processed and the colours to suit your needs.

## Final Result
The resulting Excel file includes all the metrics and styles necessary for detailed document tracking.

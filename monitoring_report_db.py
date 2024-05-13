import pandas as pd
from sqlalchemy import create_engine

# Conexión a la base de datos (reemplaza 'nombre_base_datos' por el nombre de tu base de datos)
engine = create_engine('sqlite:///nombre_base_datos.db')

query_pending = """SELECT * 
                FROM table.table
                WHERE "Estado" LIKE 'Com. Menores' OR "Estado" LIKE 'Com. Mayores' OR "Estado" LIKE 'Rechazado'"""
df = pd.read_sql_query(query_pending, engine)

# Consulta SQL para obtener registros con 'Estado' igual a 'Enviado'
query_under_review = """SELECT *
                    FROM table.table 
                    WHERE "Estado"
                    LIKE 'Enviado'"""
df2 = pd.read_sql_query(query_under_review, engine)

# Consulta SQL para obtener registros con 'Estado' que contenga 'Com. Menores'
query_to_upload = """SELECT *
                    FROM table.table
                    WHERE "Estado"
                    LIKE 'np.nan'"""
df3 = pd.read_sql_query(query_to_upload, engine)

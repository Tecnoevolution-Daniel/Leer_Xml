import pymssql
import json

conn = pymssql.connect(
    server='127.0.0.1',
    user='sa',
    password='SQL#1234',
    database='prueba',
    as_dict=True
)

cursor = conn.cursor()
DML_INSERT = """
    INSERT INTO factura (Empresa, NIT, Factura, IVA, Valor_IVA, Total) VALUES ('PARTNERS TELECOM COLOMBIA S.A.S.', '901354361', 'FCME24528235', '10187.00', '53616.00', '63990.00')
    """

cursor.execute(DML_INSERT)
conn.commit()
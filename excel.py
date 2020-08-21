import pyodbc
import pandas as pd

df = pd.read_excel('C:\\Users\\juli_\\Documents\\prueba.xlsx')
sql_conn = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server}; SERVER=DESKTOP-28IM258\SQLEXPRESS; DATABASE=Dicsys;   Trusted_Connection=yes')

cursor = sql_conn.cursor()
for index, row in df.iterrows():
    print(row)
    cursor.execute("INSERT INTO Articulo([id],[descripcion],[categoria],[cantidad]) values(?,?,?,?)",
        row['id'], row['descripcion'], row['categoria'], row['cantidad'])


sql_conn.commit()
cursor.close()
sql_conn.close()









# arch = 'C:\\Users\\juli_\\Documents\\prueba.xlsx'
# excel= xlrd.open_workbook(arch)
# sheet = excel.sheet_by_name("Hoja1")
# for i in range (sheet.nrows):
#    print(sheet.cell_value(i,0), "        ",sheet.cell_value(i,1),"        ",sheet.cell_value(i,2))




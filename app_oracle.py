# query.py
import csv
import datetime
import os
import pyodbc
import shutil
import cx_Oracle
import pandas as pd
from posixpath import split
from tkinter.ttk import Separator

filePath = 'C:\\'
fileName = 'tempm.csv'

#config de DSN
dsn = cx_Oracle.makedsn("SEU_DSN", port, sid="")

#path do seu instant client 
path = cx_Oracle.init_oracle_client(lib_dir= r"CAMINHO_INSTANT_CLIENT") 

# Conectando no Database
connection = cx_Oracle.connect(user='', password='',
                               dsn = dsn)

# Obtendo cursor consulta
cursor = connection.cursor()

# Sua Query Vai Aqui
sql = """SELECT 
   SYSDATE
   FROM
   DUAL
     """
     
     #Executando a query
cursor.execute(sql)

results = cursor.fetchall()

        # Extraindo Headers
headers = [i[0] for i in cursor.description]

        # Abrindo CSV
with open(filePath + fileName, 'w') as csvFile:

           
            writer = csv.writer(csvFile, delimiter=';', lineterminator='\r',
                                quoting=csv.QUOTE_ALL, escapechar='\\')

            # Gravando Data Headers e Results
            writer.writerow(headers)
            writer.writerows(results)

today = datetime.datetime.now().date()

print("Data Exportada")
connection.close()


    # lendo CSV
text = open("C:\\temp.csv", "r")
  
text = ''.join([i for i in text]) 
  
# replace de caracteres
text = text.replace('"', '') 
text = text.replace(',','.')
text = text.replace(';',',')

  
# Arquivo de Saida
x = open("C:\\temp.csv","w")
x.writelines(text)
x.close()

csv = pd.read_csv("C:\\temp.csv")

excelWriter = pd.ExcelWriter("C:\\final.xlsx")

csv.to_excel(
    
    excelWriter,
    index= False
    
    )

excelWriter.save()

os.remove ("C:\\tempm.csv")
os.remove ("C:\\temp.csv")

quit()



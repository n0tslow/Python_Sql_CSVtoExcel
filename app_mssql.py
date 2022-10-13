import csv
import datetime
import os
import pyodbc
import shutil
from posixpath import split
from tkinter.ttk import Separator
import pandas as pd

# Path do Arquivo
filePath = 'C:\\'
fileName = 'tempm.csv'

# Variavel de conexao
connect = None

if os.path.exists(filePath):

    try:

        # Conexao com database SQL SERVER (instalar Driver de conexao na sua maquina)
        connect = pyodbc.connect(Driver='SQL Server', host='localhost',
                                 database='', user='',
                                 password='')

    except pyodbc.Error as e:
        print("Erro de Conexão com DataBase")
        print(e)
        quit()

    # Chamando Cursor pra executar Query
    cursor = connect.cursor()

    # Sua query aqui
    sqlSelect = \
"""SELECT 
     SYSDATETIME()  
    ,SYSDATETIMEOFFSET()  
    ,SYSUTCDATETIME()  
    ,CURRENT_TIMESTAMP  
    ,GETDATE()  
    ,GETUTCDATE()
"""
    try:

        # Executando a query.
        cursor.execute(sqlSelect)

        # Capturando Dados Retornados na qeury.
        results = cursor.fetchall()

        # Extraindo Headers das Tabelas, linha 0 = headers 
        headers = [i[0] for i in cursor.description]

        # abrindo CSV para escrever os dados
        with open(filePath + fileName, 'w') as csvFile:
  
            writer = csv.writer(csvFile, delimiter=';', lineterminator='\r',
                                quoting=csv.QUOTE_ALL, escapechar='\\')

            # escrevendo os headers e os results
            writer.writerow(headers)
            writer.writerows(results)
            
        print("Arquivo exportado")

    except pyodbc.Error as e:

        print("Erro na exportacao")
        print(e)
        quit()

    finally:

        # Close database connection.
        connect.close()

else:

    print("File path Nao existe")

    quit()

    # abrindo Arquivo Csv para formatar
text = open("tempm.csv", "r")
  
text = ''.join([i for i in text]) 
  
# substituindo caracteres
text = text.replace('"', '') 
text = text.replace(',','.')
text = text.replace(';',',')

  
# Path do arquivo final
x = open("C:\\temp.csv","w")
x.writelines(text)
x.close()

#convertendo o arquivo CSV em Excel XLSX

csv = pd.read_csv("C:\\temp.csv")

excelWriter = pd.ExcelWriter("C:\\final.xlsx")
#caso queira caluna indexa no XLSX manter index= true
csv.to_excel(
    
    excelWriter,
    index= False
    
    )

#salvando
excelWriter.save()

#removendo arquivos temporarios
os.remove ("C:\\temp.csv")
os.remove ("C:\\tempm.csv")

quit()

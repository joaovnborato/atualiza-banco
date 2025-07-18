import pandas as pd
import pymssql
from datetime import datetime

def horario_atual():
    return datetime.now().strftime('%H:%M:%S')

print("Sistema iniciado às", horario_atual())
# 1. Conexão com o banco SQL (no meu caso Azure)
conn = pymssql.connect(
    server='string-de-conexao',
    user='seuLogin',
    password='suaSenha',
    database='nomeDoBanco',
)
print("Conexão efetuada")
cursor = conn.cursor()
inseridos = 0
duplicados = 0
erros = 0

# 2. Ler o arquivo Excel
# DEVOLUÇÃO <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
arquivo = 'devolucao.xlsx'
df = pd.read_excel(arquivo, engine='openpyxl')
inseridos_dev = 0
duplicados_dev = 0
erros_dev = 0
# 3. Inserir linha por linha
for index, row in df.iterrows():
    try:
        cursor.execute("""
            SELECT COUNT(*) FROM Devolucao
            WHERE Data = %s AND ID_Motorista = %s
        """, (row['Data'], row['ID']))
        
        exists = cursor.fetchone()[0]
        
        if exists == 0:
            cursor.execute("""
                INSERT INTO Devolucao (Data, ID_Motorista, Nome_Abrev, Devolucao_Porcentagem)
                VALUES (%s, %s, %s, %s)
            """, (row['Data'], row['ID'], row['Nome'], row['Valor']))
            print(f"[INSERIDO] {row['ID']} - {row['Nome']} - {row['Data']} - {arquivo}")
            inseridos += 1
            inseridos_dev += 1
            
        else:
            print(f"[DUPLICADO - IGNORADO] {row['ID']} - {row['Data']} - {arquivo}")
            duplicados +=1
            duplicados_dev +=1
    
    except Exception as e:
        print(f"[ERRO] {row['ID']} - {row['Valor']} - {e}")
        erros += 1
        erros_dev +=1

conn.commit()
print(f"[INSERIDOS = {inseridos_dev}] - [DUPLICADOS = {duplicados_dev}] - [ERROS = {erros_dev}] - {arquivo}")

# REFUGO <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
arquivo = 'refugo.xlsx'
df = pd.read_excel(arquivo, engine='openpyxl')
inseridos_ref = 0
duplicados_ref = 0
erros_ref = 0
for index, row in df.iterrows():
    try:
        cursor.execute("""
            SELECT COUNT(*) FROM Refugo
            WHERE Data = %s AND ID_Motorista = %s
        """, (row['Data'], row['ID']))
        
        exists = cursor.fetchone()[0]
        
        if exists == 0:
            cursor.execute("""
                INSERT INTO Refugo (Data, ID_Motorista, Nome_Abrev, Refugo_Porcentagem)
                VALUES (%s, %s, %s, %s)
            """, (row['Data'], row['ID'], row['Nome'], row['Valor']))
            print(f"[INSERIDO] {row['ID']} - {row['Nome']} - {row['Data']} - {arquivo}")
            inseridos += 1
            inseridos_ref += 1
        else:
            print(f"[DUPLICADO - IGNORADO] {row['ID']} - {row['Data']} - {arquivo}")
            duplicados += 1
            duplicados_ref += 1
    
    except Exception as e:
        print(f"[ERRO] {row['ID']} - {row['Valor']} - {e}")
        erros += 1
        erros_ref += 1

conn.commit()
print(f"[INSERIDOS = {inseridos_ref}] - [DUPLICADOS = {duplicados_ref}] - [ERROS = {erros_ref}] - {arquivo}")

# RATING <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
arquivo = 'rating.xlsx'
df = pd.read_excel(arquivo, engine='openpyxl')
inseridos_rat = 0
duplicados_rat = 0
erros_rat = 0
for index, row in df.iterrows():
    try:
        cursor.execute("""
            SELECT COUNT(*) FROM Rating
            WHERE Data = %s AND ID_Motorista = %s
        """, (row['Data'], row['ID']))
        
        exists = cursor.fetchone()[0]
        
        if exists == 0:
            cursor.execute("""
                INSERT INTO Rating (Data, ID_Motorista, Nome_Abrev, Rating)
                VALUES (%s, %s, %s, %s)
            """, (row['Data'], row['ID'], row['Nome'], row['Valor']))
            print(f"[INSERIDO] {row['ID']} - {row['Nome']} - {row['Data']} - {arquivo}")
            inseridos += 1
            inseridos_rat += 1
        else:
            print(f"[DUPLICADO - IGNORADO] {row['ID']} - {row['Data']} - {arquivo}")
            duplicados += 1
            duplicados_rat += 1
    
    except Exception as e:
        print(f"[ERRO] {row['ID']} - {row['Valor']} - {e}")
        erros += 1
        erros_rat += 1

conn.commit()
print(f"[INSERIDOS = {inseridos_rat}] - [DUPLICADOS = {duplicados_rat}] - [ERROS = {erros_rat}] - {arquivo}")

# DISPERSÃO KM <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
arquivo = 'dispersao.xlsx'
df = pd.read_excel(arquivo, engine='openpyxl')
inseridos_dis = 0
duplicados_dis = 0
erros_dis = 0
for index, row in df.iterrows():
    try:
        cursor.execute("""
            SELECT COUNT(*) FROM Dispersao
            WHERE Data = %s AND ID_Motorista = %s
        """, (row['Data'], row['ID']))
        
        exists = cursor.fetchone()[0]
        
        if exists == 0:
            cursor.execute("""
                INSERT INTO Dispersao (Data, ID_Motorista, Nome_Abrev, Dispersao_KM)
                VALUES (%s, %s, %s, %s)
            """, (row['Data'], row['ID'], row['Nome'], row['Valor']))
            print(f"[INSERIDO] {row['ID']} - {row['Nome']} - {row['Data']} - {arquivo}")
            inseridos += 1
            inseridos_dis += 1
        else:
            print(f"[DUPLICADO - IGNORADO] {row['ID']} - {row['Data']} - {arquivo}")
            duplicados += 1
            duplicados_dis += 1
    
    except Exception as e:
        print(f"[ERRO] {row['ID']} - {row['Valor']} - {e}")
        erros += 1
        erros_dis += 1

conn.commit()
print(f"[INSERIDOS = {inseridos_dis}] - [DUPLICADOS = {duplicados_dis}] - [ERROS = {erros_dis}] - {arquivo}")

# REPOSIÇÃO <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
arquivo = 'reposicao.xlsx'
df = pd.read_excel(arquivo, engine='openpyxl')
inseridos_rep = 0
duplicados_rep = 0
erros_rep = 0
for index, row in df.iterrows():
    try:
        cursor.execute("""
            SELECT COUNT(*) FROM Reposicao
            WHERE Data = %s AND ID_Motorista = %s
        """, (row['Data'], row['ID']))
        
        exists = cursor.fetchone()[0]
        
        if exists == 0:
            cursor.execute("""
                INSERT INTO Reposicao (Data, ID_Motorista, Nome_Abrev, Reposicao_Valor)
                VALUES (%s, %s, %s, %s)
            """, (row['Data'], row['ID'], row['Nome'], row['Valor']))
            print(f"[INSERIDO] {row['ID']} - {row['Nome']} - {row['Data']} - {arquivo}")
            inseridos += 1
            inseridos_rep += 1
        else:
            print(f"[DUPLICADO - IGNORADO] {row['ID']} - {row['Data']} - {arquivo}")
            duplicados += 1
            duplicados_rep += 1
    except Exception as e:
        print(f"[ERRO] {row['ID']} - {row['Valor']} - {e}" - {arquivo})
        erros += 1
        erros_rep += 1

# 4. Commit
conn.commit()
print(f"[INSERIDOS = {inseridos_rep}] - [DUPLICADOS = {duplicados_rep}] - [ERROS = {erros_rep}] - {arquivo}")
cursor.close()
conn.close()
agora = datetime.now()

print("-----------------------------------------RESUMO-------------------------------------------")
print("Sistema encerrado às", horario_atual())
print(f"TOTAL     >>>>> INSERIDOS = {inseridos} - DUPLICADOS = {duplicados} - ERROS = {erros}")
print(" ")
print(f"DEVOLUÇÃO >>>>> {inseridos_dev} inseridos, {duplicados_dev} duplicados, e {erros_dev} erros.")
print(f"REFUGO    >>>>> {inseridos_ref} inseridos, {duplicados_ref} duplicados, e {erros_ref} erros.")
print(f"RATING    >>>>> {inseridos_rat} inseridos, {duplicados_rat} duplicados, e {erros_rat} erros.")
print(f"DISPERSÃO >>>>> {inseridos_dis} inseridos, {duplicados_dis} duplicados, e {erros_dis} erros.")
print(f"REPOSIÇÃO >>>>> {inseridos_rep} inseridos, {duplicados_rep} duplicados, e {erros_rep} erros.")



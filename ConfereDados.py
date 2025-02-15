# Conexão Sheets
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Planilhas
from openpyxl import load_workbook
import pandas as pd

# Arquivos
import os
from pathlib import Path

# CONEXÃO COM GOOGLE SHEETS
filename = "ARQUIVO_G_SHEET.json" # ARQUIVO GERADO NO GOOGLE CLOUD PELAS CONTAS DE SERVIÇO

    # PADRÃO
scopes = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

    # PADRÃO
creds = ServiceAccountCredentials.from_json_keyfile_name(filename=filename,scopes=scopes)
client = gspread.authorize(creds)


titulo = "NOME_PLANILHA" # NOME DO SHEET
folder_id = "ID_PASTA_DRIVE" # ID DA PASTA NO GOOGLE DRIVE

    # PADRÃO
planilhaCompleta = client.open(title = titulo, folder_id = folder_id)

origem = r"C:\Users\Padrão\Desktop\BeL Tech\VJ\Royal\Planilhas"
for caminho, subpasta, arquivos in os.walk(origem):
    for contador,(nome) in enumerate(arquivos): # UMA EXECUÇÃO POR ARQUIVO
        planilha = planilhaCompleta.get_worksheet(contador)
        print(contador)
        dados = planilha.get_values()
        df = pd.DataFrame(dados).applymap(str)
        arq = caminho+"\\"+nome
        df_planilha = pd.read_excel(arq,header=None)
        df_planilha.fillna("",inplace=True)
        df_planilha.applymap(str)
        print(df)
        print(df_planilha)
        print(df.equals(df_planilha))

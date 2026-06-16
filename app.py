import streamlit as st
import pandas as pd
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import re
import unicodedata
import time

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Central de Relatórios WLM", layout="wide", page_icon="🔒")
ID_PLANILHA_MESTRA = "1XibBlm2x46Dk5bf4JvfrMepD4gITdaOtTALSgaFcwV0"

# --- AUXILIARES ---
def remover_acentos(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

def conectar_sheets():
    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    credentials_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(credentials_dict, scopes=scope)
    client = gspread.authorize(creds)
    return client

def converter_br_para_float(valor):
    if pd.isna(valor) or valor == "": 
        return 0.0
    if isinstance(valor, (int, float)): 
        return float(valor)
    valor_str = str(valor).strip()
    valor_str = valor_str.replace('\xa0', '').replace('R$', '').strip()
    if not valor_str:
        return 0.0
    if '.' in valor_str and ',' in valor_str: 
        valor_str = valor_str.replace('.', '')
    valor_str = valor_str.replace(',', '.')
    try: 
        return float(valor_str)
    except: 
        return 0.0

def padronizar_data_quatro_digitos(data_str):
    if pd.isna(data_str) or data_str == "":
        return ""
    data_str = str(data_str).strip()
    if '/' in data_str:
        partes = data_str.split('/')
        if len(partes) == 3:
            dia, mes, ano = partes
            if len(ano) == 2:
                ano = '20' + ano
            return f"{dia.zfill(2)}/{mes.zfill(2)}/{ano}"
    return data_str

def verificar_acesso():
    try:
        client = conectar_sheets()
        sh = client.open_by_key(ID_PLANILHA_MESTRA)
        try: return sh.worksheet("Config").acell('B1').value
        except: return 'admin'
    except: return None

def obter_ultima_atualizacao():
    try:
        client = conectar_sheets()
        sh = client.open_by_key(ID_PLANILHA_MESTRA)
        ws = sh.worksheet("Consolidado")
        datas = ws.col_values(1)[1:] 
        if not datas: return "Nenhuma"
        datas_dt = pd.to_datetime(datas, dayfirst=True, errors='coerce')
        ultima_data = datas_dt.max().strftime('%d/%m/%Y')
        return ultima_data
    except:
        return "Erro ao ler"

# --- PARSERS (LEITURA) ---
def parse_comissoes(arquivos):
    dados = []
    for arquivo in arquivos:
        try:
            arquivo.seek(0)
            try: conteudo = arquivo.read().decode("utf-8")
            except: 
                arquivo.seek(0)
                conteudo = arquivo.read().decode("latin-1")
            soup = BeautifulSoup(conteudo, "html.parser")
            texto_completo = soup.get_text(separator=" ", strip=True)
            match_data = re.search(r"até\s+(\d{2}/\d{2}/\d{4})", texto_completo, re.IGNORECASE)
            data_relatorio = match_data.group(1) if match_data else datetime.now().strftime("%d/%m/%Y")
            tecnico_atual = None
            for linha in soup.find_all("tr"):
                texto_linha = linha.get_text(separator=" ", strip=True).upper()
                if "TOTAL DA FILIAL" in texto_linha or "TOTAL DA EMPRESA" in texto_linha: break
                if "TOTAL DO FUNCIONARIO" in texto_linha:
                    try: tecnico_atual = texto_linha.split("TOTAL DO FUNCIONARIO")[1].replace(":", "").strip().split()[0]
                    except: continue 
                if tecnico_atual and "HORAS VENDIDAS:" in texto_linha:
                    celulas = linha.find_all("td")
                    for celula in celulas:
                        txt = celula.get_text(strip=True).upper()
                        if "HORAS" in txt and any(c.isdigit() for c in txt) and "VENDIDAS" not in txt:
                            valor_limpo = txt.replace("HORAS", "").strip()
                            dados.append([data_relatorio, arquivo.name, tecnico_atual, valor_limpo])
                            break 
        except Exception as e: st.error(f"Erro no arquivo {arquivo.name}: {e}")
    return dados

def parse_aproveitamento(arquivos):
    dados = []
    for arquivo in arquivos:
        try:
            arquivo.seek(0)
            try: conteudo = arquivo.read().decode("utf-8")
            except:
                try: conteudo = arquivo.read().decode("latin-1")
                except: conteudo = arquivo.read().decode("utf-16")
            soup = BeautifulSoup(conteudo, "html.parser")
            tecnico_atual_aprov = None
            linhas = soup.find_all("tr")
            for linha in linhas:
                texto_original = linha.get_text(separator=" ", strip=True).upper()
                texto_limpo = remover_acentos(texto_original)
                if "TOTAL FILIAL:" in texto_original: break
                if "MECANICO" in texto_limpo and "TOT.MEC" not in texto_limpo:
                    try:
                        parte_direita = texto_limpo.split("MECANICO")[1].replace(":", "").strip()
                        if "-" in parte_direita: tecnico_atual_aprov = parte_direita.split("-")[0].strip()
                        else: tecnico_atual_aprov = parte_direita.split()[0]
                    except: continue
                if "TOT.MEC.:" in texto_original: tecnico_atual_aprov = None; continue
                if tecnico_atual_aprov:
                    celulas = linha.find_all("td")
                    if not celulas: continue
                    txt_cel0 = celulas[0].get_text(strip=True)
                    if re.match(r"\d{2}/\d{2}/\d{2}", txt_cel0):
                        try:
                            if len(celulas) >= 4:
                                dados.append([
                                    txt_cel0.split()[0], 
                                    arquivo.name, 
                                    tecnico_atual_aprov, 
                                    

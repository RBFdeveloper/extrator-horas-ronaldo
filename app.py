import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from gspread_dataframe import set_with_dataframe
from datetime import datetime, timedelta
import re
import html

# ===================== CONFIGURAÇÃO DO GOOGLE =====================
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

SPREADSHEET_ID = "1KcbT6HBYf0b-7qLI70T5Xon7Xv7yqj-7qcIBlWFbHnw"

def conectar_google():
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=SCOPES)
    client = gspread.authorize(creds)
    return client, creds

@st.cache_resource
def get_drive_service():
    _, creds = conectar_google()
    return build('drive', 'v3', credentials=creds)

@st.cache_resource
def get_sheets_client():
    client, _ = conectar_google()
    return client

# ===================== PARSERS DE HTML =====================
def parse_html_table(html_text):
    linhas = html_text.strip().split('\n')
    dados = []
    for linha in linhas:
        linha = linha.strip()
        if not linha.startswith('<tr>') and not linha.startswith('<TR>'):
            continue
        valores = re.findall(r'<td>(.*?)</td>', linha, flags=re.IGNORECASE)
        valores = [html.unescape(v) for v in valores]
        if valores:
            dados.append(valores)
    if not dados:
        return pd.DataFrame()
    colunas = dados[0]
    dados = dados[1:]
    if len(dados) > 0 and len(dados[0]) != len(colunas):
        colunas = [f"Col_{i+1}" for i in range(len(dados[0]))]
    return pd.DataFrame(dados, columns=colunas)

def parse_html_table_com_cabecalho(html_text, cabecalho_esperado):
    df = parse_html_table(html_text)
    if df.empty:
        return pd.DataFrame(columns=cabecalho_esperado)
    df.columns = [str(c).strip() for c in df.columns]
    for col in cabecalho_esperado:
        if col not in df.columns:
            df[col] = None
    return df[cabecalho_esperado]

# ===================== FUNÇÕES AUXILIARES =====================
def get_sheet_as_df(sheet_name, header_row=1):
    try:
        client = get_sheets_client()
        sheet = client.open_by_key(SPREADSHEET_ID).worksheet(sheet_name)
        dados = sheet.get_all_records(head=header_row)
        return pd.DataFrame(dados)
    except Exception as e:
        st.error(f"Erro ao ler a aba '{sheet_name}': {e}")
        return pd.DataFrame()

def clear_sheet(sheet_name):
    client = get_sheets_client()
    sheet = client.open_by_key(SPREADSHEET_ID).worksheet(sheet_name)
    sheet.clear()

def append_to_sheet(sheet_name, df, header=True):
    client = get_sheets_client()
    sheet = client.open_by_key(SPREADSHEET_ID).worksheet(sheet_name)
    if header:
        sheet.append_row(df.columns.tolist())
    for _, row in df.iterrows():
        sheet.append_row(row.astype(str).tolist())

def salvar_df_no_sheet(sheet_name, df, header=True):
    client = get_sheets_client()
    planilha = client.open_by_key(SPREADSHEET_ID)
    try:
        worksheet = planilha.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = planilha.add_worksheet(title=sheet_name, rows=1000, cols=50)
    worksheet.clear()
    if df.empty:
        return
    set_with_dataframe(worksheet, df, include_index=False, include_column_header=header)

# ===================== UNIFICAÇÃO COMISSES X APROVEITAMENTO =====================
def processar_unificacao(df_comissoes, df_aproveitamento):
    if df_comissoes.empty:
        st.warning("Aba 'Comissoes' está vazia. Nada para unificar.")
        return pd.DataFrame()

    colunas_comissoes = ['Data', 'Sigla', 'OS', 'Cliente', 'Servico', 'Local', 'Valor', 'Comissao', 'Observacao', 'Comissao Real']
    for col in colunas_comissoes:
        if col not in df_comissoes.columns:
            df_comissoes[col] = None

    df_comissoes = df_comissoes[colunas_comissoes].copy()
    df_comissoes['Sigla'] = df_comissoes['Sigla'].astype(str).str.strip().str.upper()
    df_comissoes['Data'] = pd.to_datetime(df_comissoes['Data'], dayfirst=True, errors='coerce')

    df_comissoes['Disp'] = 0
    df_comissoes['TP'] = 0
    df_comissoes['TG'] = 0

    if not df_aproveitamento.empty:
        colunas_aproveitamento = ['Data', 'Técnico', 'Disp', 'TP', 'TG']
        for col in colunas_aproveitamento:
            if col not in df_aproveitamento.columns:
                df_aproveitamento[col] = 0
        df_aproveitamento = df_aproveitamento[colunas_aproveitamento].copy()
        df_aproveitamento['Técnico'] = df_aproveitamento['Técnico'].astype(str).str.strip().str.upper()
        df_aproveitamento['Data'] = pd.to_datetime(df_aproveitamento['Data'], dayfirst=True, errors='coerce')

        for col in ['Disp', 'TP', 'TG']:
            df_aproveitamento[col] = pd.to_numeric(df_aproveitamento[col], errors='coerce').fillna(0)

        df_aproveitamento = df_aproveitamento.groupby(['Data', 'Técnico'], as_index=False).agg({
            'Disp': 'sum',
            'TP': 'sum',
            'TG': 'sum'
        })

        df_comissoes = df_comissoes.merge(
            df_aproveitamento,
            left_on=['Data', 'Sigla'],
            right_on=['Data', 'Técnico'],
            how='left'
        )
        df_comissoes.drop(columns=['Técnico'], inplace=True, errors='ignore')
    else:
        st.warning("Aba 'Aproveitamento' está vazia. Unificação sem dados de aproveitamento.")

    for col in ['Disp', 'TP', 'TG']:
        df_comissoes[col] = pd.to_numeric(df_comissoes[col], errors='coerce').fillna(0)

    df_comissoes['Data'] = df_comissoes['Data'].dt.strftime('%d/%m/%Y')

    colunas_finais = ['Data', 'Sigla', 'OS', 'Cliente', 'Servico', 'Local', 'Valor', 'Comissao', 'Observacao', 'Comissao Real', 'Disp', 'TP', 'TG']
    df_comissoes = df_comissoes[[c for c in colunas_finais if c in df_comissoes.columns]]
    return df_comissoes

# ===================== INTERFACE STREAMLIT =====================
st.set_page_config(page_title="Unificador de Dados", layout="wide")
st.title("Unificador de Comissões e Aproveitamento")

aba_origem = st.sidebar.selectbox("Aba de origem", ["Comissoes", "Aproveitamento", "Unificacao"])

if aba_origem in ["Comissoes", "Aproveitamento"]:
    st.subheader(f"Importar dados para a aba: {aba_origem}")
    html_input = st.text_area("Cole aqui o HTML da tabela", height=300)
    if st.button(f"Importar {aba_origem}"):
        if html_input.strip():
            if aba_origem == "Comissoes":
                cabecalho = ['Data', 'Sigla', 'OS', 'Cliente', 'Servico', 'Local', 'Valor', 'Comissao', 'Observacao', 'Comissao Real']
            else:
                cabecalho = ['Data', 'Técnico', 'Disp', 'TP', 'TG']
            df = parse_html_table_com_cabecalho(html_input, cabecalho)
            if not df.empty:
                salvar_df_no_sheet(aba_origem, df, header=True)
                st.success(f"{len(df)} registros importados para '{aba_origem}'.")
                st.dataframe(df)
            else:
                st.warning("Nenhum dado válido encontrado no HTML.")
        else:
            st.warning("Por favor, cole o HTML antes de importar.")

elif aba_origem == "Unificacao":
    st.subheader("Unificar Comissões + Aproveitamento")
    if st.button("Executar Unificação"):
        df_comissoes = get_sheet_as_df("Comissoes")
        df_aproveitamento = get_sheet_as_df("Aproveitamento")
        df_unificado = processar_unificacao(df_comissoes, df_aproveitamento)
        if not df_unificado.empty:
            salvar_df_no_sheet("Unificacao", df_unificado, header=True)
            st.success(f"Unificação concluída! {len(df_unificado)} registros salvos.")
            st.dataframe(df_unificado)
        else:
            st.warning("Nenhum registro unificado para salvar.")

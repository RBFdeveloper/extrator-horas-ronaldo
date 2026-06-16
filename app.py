import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime, date

# -----------------------------------------------------------------------------
# CONFIGURAÇÕES
# -----------------------------------------------------------------------------
SPREADSHEET_ID = 'SUA_PLANILHA_AQUI'
ABA_COMISSOES = 'Comissoes'
ABA_APROVEITAMENTO = 'Aproveitamento'
ABA_CONSOLIDADO = 'Consolidado'
CREDENTIALS_FILE = 'credentials.json'

# -----------------------------------------------------------------------------
# AUTENTICAÇÃO
# -----------------------------------------------------------------------------
def autenticar_google_sheets():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = service_account.Credentials.from_service_account_file(
        CREDENTIALS_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    return service

# -----------------------------------------------------------------------------
# LEITURA DO GOOGLE SHEETS
# -----------------------------------------------------------------------------
def ler_aba(service, spreadsheet_id, aba):
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=aba
        ).execute()
        values = result.get('values', [])
        if not values:
            return pd.DataFrame()

        header = values[0]
        dados = values[1:]
        df = pd.DataFrame(dados, columns=header)
        return df
    except HttpError as e:
        print(f'Erro ao ler a aba {aba}: {e}')
        return pd.DataFrame()

# -----------------------------------------------------------------------------
# ESCREVER NO GOOGLE SHEETS
# -----------------------------------------------------------------------------
def escrever_aba(service, spreadsheet_id, aba, df):
    try:
        service.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id,
            range=aba,
            body={}
        ).execute()

        values = [df.columns.tolist()] + df.values.tolist()
        body = {'values': values}

        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=aba,
            valueInputOption='RAW',
            body=body
        ).execute()
        print(f'Aba {aba} atualizada com sucesso.')
    except HttpError as e:
        print(f'Erro ao escrever na aba {aba}: {e}')

# -----------------------------------------------------------------------------
# LIMPEZA DE NOMES / SIGLAS
# -----------------------------------------------------------------------------
def limpar_texto(coluna):
    if coluna is None or not isinstance(coluna, pd.Series):
        return coluna
    return (
        coluna
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace(r'\s+', ' ', regex=True)
    )

# -----------------------------------------------------------------------------
# CONVERSÃO DE DATAS
# -----------------------------------------------------------------------------
def converter_para_datetime(coluna, dayfirst=True):
    if coluna is None or not isinstance(coluna, pd.Series):
        return coluna
    return pd.to_datetime(coluna, dayfirst=dayfirst, errors='coerce')

# -----------------------------------------------------------------------------
# FORMATAÇÃO PARA O BI / GOOGLE SHEETS
# -----------------------------------------------------------------------------
def formatar_data(df, coluna_data='DATA'):
    if coluna_data in df.columns:
        df[coluna_data] = pd.to_datetime(df[coluna_data], dayfirst=True, errors='coerce')
        df[coluna_data] = df[coluna_data].dt.strftime('%d/%m/%Y')
    return df

# -----------------------------------------------------------------------------
# AJUSTE DE ESCALA NUMÉRICA
# -----------------------------------------------------------------------------
def ajustar_escala(df, colunas, divisao=100):
    for col in colunas:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace('.', '', regex=False)
                .str.replace(',', '.', regex=False)
            )
            df[col] = pd.to_numeric(df[col], errors='coerce')
            if divisao:
                df[col] = df[col] / divisao
    return df

# -----------------------------------------------------------------------------
# UNIFICAÇÃO DE DADOS
# -----------------------------------------------------------------------------
def processar_unificacao(df_comissoes, df_aproveitamento):
    if df_comissoes.empty or df_aproveitamento.empty:
        raise ValueError('Uma ou ambas as abas de origem estão vazias.')

    col_tecnico_com = 'TECNICO'
    col_tecnico_apr = 'TECNICO'
    col_data_com = 'DATA'
    col_data_apr = 'DATA'

    df_comissoes[col_tecnico_com] = limpar_texto(df_comissoes[col_tecnico_com])
    df_aproveitamento[col_tecnico_apr] = limpar_texto(df_aproveitamento[col_tecnico_apr])

    df_comissoes[col_data_com] = converter_para_datetime(df_comissoes[col_data_com])
    df_aproveitamento[col_data_apr] = converter_para_datetime(df_aproveitamento[col_data_apr])

    df_comissoes = df_comissoes.dropna(subset=[col_tecnico_com, col_data_com])
    df_aproveitamento = df_aproveitamento.dropna(subset=[col_tecnico_apr, col_data_apr])

    df_unificado = pd.merge(
        df_comissoes,
        df_aproveitamento,
        how='outer',
        left_on=[col_tecnico_com, col_data_com],
        right_on=[col_tecnico_apr, col_data_apr],
        suffixes=('_COM', '_APR')
    )

    if col_data_com in df_unificado.columns and col_data_apr in df_unificado.columns:
        df_unificado[col_data_com] = df_unificado[col_data_com].fillna(df_unificado[col_data_apr])
        df_unificado = df_unificado.drop(columns=[col_data_apr])
        df_unificado = df_unificado.rename(columns={col_data_com: 'DATA'})
    elif col_data_apr in df_unificado.columns:
        df_unificado = df_unificado.rename(columns={col_data_apr: 'DATA'})
    elif col_data_com in df_unificado.columns:
        df_unificado = df_unificado.rename(columns={col_data_com: 'DATA'})

    if col_tecnico_com in df_unificado.columns and col_tecnico_apr in df_unificado.columns:
        df_unificado[col_tecnico_com] = df_unificado[col_tecnico_com].fillna(df_unificado[col_tecnico_apr])
        df_unificado = df_unificado.drop(columns=[col_tecnico_apr])
        df_unificado = df_unificado.rename(columns={col_tecnico_com: 'TECNICO'})
    elif col_tecnico_apr in df_unificado.columns:
        df_unificado = df_unificado.rename(columns={col_tecnico_apr: 'TECNICO'})
    elif col_tecnico_com in df_unificado.columns:
        df_unificado = df_unificado.rename(columns={col_tecnico_com: 'TECNICO'})

    ultima_data_processada = df_unificado['DATA'].max()
    if pd.isna(ultima_data_processada):
        ultima_data_processada = None
    else:
        ultima_data_processada = pd.Timestamp(ultima_data_processada).to_pydatetime()

    df_unificado = df_unificado.sort_values(by=['DATA', 'TECNICO'], ascending=[False, True])

    colunas_numericas = [col for col in df_unificado.columns if any(
        palavra in col.upper() for palavra in ['VALOR', 'COMISSAO', 'RECEITA', 'TOTAL', 'QTD', 'QUANTIDADE'])]
    df_unificado = ajustar_escala(df_unificado, colunas_numericas, divisao=100)

    df_unificado = formatar_data(df_unificado, 'DATA')

    df_unificado = df_unificado.reset_index(drop=True)

    return df_unificado, ultima_data_processada

# -----------------------------------------------------------------------------
# EXECUÇÃO PRINCIPAL
# -----------------------------------------------------------------------------
def main():
    service = autenticar_google_sheets()

    df_comissoes = ler_aba(service, SPREADSHEET_ID, ABA_COMISSOES)
    df_aproveitamento = ler_aba(service, SPREADSHEET_ID, ABA_APROVEITAMENTO)

    df_consolidado, ultima_data = processar_unificacao(df_comissoes, df_aproveitamento)

    escrever_aba(service, SPREADSHEET_ID, ABA_CONSOLIDADO, df_consolidado)

    if ultima_data:
        print(f'Última data processada: {ultima_data.strftime("%d/%m/%Y")}')
    else:
        print('Nenhuma data válida encontrada no consolidado.')


if __name__ == '__main__':
    main()

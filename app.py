import streamlit as st
import pandas as pd
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import re
import unicodedata
import time
import traceback

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
        try:
            return sh.worksheet("Config").acell('B1').value
        except:
            return 'admin'
    except:
        return None

def obter_ultima_atualizacao():
    try:
        client = conectar_sheets()
        sh = client.open_by_key(ID_PLANILHA_MESTRA)
        ws = sh.worksheet("Consolidado")
        datas = ws.col_values(1)[1:]
        if not datas:
            return "Nenhuma"
        datas_dt = pd.to_datetime(datas, dayfirst=True, errors='coerce')
        ultima_data = datas_dt.max().strftime('%d/%m/%Y')
        return ultima_data
    except:
        return "Erro ao ler"

# --- PARSERS ---
def parse_comissoes(arquivos):
    dados = []
    for arquivo in arquivos:
        try:
            arquivo.seek(0)
            try:
                conteudo = arquivo.read().decode("utf-8")
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
                if "TOTAL DA FILIAL" in texto_linha or "TOTAL DA EMPRESA" in texto_linha:
                    break
                if "TOTAL DO FUNCIONARIO" in texto_linha:
                    try:
                        tecnico_atual = texto_linha.split("TOTAL DO FUNCIONARIO")[1].replace(":", "").strip().split()[0]
                    except:
                        continue
                if tecnico_atual and "HORAS VENDIDAS:" in texto_linha:
                    celulas = linha.find_all("td")
                    for celula in celulas:
                        txt = celula.get_text(strip=True).upper()
                        if "HORAS" in txt and any(c.isdigit() for c in txt) and "VENDIDAS" not in txt:
                            valor_limpo = txt.replace("HORAS", "").strip()
                            dados.append([data_relatorio, arquivo.name, tecnico_atual, valor_limpo])
                            break
        except Exception as e:
            st.error(f"Erro no arquivo {arquivo.name}: {e}")
    return dados

def parse_aproveitamento(arquivos):
    dados = []
    for arquivo in arquivos:
        try:
            arquivo.seek(0)
            try:
                conteudo = arquivo.read().decode("utf-8")
            except:
                try:
                    conteudo = arquivo.read().decode("latin-1")
                except:
                    conteudo = arquivo.read().decode("utf-16")

            soup = BeautifulSoup(conteudo, "html.parser")
            tecnico_atual_aprov = None
            linhas = soup.find_all("tr")
            for linha in linhas:
                texto_original = linha.get_text(separator=" ", strip=True).upper()
                texto_limpo = remover_acentos(texto_original)

                if "TOTAL FILIAL:" in texto_original:
                    break
                if "MECANICO" in texto_limpo and "TOT.MEC" not in texto_limpo:
                    try:
                        parte_direita = texto_limpo.split("MECANICO")[1].replace(":", "").strip()
                        if "-" in parte_direita:
                            tecnico_atual_aprov = parte_direita.split("-")[0].strip()
                        else:
                            tecnico_atual_aprov = parte_direita.split()[0]
                    except:
                        continue

                if "TOT.MEC.:" in texto_original:
                    tecnico_atual_aprov = None
                    continue

                if tecnico_atual_aprov:
                    celulas = linha.find_all("td")
                    if not celulas:
                        continue
                    txt_cel0 = celulas[0].get_text(strip=True)
                    if re.match(r"\d{2}/\d{2}/\d{2}", txt_cel0):
                        try:
                            if len(celulas) >= 4:
                                dados.append([
                                    txt_cel0.split()[0],
                                    arquivo.name,
                                    tecnico_atual_aprov,
                                    celulas[1].get_text(strip=True),
                                    celulas[2].get_text(strip=True),
                                    celulas[3].get_text(strip=True)
                                ])
                        except:
                            continue
        except Exception as e:
            st.error(f"Erro no arquivo {arquivo.name}: {e}")
    return dados

# --- GRAVAÇÃO ---
def atualizar_planilha_preservando_formato(sh, nome_aba, df_final):
    try:
        ws = sh.worksheet(nome_aba)
    except:
        ws = sh.add_worksheet(title=nome_aba, rows=2000, cols=20)

    if not ws.get_all_values():
        ws.update('A1', [df_final.columns.values.tolist()])
        try:
            ws.format('A1:Z1', {'textFormat': {'bold': True}})
        except:
            pass

    ws.batch_clear(["A2:Z10000"])
    df_final = df_final.fillna(0.0)
    dados_para_enviar = df_final.values.tolist()
    if dados_para_enviar:
        ws.update('A2', dados_para_enviar)
    return True

# --- UPSERT ---
def salvar_com_upsert(nome_aba, novos_dados_df, colunas_chaves):
    client = conectar_sheets()
    sh = client.open_by_key(ID_PLANILHA_MESTRA)

    try:
        time.sleep(1)
        ws = sh.worksheet(nome_aba)
        dados_antigos = ws.get_all_records()
        df_antigo = pd.DataFrame(dados_antigos)
    except Exception as e:
        df_antigo = pd.DataFrame()

    for col in colunas_chaves:
        if col in novos_dados_df.columns:
            novos_dados_df[col] = novos_dados_df[col].astype(str).str.strip().str.upper()
        if not df_antigo.empty and col in df_antigo.columns:
            df_antigo[col] = df_antigo[col].astype(str).str.strip().str.upper()

    if not df_antigo.empty:
        col_data = 'Data Processamento' if 'Data Processamento' in novos_dados_df.columns else 'Data'
        datas_novas = novos_dados_df[col_data].unique()
        df_final = df_antigo[~df_antigo[col_data].isin(datas_novas)]
        df_final = pd.concat([df_final, novos_dados_df], ignore_index=True)
    else:
        df_final = novos_dados_df

    df_final = df_final.fillna(0.0)
    time.sleep(1.5)
    atualizar_planilha_preservando_formato(sh, nome_aba, df_final)
    return len(df_final)

# --- SALVAR AJUSTE MANUAL ---
def salvar_ajuste_manual(data, tecnico, metrica, valor, motivo):
    client = conectar_sheets()
    sh = client.open_by_key(ID_PLANILHA_MESTRA)
    try:
        ws = sh.worksheet("Ajustes")
    except:
        ws = sh.add_worksheet(title="Ajustes", rows=1000, cols=10)
        ws.append_row(["Data", "Técnico", "Métrica", "Valor", "Motivo", "Data do Registro"])

    ws.append_row([
        str(data.strftime('%d/%m/%Y')),
        tecnico,
        metrica,
        float(valor),
        motivo,
        datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    ])

# --- APLICAR AJUSTES ---
def aplicar_logica_ajustes(df_base):
    try:
        client = conectar_sheets()
        sh = client.open_by_key(ID_PLANILHA_MESTRA)
        ws_ajustes = sh.worksheet("Ajustes")
        dados_ajustes = ws_ajustes.get_all_records()

        if not dados_ajustes:
            return df_base

        df_ajustes = pd.DataFrame(dados_ajustes)

        mapa = {
            "Horas Vendidas (HV)": "Horas Vendidas",
            "Tempo Padrão (TP)": "TP",
            "Tempo Disponível (Disp)": "Disp",
            "Tempo Garantia (TG)": "TG"
        }

        df_base['Key_D_Comp'] = pd.to_datetime(df_base['Data'], dayfirst=True, errors='coerce')

        for _, row in df_ajustes.iterrows():
            try:
                dt_ajuste = pd.to_datetime(row['Data'], dayfirst=True, errors='coerce')
                tec_ajuste = str(row['Técnico']).strip()
                metrica_ajuste = mapa.get(row['Métrica'])
                valor_ajuste = float(str(row['Valor']).replace(',', '.'))

                if metrica_ajuste and metrica_ajuste in df_base.columns:
                    mask = (df_base['Key_D_Comp'] == dt_ajuste) & (df_base['Técnico'] == tec_ajuste)
                    if mask.any():
                        df_base.loc[mask, metrica_ajuste] += valor_ajuste
            except:
                continue

        if 'Key_D_Comp' in df_base.columns:
            df_base.drop(columns=['Key_D_Comp'], inplace=True)

        return df_base

    except Exception as e:
        print(f"Erro ajustes: {e}")
        return df_base

# --- TRADUZIR NOMES ---
def aplicar_traducao_nomes(df_final):
    try:
        client = conectar_sheets()
        sh = client.open_by_key(ID_PLANILHA_MESTRA)

        try:
            ws_nomes = sh.worksheet("Nomes")
            todas_linhas = ws_nomes.get_all_values()

            dicionario_nomes = {}
            for row in todas_linhas[1:]:
                if len(row) >= 2:
                    sigla = str(row[0]).strip().upper()
                    nome = str(row[1]).strip()
                    if sigla and nome:
                        dicionario_nomes[sigla] = nome

            if dicionario_nomes:
                df_final['Técnico'] = df_final['Técnico'].apply(
                    lambda sigla: dicionario_nomes.get(str(sigla).strip().upper(), sigla)
                )
                print(f"Tradução aplicada: {len(dicionario_nomes)} nomes encontrados.")

        except Exception as e:
            print(f"Aba 'Nomes' não lida: {e}")

        return df_final

    except Exception as e:
        print(f"Erro na tradução de nomes: {e}")
        return df_final

# --- UNIFICAÇÃO (CORRIGIDA) ---
def processar_unificacao():
    """
    CORREÇÕES APLICADAS:
    1. fillna(0.0) removido antes do combine_first — não envenena mais as chaves
    2. replace() com regex=True para limpar '0', '0.0', 'nan' corretamente
    3. dropna substituído por filtro explícito — não remove linhas legítimas
    """
    try:
        client = conectar_sheets()
        sh = client.open_by_key(ID_PLANILHA_MESTRA)
        ws_com = sh.worksheet("Comissoes")
        ws_aprov = sh.worksheet("Aproveitamento")

        dados_com = ws_com.get_all_records()
        dados_aprov = ws_aprov.get_all_records()

        # Encerra só se AMBAS as abas estiverem vazias
        if not dados_com and not dados_aprov:
            print("Abas Comissoes e Aproveitamento vazias. Nada a unificar.")
            return False

        # Cria DataFrames mesmo se uma das abas vier vazia
        df_com = (
            pd.DataFrame(dados_com)
            if dados_com
            else pd.DataFrame(columns=['Data', 'Técnico', 'Horas Vendidas'])
        )
        df_aprov = (
            pd.DataFrame(dados_aprov)
            if dados_aprov
            else pd.DataFrame(columns=['Data', 'Técnico', 'Disp', 'TP', 'TG'])
        )

        # Sanitiza cabeçalhos
        df_com.columns = [c.strip() for c in df_com.columns]
        df_aprov.columns = [c.strip() for c in df_aprov.columns]

        # Renomeia colunas de Comissões para padrão interno
        df_com.rename(
            columns={"Data Processamento": "Data", "Sigla Técnico": "Técnico"},
            inplace=True
        )

        # Filtra apenas colunas necessárias
        cols_com = ['Data', 'Técnico', 'Horas Vendidas']
        df_com = df_com[[c for c in cols_com if c in df_com.columns]]

        cols_aprov = ['Data', 'Técnico', 'Disp', 'TP', 'TG']
        df_aprov = df_aprov[[c for c in cols_aprov if c in df_aprov.columns]]

        # Padroniza datas ('08/12/25' → '08/12/2025')
        if 'Data' in df_com.columns:
            df_com['Data'] = df_com['Data'].apply(padronizar_data_quatro_digitos)
        if 'Data' in df_aprov.columns:
            df_aprov['Data'] = df_aprov['Data'].apply(padronizar_data_quatro_digitos)

        # Converte formatação numérica brasileira
        cols_numericas = ['Horas Vendidas', 'Disp', 'TP', 'TG']
        for col in cols_numericas:
            if col in df_com.columns:
                df_com[col] = df_com[col].apply(converter_br_para_float)
            if col in df_aprov.columns:
                df_aprov[col] = df_aprov[col].apply(converter_br_para_float)

        # Chaves textuais para o merge
        df_com['Key_D'] = df_com['Data'].astype(str).str.strip()
        df_com['Key_T'] = df_com['Técnico'].astype(str).str.strip().str.upper()
        df_aprov['Key_D'] = df_aprov['Data'].astype(str).str.strip()
        df_aprov['Key_T'] = df_aprov['Técnico'].astype(str).str.strip().str.upper()

        # Outer join — preserva técnicos que existem em apenas uma das abas
        df_final = pd.merge(
            df_com, df_aprov,
            on=['Key_D', 'Key_T'],
            how='outer',
            suffixes=('_C', '_A')
        )

        # ---------------------------------------------------------------
        # CORREÇÃO 1: NÃO fazer fillna(0.0) aqui.
        # O fillna prematuro transformava NaN em 0.0 nas colunas-chave,
        # impedindo o combine_first de detectar os lados vazios.
        # ---------------------------------------------------------------

        # CORREÇÃO 2: regex=True para substituir '0', '0.0', 'nan' por None
        for col in ['Data_C', 'Data_A', 'Técnico_C', 'Técnico_A']:
            if col in df_final.columns:
                df_final[col] = (
                    df_final[col]
                    .astype(str)
                    .str.strip()
                    .replace(r'^(0|0\.0|nan|None|)$', None, regex=True)
                )

        # Fusão inteligente: preenche o lado vazio com o dado válido do outro
        df_final['Data'] = df_final['Data_C'].combine_first(df_final['Data_A'])
        df_final['Técnico'] = df_final['Técnico_C'].combine_first(df_final['Técnico_A'])

        # CORREÇÃO 3: filtro explícito no lugar do dropna agressivo
        # O dropna eliminava linhas legítimas quando o combine_first falhava por causa do bug 1+2
        df_final = df_final[
            df_final['Data'].notna() & (df_final['Data'].astype(str).str.strip() != '') &
            df_final['Técnico'].notna() & (df_final['Técnico'].astype(str).str.strip() != '')
        ]

        if df_final.empty:
            print("Erro: df_final ficou vazio após o merge. Verifique as chaves Data e Técnico nas abas.")
            return False

        # Reordena para o layout final
        cols_finais = ['Data', 'Técnico', 'Horas Vendidas', 'Disp', 'TP', 'TG']
        df_final = df_final[[c for c in cols_finais if c in df_final.columns]]

        # Agora sim: fillna(0.0) apenas nas colunas numéricas (seguro aqui)
        for col in ['Horas Vendidas', 'Disp', 'TP', 'TG']:
            if col in df_final.columns:
                df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0.0)

        # Ajuste de escala decimal para o BI
        for col in ['Horas Vendidas', 'Disp', 'TP', 'TG']:
            if col in df_final.columns:
                df_final[col] = df_final[col] / 100.0

        # Regras de negócio
        df_final = aplicar_logica_ajustes(df_final)
        df_final = aplicar_traducao_nomes(df_final)

        # Grava no Sheets
        atualizar_planilha_preservando_formato(sh, "Consolidado", df_final)
        return True

    except Exception as e:
        print(f"Erro crítico na rotina de unificação: {e}")
        print(traceback.format_exc())
        return False

# --- ROTINA MESTRA ---
def executar_rotina_global(df_com=None, df_aprov=None):
    status_msg = st.empty()
    bar = st.progress(0)
    try:
        if df_com is not None and not df_com.empty:
            status_msg.info("💾 Salvando Comissões...")
            salvar_com_upsert("Comissoes", df_com, ["Data Processamento", "Sigla Técnico"])
            bar.progress(40)

        if df_aprov is not None and not df_aprov.empty:
            status_msg.info("💾 Salvando Aproveitamento...")
            salvar_com_upsert("Aproveitamento", df_aprov, ["Data", "Técnico"])
            bar.progress(70)

        status_msg.info("🔄 Unificando, Ajustando e Traduzindo Nomes...")
        sucesso = processar_unificacao()
        bar.progress(100)

        if sucesso:
            status_msg.success("✅ Sucesso! Dados Consolidados e Enviados para o BI.")
            st.balloons()
        else:
            status_msg.warning("⚠️ Salvo, mas erro na unificação.")

    except Exception as e:
        status_msg.error(f"Erro: {e}")

# --- HELPER: LISTAR TÉCNICOS ---
def listar_tecnicos_unicos():
    try:
        client = conectar_sheets()
        sh = client.open_by_key(ID_PLANILHA_MESTRA)
        try:
            vals = sh.worksheet("Consolidado").col_values(2)[1:]
        except:
            vals = []
        unicos = sorted(list(set([v for v in vals if v])))
        return unicos
    except:
        return []

# --- INTERFACE ---
st.sidebar.title("Login Seguro")
senha = st.sidebar.text_input("Senha:", type="password")

if senha == verificar_acesso():
    st.sidebar.success("Acesso Liberado")
    st.title("🏭 Central de Processamento WLM")

    ultima_data_base = obter_ultima_atualizacao()
    st.info(f"📅 **Último dado processado na base:** {ultima_data_base}")

    aba1, aba2, aba3 = st.tabs(["💰 Comissões", "⚙️ Aproveitamento", "🔧 Ajustes Manuais"])
    df_comissao_global = None
    df_aprov_global = None

    with aba1:
        st.header("Upload Comissões")
        files_com = st.file_uploader("Arquivos HTML", accept_multiple_files=True, key="up_com")
        if files_com:
            dados_c = parse_comissoes(files_com)
            if dados_c:
                df_comissao_global = pd.DataFrame(dados_c, columns=["Data Processamento", "Nome do Arquivo", "Sigla Técnico", "Horas Vendidas"])
                st.dataframe(df_comissao_global, height=200)

    with aba2:
        st.header("Upload Aproveitamento")
        files_aprov = st.file_uploader("Arquivos HTML/SLK", accept_multiple_files=True, key="up_aprov")
        if files_aprov:
            dados_a = parse_aproveitamento(files_aprov)
            if dados_a:
                df_aprov_global = pd.DataFrame(dados_a, columns=["Data", "Arquivo", "Técnico", "Disp", "TP", "TG"])
                st.dataframe(df_aprov_global, height=200)

    with aba3:
        st.header("Correção e Ajustes")
        st.info("Use esta tela para corrigir dias fechados errados ou transferir horas.")

        with st.form("form_ajustes"):
            col_a, col_b = st.columns(2)
            data_adj = col_a.date_input("Data do Ajuste")
            lista_tec = listar_tecnicos_unicos()
            if not lista_tec:
                lista_tec = ["Digite Manualmente Abaixo"]

            tec_adj = col_b.selectbox("Selecione o Técnico", lista_tec)
            tec_manual = st.text_input("Ou digite a Sigla do Técnico (se não estiver na lista acima)")

            col_c, col_d = st.columns(2)
            metrica_adj = col_c.selectbox("Métrica", [
                "Horas Vendidas (HV)",
                "Tempo Padrão (TP)",
                "Tempo Disponível (Disp)",
                "Tempo Garantia (TG)"
            ])
            valor_adj = col_d.number_input("Valor (+/-)", step=0.5, format="%.2f")

            motivo_adj = st.text_input("Motivo da Correção")

            if st.form_submit_button("💾 Salvar Ajuste e Atualizar BI"):
                tecnico_final = tec_manual.upper().strip() if tec_manual else tec_adj
                if tecnico_final:
                    salvar_ajuste_manual(data_adj, tecnico_final, metrica_adj, valor_adj, motivo_adj)
                    st.success(f"Ajuste salvo para {tecnico_final}!")
                    with st.spinner("Atualizando BI..."):
                        sucesso = processar_unificacao()
                        if sucesso:
                            st.success("BI Atualizado!")
                else:
                    st.error("Selecione um técnico.")

        st.markdown("### Últimos Ajustes")
        try:
            client = conectar_sheets()
            sh = client.open_by_key(ID_PLANILHA_MESTRA)
            try:
                df_ajustes_view = pd.DataFrame(sh.worksheet("Ajustes").get_all_records())
                if not df_ajustes_view.empty:
                    st.dataframe(df_ajustes_view.tail(5))
            except:
                st.write("Nenhum ajuste.")
        except:
            pass

    st.divider()
    col_btn, col_txt = st.columns([1, 4])
    with col_btn:
        if st.button("🚀 GRAVAR TUDO E ATUALIZAR", type="primary"):
            if df_comissao_global is None and df_aprov_global is None:
                st.warning("Sem arquivos.")
            else:
                executar_rotina_global(df_comissao_global, df_aprov_global)

else:
    if senha:
        st.error("Senha incorreta.")

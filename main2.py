import streamlit as st
import pandas as pd
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re

st.set_page_config(page_title="Controle de RM atendidas", layout="wide")
st.title("📦 Controle de RMs - Estocagem e Expedição")
st.markdown("Sistema atualizado: conferência por VOLUME. A capa só será liberada para MAPA quando todas as RMs estiverem 100% atendidas via volumes da expedição.")

# ----------------------
# Utilitários / Normalização
# ----------------------
def clean_colnames(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    newcols = []
    for c in df.columns:
        c2 = str(c).replace("'", "").replace('"', '').strip()
        newcols.append(c2)
    df.columns = newcols
    return df

def normalizar_codigo_rm(valor):
    if pd.isna(valor) or str(valor).strip() == '':
        return ''
    s = str(valor).strip().replace("'", "").replace('"', "")
    s = s.replace(".", "").replace(",", "").replace(" ", "")
    return s

def normalizar_lote(valor):
    if pd.isna(valor):
        return ''
    return str(valor).strip().replace("'", "").replace('"', '')

def singra_indica_em_expedicao(val: str) -> bool:
    if pd.isna(val) or str(val).strip() == '':
        return False
    v = str(val).strip().upper()
    return ('EXPED' in v) or ('EM EXPED' in v) or ('EXPEDIÇÃO' in v) or ('EXPEDICAO' in v)

# ----------------------
# Cache: carregamento arquivos
# ----------------------
@st.cache_data
def carregar_singra(file):
    df = pd.read_csv(file, sep=';', encoding='latin1', dtype=str, low_memory=False)
    df = clean_colnames(df)
    df = df.fillna('')
    if 'ID' in df.columns:
        df['ID'] = df['ID'].apply(normalizar_codigo_rm)
    if 'SITUACAO' in df.columns:
        df['SITUACAO'] = df['SITUACAO'].astype(str).str.strip()
    if 'OMS' in df.columns:
        df['OMS'] = df['OMS'].astype(str).str.strip()
    if 'LISTA_WMS_ID' in df.columns:
        df['LISTA_WMS_ID'] = df['LISTA_WMS_ID'].astype(str).str.strip()
    return df

@st.cache_data
def carregar_pwa(file):
    df = pd.read_excel(file, sheet_name=0, dtype=str)
    df = clean_colnames(df)
    df = df.fillna('')
    for col in ['PEDIDO', 'CAPA', 'MAPA', 'STC', 'CAM', 'LOTE', 'STATUS', 'VOLUME']:
        if col in df.columns:
            df[col] = df[col].astype(str).strip()

    if 'PEDIDO' in df.columns:
        df['PEDIDO_LIMPO'] = df['PEDIDO'].apply(normalizar_codigo_rm)
    else:
        df['PEDIDO'] = ''
        df['PEDIDO_LIMPO'] = ''

    if 'MAPA' in df.columns:
        def mapa_to_intstr(x):
            x = str(x).strip()
            if x == '' or x.upper() == 'NAN':
                return ''
            try:
                if '.' in x:
                    return str(int(float(x)))
                return x
            except:
                return x
        df['MAPA'] = df['MAPA'].apply(mapa_to_intstr)

    if 'STATUS' in df.columns:
        df['STATUS'] = df['STATUS'].astype(str).str.strip().upper()

    return df

@st.cache_data(ttl=3600)
def carregar_lotes_google(credentials_dict: dict, sheet_url: str):
    creds = ServiceAccountCredentials.from_json_keyfile_dict(
        credentials_dict,
        ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    )
    client = gspread.authorize(creds)
    sheet = client.open_by_url(sheet_url)
    worksheet = sheet.get_worksheet(0)
    data = worksheet.get_all_records()
    df = pd.DataFrame(data)
    df = clean_colnames(df)
    df = df.fillna('')
    if 'LOTE' in df.columns:
        df['LOTE'] = df['LOTE'].apply(normalizar_lote)
    return df

# ----------------------
# UI: Uploads
# ----------------------
with st.expander("📄 Upload de arquivos"):
    singra_file = st.file_uploader("Upload planilha do SINGRA (.csv)", type=["csv"])
    pwa_file = st.file_uploader("Upload planilha do PWA (.xlsx)", type=["xlsx"])

if not (singra_file and pwa_file):
    st.info("Faça upload do SINGRA (.csv) e do PWA (.xlsx) para prosseguir.")
    st.stop()

# ----------------------
# Carrega dados
# ----------------------
df_singra = carregar_singra(singra_file)
df_pwa = carregar_pwa(pwa_file)

# Carregar planilha de lotes
SHEET_URL = "https://docs.google.com/spreadsheets/d/1naVnAlUGmeAMb_YftLGYit-1e1BcYFJgiJwSnOcgJf4/edit?gid=0"
service_account_dict = dict(st.secrets["gcp_service_account"])
df_lotes_user = carregar_lotes_google(service_account_dict, SHEET_URL)

# ----------------------
# EXTRAÇÃO DOS VOLUMES
# ----------------------
def extrair_lote_volume(valor: str):
    v = str(valor).strip()
    if "-" in v:
        base, vol = v.split("-", 1)
        return base.strip(), vol.strip()
    return v, None

lotes_volumes_recebidos = {}

for raw in df_lotes_user['LOTE'].astype(str):
    lote_base, vol = extrair_lote_volume(raw)
    if lote_base not in lotes_volumes_recebidos:
        lotes_volumes_recebidos[lote_base] = set()
    if vol is None:
        lotes_volumes_recebidos[lote_base].add("ALL")
    else:
        lotes_volumes_recebidos[lote_base].add(vol)

# ----------------------
# MAPEAMENTO DO PWA
# ----------------------
pwa_lotes_map = {pedido: sorted(set(g['LOTE'].astype(str).tolist())) for pedido, g in df_pwa.groupby('PEDIDO_LIMPO')}
capa_to_rms = {capa: sorted(df_pwa[df_pwa['CAPA'] == capa]['PEDIDO_LIMPO'].unique().tolist()) for capa in sorted(df_pwa['CAPA'].unique().tolist())}

if "VOLUME" not in df_pwa.columns:
    st.error("A planilha do PWA precisa ter a coluna 'VOLUME' para a conferência por volume.")
    st.stop()

lotes_volumes_previstos = {}
for lote, grp in df_pwa.groupby('LOTE'):
    lotes_volumes_previstos[lote] = sorted(set(grp['VOLUME'].astype(str).tolist()))

# ----------------------
# LÓGICA: RM COMPLETA
# ----------------------
def rm_completamente_atendida(rm):
    lotes_da_rm = pwa_lotes_map.get(rm, [])
    for lote in lotes_da_rm:
        previstos = lotes_volumes_previstos.get(lote, [])
        recebidos = lotes_volumes_recebidos.get(lote, set())

        if "ALL" in recebidos:
            continue

        for vol in previstos:
            if vol not in recebidos:
                return False
    return True

# ----------------------
# LÓGICA: CAPA PRONTA
# ----------------------
def capa_pronta_para_mapa(capa):
    rms = capa_to_rms.get(capa, [])
    return all(rm_completamente_atendida(rm) for rm in rms)

# ----------------------
# INTERFACE / RESULTADOS
# ----------------------
st.subheader("📦 CAPAS prontas para MAPA")

capas_ok = [c for c in capa_to_rms if capa_pronta_para_mapa(c)]
df_capas_ok = pd.DataFrame({"CAPA": capas_ok})

st.dataframe(df_capas_ok, use_container_width=True)

st.success(f"Total de CAPAS prontas: {len(capas_ok)}")

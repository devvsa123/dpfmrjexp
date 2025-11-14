import streamlit as st
import pandas as pd
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import re

st.set_page_config(page_title="Controle de RM atendidas", layout="wide")
st.title("📦 Controle de RMs - Estocagem e Expedição")
st.markdown("O sistema organiza as RMs em três blocos: conferência por lote, MAPA sem STC e STC não expedido.")

# ==========================================================
# 🔹 Funções utilitárias
# ==========================================================

def normalizar_codigo(valor):
    """
    Normaliza códigos numéricos:
    - Remove pontos, vírgulas e espaços
    - Converte para inteiro
    - Retorna string com zero à esquerda garantindo 8 dígitos
    """
    try:
        return f"{int(str(valor).replace('.', '').replace(',', '').strip()):08d}"
    except:
        return str(valor).strip()


# ==========================================================
# 🔹 Funções com cache
# ==========================================================

@st.cache_data
def carregar_singra(file):
    df = pd.read_csv(file, sep=';', encoding='latin1')
    df.columns = df.columns.str.replace("'", "").str.strip()
    df = df.fillna('')
    for col in ['ID', 'OMS', 'LISTA_WMS_ID']:
        if col in df.columns:
            df[col] = df[col].astype(str).apply(normalizar_codigo)
    return df

@st.cache_data
def carregar_pwa(file):
    df = pd.read_excel(file, sheet_name=0)
    df = df.fillna('')

    colunas_normalizar = ['PEDIDO', 'LOTE', 'CAPA', 'MAPA', 'STC', 'CAM']
    for col in colunas_normalizar:
        if col in df.columns:
            df[col] = df[col].astype(str).apply(normalizar_codigo)

    df['PEDIDO_LIMPO'] = df['PEDIDO']
    return df

@st.cache_data(ttl=3600)
def carregar_lotes(credentials_dict):
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(
        credentials_dict,
        ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    )
    client = gspread.authorize(credentials)
    spreadsheet = client.open_by_url(
        "https://docs.google.com/spreadsheets/d/1naVnAlUGmeAMb_YftLGYit-1e1BcYFJgiJwSnOcgJf4/edit?gid=0"
    )
    worksheet = spreadsheet.get_worksheet(0)
    data = worksheet.get_all_records()
    df_lotes = pd.DataFrame(data)
    df_lotes['LOTE'] = df_lotes['LOTE'].astype(str).apply(normalizar_codigo)
    return df_lotes


# ==========================================================
# 🔹 Upload dos arquivos
# ==========================================================

with st.expander("📄 Upload de arquivos"):
    singra_file = st.file_uploader("Upload planilha do SINGRA (.csv)", type=["csv"])
    pwa_file = st.file_uploader("Upload planilha do PWA (.xlsx)", type=["xlsx"])

if not (singra_file and pwa_file):
    st.warning("Envie os dois arquivos para continuar.")
    st.stop()

df_singra = carregar_singra(singra_file)
df_pwa = carregar_pwa(pwa_file)

# Carregar lotes do Google Sheets
service_account_dict = dict(st.secrets["gcp_service_account"])
df_lotes_user = carregar_lotes(service_account_dict)


# ==========================================================
# 🔹 BLOCO 1 — Conferência de CAPA atendidas ou não
# ==========================================================

st.markdown("### 📊 Resumo Conferência por CAPA")

capa_completa = []
capa_incompleta = []

capas_unicas = df_singra['LISTA_WMS_ID'].unique()
lotes_disponiveis = df_lotes_user['LOTE'].unique().tolist()

for capa in capas_unicas:

    rms_da_capa = df_singra[df_singra['LISTA_WMS_ID'] == capa]['ID'].unique()

    # Todas RMs expedidas?
    df_pwa_rms = df_pwa[df_pwa['PEDIDO'].isin(rms_da_capa)]

    if df_pwa_rms.empty:
        continue

    todas_expedidas = (df_pwa_rms['STATUS'] == 'EXPEDIDO').all()

    # Verificar lotes da RM
    lotes_rm = df_pwa_rms['LOTE'].unique().tolist()
    faltando_lotes = [l for l in lotes_rm if l not in lotes_disponiveis]

    cam = df_singra[df_singra['LISTA_WMS_ID'] == capa]['OMS'].iloc[0]

    if todas_expedidas and not faltando_lotes:
        capa_completa.append({
            "CAM": cam,
            "CAPA": capa,
            "RMs": ', '.join(rms_da_capa)
        })
    else:
        capa_incompleta.append({
            "CAM": cam,
            "CAPA": capa,
            "RMs": ', '.join(rms_da_capa),
            "LOTES_FALTANTES": ', '.join(faltando_lotes)
        })

col1, col2 = st.columns(2)
col1.success(f"✅ CAPA completamente atendidas: {len(capa_completa)}")
col2.warning(f"⚠️ CAPA parcialmente atendidas: {len(capa_incompleta)}")

df_capa_completa = pd.DataFrame(capa_completa)
df_capa_incompleta = pd.DataFrame(capa_incompleta)

st.subheader("✅ CAPA completamente atendidas")
st.dataframe(df_capa_completa, use_container_width=True)

st.subheader("⚠️ CAPA parcialmente atendidas")
st.dataframe(df_capa_incompleta, use_container_width=True)


# ==========================================================
# 🔹 BLOCO 2 — MAPA sem STC (remover lotes expedidos)
# ==========================================================

st.markdown("### 📋 MAPA sem STC (filtrando apenas RMs NÃO expedidas)")

df_mapa_sem_stc = df_pwa[
    (df_pwa['MAPA'] != '') &
    (df_pwa['STC'] == '') &
    (df_pwa['STATUS'] != 'EXPEDIDO')
]

df_agrupado_mapa = (
    df_mapa_sem_stc.groupby(['CAM', 'MAPA'])
    .agg({'CAPA': lambda x: ', '.join(sorted(set(x)))})
    .reset_index()
)

cams = ["Todos"] + sorted(df_agrupado_mapa['CAM'].unique())
cam_escolhido = st.selectbox("Filtrar CAM (Bloco 2)", cams)

if cam_escolhido != "Todos":
    df_agrupado_mapa = df_agrupado_mapa[df_agrupado_mapa['CAM'] == cam_escolhido]

st.dataframe(df_agrupado_mapa, use_container_width=True)


# ==========================================================
# 🔹 BLOCO 3 — STC não expedidas
# ==========================================================

st.markdown("### 🚚 STC não expedidas")

df_stc_nao_expedida = df_pwa[
    (df_pwa['STC'] != '') &
    (df_pwa['STATUS'] != 'EXPEDIDO') &
    (df_pwa['STATUS'] != 'CANCELADO')
]

df_agrupado_stc = (
    df_stc_nao_expedida.groupby(['CAM', 'STC'])
    .agg({'MAPA': lambda x: ', '.join(sorted(set(x)))})
    .reset_index()
)

cams2 = ["Todos"] + sorted(df_agrupado_stc['CAM'].unique())
cam_escolhido2 = st.selectbox("Filtrar CAM (Bloco 3)", cams2)

if cam_escolhido2 != "Todos":
    df_agrupado_stc = df_agrupado_stc[df_agrupado_stc['CAM'] == cam_escolhido2]

st.dataframe(df_agrupado_stc, use_container_width=True)


# ==========================================================
# 🔹 Exportação Excel
# ==========================================================

def to_excel(dfs, names):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for df, name in zip(dfs, names):
            df.to_excel(writer, sheet_name=name, index=False)
    return output.getvalue()

with st.expander("📥 Exportar resultados"):
    if st.button("Baixar Excel"):
        excel_bytes = to_excel(
            [df_capa_completa, df_capa_incompleta, df_agrupado_mapa, df_agrupado_stc],
            ["CAPA_Atendidas", "CAPA_Pendentes", "MAPA_sem_STC", "STC_nao_expedida"]
        )
        st.download_button(
            label="Download Excel",
            data=excel_bytes,
            file_name="resultado_controle_rm.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

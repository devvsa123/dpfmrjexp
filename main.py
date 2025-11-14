import streamlit as st
import pandas as pd
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re

st.set_page_config(page_title="Controle de RM atendidas", layout="wide")
st.title("📦 Controle de RM - Estocagem e Expedição")
st.markdown("CAPA completa apenas se todas as RMs tiverem **todos os LOTES** lançados na planilha de conferência (Google Sheets).")

# -------------------------
# Utilitárias de normalização
# -------------------------
def normalizar_codigo_8(valor):
    """Normaliza códigos numéricos/tipo RM/CAPA/MAPA/STC/CAM para string (remove pontuações) e tenta pad 8 dígitos."""
    if pd.isna(valor) or str(valor).strip() == '':
        return ''
    s = str(valor).strip().replace("'", "").replace('"', "")
    s = s.replace(".", "").replace(",", "").replace(" ", "")
    try:
        return f"{int(s):08d}"
    except:
        return s

def normalizar_lote(valor):
    """Normaliza lote como string (mantém letras se houver)."""
    if pd.isna(valor):
        return ''
    return str(valor).strip().replace("'", "").replace('"', "")

# -------------------------
# Funções com cache
# -------------------------
@st.cache_data
def carregar_singra(file):
    df = pd.read_csv(file, sep=';', encoding='latin1', dtype=str, low_memory=False)
    df.columns = df.columns.str.replace("'", "").str.strip()
    df = df.fillna('')
    if 'ID' in df.columns:
        df['ID'] = df['ID'].apply(normalizar_codigo_8)
    if 'LISTA_WMS_ID' in df.columns:
        df['LISTA_WMS_ID'] = df['LISTA_WMS_ID'].apply(lambda x: str(x).strip())
    if 'OMS' in df.columns:
        df['OMS'] = df['OMS'].apply(lambda x: str(x).strip())
    return df

@st.cache_data
def carregar_pwa(file):
    df = pd.read_excel(file, sheet_name=0, dtype=str)
    df = df.fillna('')
    # Normalizar campos-chave (quando aplicável)
    for col in ['PEDIDO', 'CAPA', 'MAPA', 'STC', 'CAM']:
        if col in df.columns:
            df[col] = df[col].apply(normalizar_codigo_8)
    if 'LOTE' in df.columns:
        df['LOTE'] = df['LOTE'].apply(normalizar_lote)
    if 'STATUS' in df.columns:
        df['STATUS'] = df['STATUS'].astype(str).str.strip().str.upper()
    # Campo auxiliar já limpo (PEDIDO)
    if 'PEDIDO' not in df.columns:
        df['PEDIDO'] = ''
    df['PEDIDO_LIMPO'] = df['PEDIDO']
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
    if 'LOTE' in df.columns:
        df['LOTE'] = df['LOTE'].apply(normalizar_lote)
    else:
        df['LOTE'] = []
    return df

@st.cache_data
def agrupar_mapa_sem_stc_cached(df):
    return (
        df[(df['MAPA'] != '') & (df['STC'] == '') & (df['STATUS'] != 'EXPEDIDO')]
        .groupby(['CAM', 'MAPA'])
        .agg({'CAPA': lambda x: ', '.join(sorted(set(x)))})
        .reset_index()
    )

@st.cache_data
def agrupar_stc_nao_expedida_cached(df):
    return (
        df[(df['STC'] != '') & (df['STATUS'] != 'EXPEDIDO') & (df['STATUS'] != 'CANCELADO')]
        .groupby(['CAM', 'STC'])
        .agg({'MAPA': lambda x: ', '.join(sorted(set(x)))})
        .reset_index()
    )

# -------------------------
# Upload dos arquivos (UI)
# -------------------------
with st.expander("📄 Upload de arquivos (SINGRA e PWA)"):
    singra_file = st.file_uploader("Upload planilha do SINGRA (.csv)", type=["csv"])
    pwa_file = st.file_uploader("Upload planilha do PWA (.xlsx)", type=["xlsx"])

if not (singra_file and pwa_file):
    st.info("Faça upload do SINGRA (.csv) e do PWA (.xlsx) para prosseguir.")
    st.stop()

# Carregar dados com cache
df_singra = carregar_singra(singra_file)
df_pwa = carregar_pwa(pwa_file)

# Carregar lotes do Google Sheets (secrets)
SHEET_URL = "https://docs.google.com/spreadsheets/d/1naVnAlUGmeAMb_YftLGYit-1e1BcYFJgiJwSnOcgJf4/edit?gid=0"
service_account_dict = dict(st.secrets["gcp_service_account"])
df_lotes_user = carregar_lotes_google(service_account_dict, SHEET_URL)

# Pré-processamentos rápidos
df_pwa['PEDIDO_LIMPO'] = df_pwa['PEDIDO'].astype(str).str.replace(".", "", regex=False)

# -------------------------
# Debug rápido (contagens)
# -------------------------
col1, col2, col3 = st.columns(3)
col1.metric("RMs (Singra)", df_singra['ID'].nunique() if 'ID' in df_singra.columns else 0)
col2.metric("Linhas PWA", len(df_pwa))
col3.metric("Lotes (planilha conf.)", df_lotes_user['LOTE'].nunique() if 'LOTE' in df_lotes_user.columns else 0)

# -------------------------
# BLOCO 1 - lógica CORRETA
# -------------------------
st.markdown("## 🔵 BLOCO 1 — Conferência por CAPA (RM completa somente se TODOS os LOTES estiverem lançados)")

# lista lotes disponíveis (conferência)
lotes_disponiveis = set(df_lotes_user['LOTE'].astype(str).tolist())

# Construir mapeamento PWA por PEDIDO para acelerar
pwa_map = {}
for pedido, g in df_pwa.groupby('PEDIDO'):
    lotes = sorted(set(g['LOTE'].astype(str).tolist()))
    statuses = sorted(set(g['STATUS'].astype(str).tolist())) if 'STATUS' in g.columns else []
    pwa_map[pedido] = {'lotes': lotes, 'statuses': statuses}

# Montar df_rm_status
rm_rows = []
singra_rms = sorted(df_singra['ID'].unique().tolist()) if 'ID' in df_singra.columns else []

for rm in singra_rms:
    p = pwa_map.get(rm, {'lotes': [], 'statuses': []})
    lotes_rm = p['lotes']
    statuses = p['statuses']
    if not lotes_rm:
        lotes_faltando = ["<SEM_LOTE_NO_PWA>"]
        completa = False
    else:
        lotes_faltando = [l for l in lotes_rm if l not in lotes_disponiveis]
        completa = (len(lotes_faltando) == 0)
    expedido = False
    if statuses:
        expedido = all([s.upper() == 'EXPEDIDO' for s in statuses])
    rm_rows.append({
        "RM": rm,
        "LOTES": ', '.join(lotes_rm),
        "LOTES_FALTANDO": ', '.join(lotes_faltando) if lotes_faltando else '',
        "COMPLETA": completa,
        "EXPEDIDO": expedido
    })

df_rm_status = pd.DataFrame(rm_rows)

# Avaliar por CAPA
capas = sorted(df_singra['LISTA_WMS_ID'].unique().tolist()) if 'LISTA_WMS_ID' in df_singra.columns else []
capa_complete_list = []
capa_incomplete_list = []

for capa in capas:
    rms_da_capa = df_singra[df_singra['LISTA_WMS_ID'] == capa]['ID'].unique().tolist()
    sub = df_rm_status[df_rm_status['RM'].isin(rms_da_capa)]
    # se sub está vazio ou alguma RM não completa -> incompleto
    if (not sub.empty) and sub['COMPLETA'].all():
        capa_complete_list.append({
            "CAM": df_singra.loc[df_singra['LISTA_WMS_ID'] == capa, 'OMS'].iloc[0] if 'OMS' in df_singra.columns else '',
            "CAPA": capa,
            "RMs": ', '.join(rms_da_capa)
        })
    else:
        # construir detalhes das RMs incompletas
        details = []
        for _, rrow in sub[~sub['COMPLETA']].iterrows():
            details.append(f"{rrow['RM']}: faltam [{rrow['LOTES_FALTANDO']}]")
        capa_incomplete_list.append({
            "CAM": df_singra.loc[df_singra['LISTA_WMS_ID'] == capa, 'OMS'].iloc[0] if 'OMS' in df_singra.columns else '',
            "CAPA": capa,
            "RMs": ', '.join(rms_da_capa),
            "DETALHES": '; '.join(details) if details else ''
        })

df_capa_completa = pd.DataFrame(capa_complete_list)
df_capa_incompleta = pd.DataFrame(capa_incomplete_list)

# Mostra resumo e tabelas
c1, c2 = st.columns([1,1])
c1.success(f"CAPAs totalmente atendidas (todas RMs com todos os lotes lançados): {len(df_capa_completa)}")
c2.warning(f"CAPAs com pendências: {len(df_capa_incompleta)}")

st.subheader("✅ CAPAs completamente atendidas")
if not df_capa_completa.empty:
    st.dataframe(df_capa_completa.style.set_properties(**{'text-align': 'left'}), use_container_width=True)
else:
    st.info("Nenhuma CAPA totalmente atendida segundo a regra de lotes lançados.")

st.subheader("⚠️ CAPAs parcialmente atendidas (detalhes por RM)")
if not df_capa_incompleta.empty:
    st.dataframe(df_capa_incompleta.style.set_properties(**{'text-align': 'left'}), use_container_width=True)
else:
    st.info("Nenhuma CAPA parcialmente atendida encontrada.")

# Expansion com detalhes por RM (com destaque)
with st.expander("🔎 Detalhes por RM (depuração)"):
    if not df_rm_status.empty:
        # Estilo: linha verde se COMPLETA True, amarela se incompleta e não expedida, cinza se expedida
        def highlight_row(row):
            if row['EXPEDIDO']:
                return ['background-color: #d3d3d3' for _ in row]
            if row['COMPLETA']:
                return ['background-color: #d4edda' for _ in row]  # verde claro
            return ['background-color: #fff3cd' for _ in row]  # amarelo claro
        st.dataframe(df_rm_status.style.apply(highlight_row, axis=1).set_properties(**{'text-align':'left'}), use_container_width=True)
    else:
        st.info("Nenhum detalhe de RM disponível.")

# -------------------------
# BLOCO 2 - MAPA sem STC (não exibir EXPEDIDOS)
# -------------------------
st.markdown("## 🔷 BLOCO 2 — MAPA sem STC (exclui RM com STATUS = EXPEDIDO)")

if all(col in df_pwa.columns for col in ['MAPA','STC','STATUS','CAM','CAPA']):
    df_mapa_sem_stc = df_pwa[
        (df_pwa['MAPA'] != '') &
        (df_pwa['STC'] == '') &
        (df_pwa['STATUS'] != 'EXPEDIDO')
    ]
    if df_mapa_sem_stc.empty:
        st.info("Nenhuma linha com MAPA sem STC (após filtrar EXPEDIDOS).")
        df_agrupado_mapa = pd.DataFrame()
    else:
        df_agrupado_mapa = df_mapa_sem_stc.groupby(['CAM','MAPA']).agg({'CAPA': lambda x: ', '.join(sorted(set(x)))}).reset_index()
        cams = ["Todos"] + sorted(df_agrupado_mapa['CAM'].unique().tolist())
        cam_sel = st.selectbox("Filtrar CAM (Bloco 2)", cams)
        display_mapa = df_agrupado_mapa if cam_sel == "Todos" else df_agrupado_mapa[df_agrupado_mapa['CAM'] == cam_sel]
        st.dataframe(display_mapa.style.set_properties(**{'text-align':'left'}), use_container_width=True)
else:
    st.info("Colunas necessárias para Bloco 2 ausentes no PWA.")

# -------------------------
# BLOCO 3 - STC não expedidas (agrupar por CAM e STC)
# -------------------------
st.markdown("## 🔶 BLOCO 3 — STC não expedidas (agrupar por CAM e STC)")

if all(col in df_pwa.columns for col in ['STC','STATUS','CAM']):
    df_stc_nao_expedida = df_pwa[
        (df_pwa['STC'] != '') &
        (df_pwa['STATUS'] != 'EXPEDIDO') &
        (df_pwa['STATUS'] != 'CANCELADO')
    ]
    if df_stc_nao_expedida.empty:
        st.info("Nenhuma STC pendente.")
        df_agrupado_stc = pd.DataFrame()
    else:
        df_agrupado_stc = df_stc_nao_expedida.groupby(['CAM','STC']).agg({'MAPA': lambda x: ', '.join(sorted(set(x)))}).reset_index()
        cams3 = ["Todos"] + sorted(df_agrupado_stc['CAM'].unique().tolist())
        cam_sel3 = st.selectbox("Filtrar CAM (Bloco 3)", cams3)
        display_stc = df_agrupado_stc if cam_sel3 == "Todos" else df_agrupado_stc[df_agrupado_stc['CAM'] == cam_sel3]
        st.dataframe(display_stc.style.set_properties(**{'text-align':'left'}), use_container_width=True)
else:
    st.info("Colunas necessárias para Bloco 3 ausentes no PWA.")

# -------------------------
# Exportação (inclui df_rm_status)
# -------------------------
def to_excel(dfs, names):
    out = BytesIO()
    with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
        for df, name in zip(dfs, names):
            try:
                df.to_excel(writer, sheet_name=name, index=False)
            except Exception:
                pd.DataFrame(df).to_excel(writer, sheet_name=name, index=False)
    return out.getvalue()

with st.expander("📥 Exportar resultados (Excel)"):
    if st.button("Gerar Excel de saída"):
        excel_bytes = to_excel(
            [df_capa_completa, df_capa_incompleta, df_agrupado_mapa if 'df_agrupado_mapa' in locals() else pd.DataFrame(),
             df_agrupado_stc if 'df_agrupado_stc' in locals() else (df_agrupado_stc if 'df_agrupado_stc' in locals() else pd.DataFrame()),
             df_rm_status],
            ["CAPA_Atendidas", "CAPA_Pendentes", "MAPA_sem_STC", "STC_nao_expedida", "RM_Status"]
        )
        st.download_button(
            label="📥 Baixar Excel completo",
            data=excel_bytes,
            file_name="resultado_controle_rm_completo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.markdown("✅ Pronto — regras aplicadas: **RM completa** = todos os LOTES da RM presentes na planilha de conferência; **CAPA completa** = todas as RMs da CAPA completas.")

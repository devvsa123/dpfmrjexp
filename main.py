import streamlit as st
import pandas as pd
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re

st.set_page_config(page_title="Controle de RM atendidas", layout="wide")
st.title("📦 Controle de RMs - Estocagem e Expedição")
st.markdown("Sistema com PWA como fonte da verdade. Verifica CAPA/RM/LOTE comparando PWA ⇄ planilha de conferência (Google) e SINGRA.")

# ----------------------
# Utilitários / Normalização
# ----------------------
def clean_colnames(df: pd.DataFrame) -> pd.DataFrame:
    """Remove aspas, espaços invisíveis e normaliza nomes de colunas."""
    df = df.copy()
    newcols = []
    for c in df.columns:
        c2 = str(c).replace("'", "").replace('"', '').strip()
        newcols.append(c2)
    df.columns = newcols
    return df

def normalizar_codigo_rm(valor):
    """Remove pontos/espacos e retorna string (preservando zeros se houver)."""
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
    """Decide se a coluna SITUACAO no singra indica que a RM está em expedição."""
    if pd.isna(val) or str(val).strip() == '':
        return False
    v = str(val).strip().upper()
    # aceitar variações: 'EM EXPEDIÇÃO', 'EXPEDICAO', 'EXPEDIDO' etc.
    return ('EXPED' in v) or ('EM EXPED' in v) or ('EXPEDIÇÃO' in v) or ('EXPEDICAO' in v)

# ----------------------
# Cache: carregamento arquivos
# ----------------------
@st.cache_data
def carregar_singra(file):
    df = pd.read_csv(file, sep=';', encoding='latin1', dtype=str, low_memory=False)
    df = clean_colnames(df)
    df = df.fillna('')
    # Normalizações comuns
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
    # Normalizar campos
    for col in ['PEDIDO', 'CAPA', 'MAPA', 'STC', 'CAM', 'LOTE', 'STATUS']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    # Normalizar PEDIDO removendo pontos para comparação (mas preservamos original)
    if 'PEDIDO' in df.columns:
        df['PEDIDO_LIMPO'] = df['PEDIDO'].apply(normalizar_codigo_rm)
    else:
        df['PEDIDO'] = ''
        df['PEDIDO_LIMPO'] = ''
    # Normalize MAPA to integer-like string when present
    if 'MAPA' in df.columns:
        def mapa_to_intstr(x):
            x = str(x).strip()
            if x == '' or x.upper() == 'NAN':
                return ''
            # try float -> int to remove .0
            try:
                if '.' in x:
                    return str(int(float(x)))
                return x
            except:
                return x
        df['MAPA'] = df['MAPA'].apply(mapa_to_intstr)
    # Upper STATUS for robust comparisons
    if 'STATUS' in df.columns:
        df['STATUS'] = df['STATUS'].astype(str).str.strip().str.upper()
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
# Uploads UI
# ----------------------
with st.expander("📄 Upload de arquivos"):
    singra_file = st.file_uploader("Upload planilha do SINGRA (.csv)", type=["csv"])
    pwa_file = st.file_uploader("Upload planilha do PWA (.xlsx)", type=["xlsx"])

if not (singra_file and pwa_file):
    st.info("Faça upload do SINGRA (.csv) e do PWA (.xlsx) para prosseguir.")
    st.stop()

# Carregar dados (cached)
df_singra = carregar_singra(singra_file)
df_pwa = carregar_pwa(pwa_file)

# Carregar lotes do Google Sheets (planilha de conferência)
SHEET_URL = "https://docs.google.com/spreadsheets/d/1naVnAlUGmeAMb_YftLGYit-1e1BcYFJgiJwSnOcgJf4/edit?gid=0"
service_account_dict = dict(st.secrets["gcp_service_account"])
df_lotes_user = carregar_lotes_google(service_account_dict, SHEET_URL)

# Pré-process rápido: set de lotes disponíveis na conferência
lotes_disponiveis = set(df_lotes_user['LOTE'].astype(str).tolist()) if 'LOTE' in df_lotes_user.columns else set()

# Adicional: mapear SINGRA por RM para checagens de migração/status
singra_map = {}
if 'ID' in df_singra.columns:
    for rm, grp in df_singra.groupby('ID'):
        # pegar primeira SITUACAO se existir
        situ = grp['SITUACAO'].iloc[0] if 'SITUACAO' in df_singra.columns else ''
        oms = grp['OMS'].iloc[0] if 'OMS' in df_singra.columns else ''
        singra_map[rm] = {'SITUACAO': situ, 'OMS': oms}

# Preprocess PWA grouping para rapidez
pwa_by_pedido = {pedido: g for pedido, g in df_pwa.groupby('PEDIDO_LIMPO')}

# Debug counters / métricas
col1, col2, col3 = st.columns(3)
col1.metric("RMs (PWA únicas)", df_pwa['PEDIDO_LIMPO'].nunique())
col2.metric("Linhas PWA", len(df_pwa))
col3.metric("Lotes (planilha conf.)", len(lotes_disponiveis))

# ----------------------
# Nova aba: Consulta rápida via texto (mantive sua UI)
# ----------------------
st.markdown("### 🔍 Consulta rápida de RMs (via texto)")
with st.expander("Consultar RMs colando mensagem"):
    texto_rms = st.text_area("Cole aqui a mensagem com as RMs", height=180)
    if st.button("🔎 Consultar RMs no sistema"):
        if texto_rms.strip():
            rms_extraidas = re.findall(r"\b\d{2}\.\d{3}\.\d{3}\b", texto_rms)
            if rms_extraidas:
                rms_sem_ponto = [rm.replace(".", "") for rm in rms_extraidas]
                # filtrar em bloco
                df_filtro = df_pwa[df_pwa['PEDIDO_LIMPO'].isin(rms_sem_ponto)]
                resultados = []
                for rm_texto, rm_limpo in zip(rms_extraidas, rms_sem_ponto):
                    dados_rm = df_filtro[df_filtro['PEDIDO_LIMPO'] == rm_limpo]
                    if not dados_rm.empty:
                        mapa = ', '.join(sorted(set([m for m in dados_rm['MAPA'] if m and m != '']))) or "Não consta"
                        stc  = ', '.join(sorted(set([s for s in dados_rm['STC'] if s and s != '']))) or "Não consta"
                        status = ', '.join(sorted(set(dados_rm['STATUS']))) if 'STATUS' in dados_rm.columns else ''
                    else:
                        mapa = "Não consta"
                        stc = "Não consta"
                        status = ""
                    resultados.append({
                        "RM (texto)": rm_texto,
                        "RM (planilha)": rm_limpo,
                        "STATUS (PWA)": status,
                        "MAPA": mapa,
                        "STC": stc
                    })
                st.dataframe(pd.DataFrame(resultados).style.set_properties(**{'text-align':'left'}), use_container_width=True)
            else:
                st.warning("⚠️ Nenhuma RM válida encontrada no texto.")
        else:
            st.info("Cole o texto e clique em Consultar.")

# ----------------------
# BLOCO 1: CAPA completamente atendidas (PWA = verdade) — LÓGICA NOVA
# ----------------------
st.markdown("## 🔵 BLOCO 1 — CAPA: verificação completa (PWA como fonte de verdade)")

# Verificar colunas essenciais
required_pwa_cols = ['PEDIDO_LIMPO', 'LOTE', 'CAPA', 'CAM', 'STATUS']
if not all(c in df_pwa.columns for c in required_pwa_cols):
    st.error("Colunas essenciais ausentes no PWA. Preciso de PEDIDO/LOTE/CAPA/CAM/STATUS.")
else:
    capas = sorted(df_pwa['CAPA'].unique().tolist())

    capa_completa_rows = []
    capa_incompleta_rows = []

    # précriar mapa de lotes por pedido (para performance)
    pwa_lotes_map = {}
    for pedido_limpo, g in df_pwa.groupby('PEDIDO_LIMPO'):
        pwa_lotes_map[pedido_limpo] = sorted(set(g['LOTE'].astype(str).tolist()))
    # também mapear CAPA -> pedidos
    capa_to_rms = {}
    for capa in capas:
        capa_to_rms[capa] = sorted(df_pwa[df_pwa['CAPA'] == capa]['PEDIDO_LIMPO'].unique().tolist())

    for capa in capas:
        rms_da_capa = capa_to_rms.get(capa, [])
        cam = df_pwa[df_pwa['CAPA'] == capa]['CAM'].iloc[0] if len(df_pwa[df_pwa['CAPA'] == capa])>0 else ''
        pendencias = []  # acumula strings descrevendo problemas por RM

        for rm in rms_da_capa:
            # pegar lotes esperados (PWA)
            lotes_pwa = pwa_lotes_map.get(rm, [])
            # verificar presença de RM no singra (por ID)
            singra_info = singra_map.get(rm)
            if not singra_info:
                # RM não existe no SINGRA: pendência crítica
                pendencias.append(f"{rm} (Status SINGRA não migrou)")
                continue

            # Verificar se SINGRA marca como em expedição
            situ = singra_info.get('SITUACAO', '')
            if not singra_indica_em_expedicao(situ):
                pendencias.append(f"{rm} (SITUACAO SINGRA: '{situ}' não indica expedição)")
                continue

            # Verificar lotes: todos os lotes da RM (PWA) devem estar na planilha de conferência
            faltando_lotes = [l for l in lotes_pwa if l not in lotes_disponiveis]
            if faltando_lotes:
                pendencias.append(f"{rm} – faltando lotes: {', '.join(faltando_lotes)}")

        # decidir situação da CAPA
        if len(pendencias) == 0:
            capa_completa_rows.append({"CAM": cam, "CAPA": capa, "RMs": ', '.join(rms_da_capa)})
        else:
            capa_incompleta_rows.append({"CAM": cam, "CAPA": capa, "Pendências": '; '.join(pendencias)})

    df_capa_completa = pd.DataFrame(capa_completa_rows)
    df_capa_incompleta = pd.DataFrame(capa_incompleta_rows)

    # Resumo
    col_a, col_b = st.columns(2)
    col_a.success(f"CAPAs totalmente atendidas: {len(df_capa_completa)}")
    col_b.warning(f"CAPAs com pendências: {len(df_capa_incompleta)}")

    st.subheader("✅ CAPAs completamente atendidas")
    if not df_capa_completa.empty:
        st.dataframe(df_capa_completa.style.set_properties(**{'text-align':'left'}), use_container_width=True)
    else:
        st.info("Nenhuma CAPA completamente atendida segundo a nova regra.")

    st.subheader("⚠️ CAPAs parcialmente atendidas (detalhes)")
    if not df_capa_incompleta.empty:
        st.dataframe(df_capa_incompleta.style.set_properties(**{'text-align':'left'}), use_container_width=True)
    else:
        st.info("Nenhuma CAPA parcialmente atendida encontrada.")

# ----------------------
# BLOCO 2: MAPA sem STC (agrupado por CAM e MAPA), excluir EXPEDIDO
# ----------------------
st.markdown("## 🔷 BLOCO 2 — MAPA sem STC (agrupar por CAM e MAPA)")

if all(c in df_pwa.columns for c in ['MAPA','STC','STATUS','CAM','CAPA']):
    df_mapa_sem_stc = df_pwa[
        (df_pwa['MAPA'] != '') &
        (df_pwa['STC'] == '') &
        (df_pwa['STATUS'] != 'EXPEDIDO')
    ]
    if df_mapa_sem_stc.empty:
        st.info("Nenhuma MAPA sem STC (após filtrar EXPEDIDO).")
    else:
        agrupado_mapa = (
            df_mapa_sem_stc.groupby(['CAM','MAPA'])
            .agg({'CAPA': lambda x: ', '.join(sorted(set(x)))})
            .reset_index()
        )
        cams = ["Todos"] + sorted(agrupado_mapa['CAM'].unique().tolist())
        cam_sel = st.selectbox("Filtrar por CAM (Bloco 2)", cams)
        display = agrupado_mapa if cam_sel == "Todos" else agrupado_mapa[agrupado_mapa['CAM'] == cam_sel]
        st.dataframe(display.style.set_properties(**{'text-align':'left'}), use_container_width=True)
else:
    st.info("Colunas necessárias para Bloco 2 ausentes no PWA.")

# ----------------------
# BLOCO 3: STC não expedidas (agrupar por CAM e STC)
# ----------------------
st.markdown("## 🔶 BLOCO 3 — STC não expedidas (agrupar por CAM e STC)")

if all(c in df_pwa.columns for c in ['STC','STATUS','CAM','MAPA']):
    df_stc_nao_expedida = df_pwa[
        (df_pwa['STC'] != '') &
        (df_pwa['STATUS'] != 'EXPEDIDO') &
        (df_pwa['STATUS'] != 'CANCELADO')
    ]
    if df_stc_nao_expedida.empty:
        st.info("Nenhuma STC pendente.")
    else:
        agrupado_stc = (
            df_stc_nao_expedida.groupby(['CAM','STC'])
            .agg({'MAPA': lambda x: ', '.join(sorted(set([m for m in x if m and m != ''])))})
            .reset_index()
        )
        cams3 = ["Todos"] + sorted(agrupado_stc['CAM'].unique().tolist())
        cam_sel3 = st.selectbox("Filtrar por CAM (Bloco 3)", cams3)
        display3 = agrupado_stc if cam_sel3 == "Todos" else agrupado_stc[agrupado_stc['CAM'] == cam_sel3]
        st.dataframe(display3.style.set_properties(**{'text-align':'left'}), use_container_width=True)
else:
    st.info("Colunas necessárias para Bloco 3 ausentes no PWA.")

# ----------------------
# Exportação Excel (inclui dados de depuração)
# ----------------------
def to_excel(dfs, names):
    out = BytesIO()
    with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
        for df, name in zip(dfs, names):
            try:
                df.to_excel(writer, sheet_name=name, index=False)
            except Exception:
                pd.DataFrame(df).to_excel(writer, sheet_name=name, index=False)
    return out.getvalue()

with st.expander("📥 Exportar resultados"):
    if st.button("Gerar Excel de saída"):
        # montar objetos para export
        export_dfs = [
            df_capa_completa if 'df_capa_completa' in locals() else pd.DataFrame(),
            df_capa_incompleta if 'df_capa_incompleta' in locals() else pd.DataFrame(),
            agrupado_mapa if 'agrupado_mapa' in locals() else pd.DataFrame(),
            agrupado_stc if 'agrupado_stc' in locals() else pd.DataFrame(),
            df_singra if 'df_singra' in locals() else pd.DataFrame(),
            df_pwa if 'df_pwa' in locals() else pd.DataFrame(),
            df_lotes_user if 'df_lotes_user' in locals() else pd.DataFrame()
        ]
        names = ["CAPA_Atendidas", "CAPA_Pendentes", "MAPA_sem_STC", "STC_nao_expedida", "SINGRA_RAW", "PWA_RAW", "LOTES_CONFERENCIA"]
        excel_bytes = to_excel(export_dfs, names)
        st.download_button(
            label="📥 Baixar Excel completo",
            data=excel_bytes,
            file_name="resultado_controle_rm_completo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.markdown("✅ Regra aplicada: **PWA = fonte da verdade**. CAPA completa somente se todas as RMs da CAPA estão com todos os lotes lançados na planilha de conferência e com SINGRA indicando expedição.")

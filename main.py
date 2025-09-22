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

# ===============================
# 🔹 Funções com cache
# ===============================
@st.cache_data
def carregar_singra(file):
    df = pd.read_csv(file, sep=';', encoding='latin1')
    df.columns = df.columns.str.replace("'", "").str.strip()
    df = df.fillna('')
    for col in ['ID', 'OMS', 'LISTA_WMS_ID']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace("'", "").str.replace('"', '').str.strip()
    return df

@st.cache_data
def carregar_pwa(file):
    df = pd.read_excel(file, sheet_name=0)
    df = df.fillna('')
    for col in ['PEDIDO', 'LOTE', 'CAPA', 'MAPA', 'STC', 'STATUS', 'CAM']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace("'", "").str.replace('"', '').str.strip()
    if 'MAPA' in df.columns:
        df['MAPA'] = df['MAPA'].apply(lambda x: str(int(float(x))) if x not in ['', None] else '')
    df['PEDIDO_LIMPO'] = df['PEDIDO'].astype(str).str.replace(".", "", regex=False)
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
    df_lotes['LOTE'] = df_lotes['LOTE'].astype(str).str.strip()
    return df_lotes

@st.cache_data
def agrupar_mapa_sem_stc(df):
    return (
        df[(df['MAPA'] != '') & (df['STC'] == '') & (df['STATUS'] != 'EXPEDIDO')]
        .groupby(['CAM', 'MAPA'])
        .agg({'CAPA': lambda x: ', '.join(sorted(set(x)))})
        .reset_index()
    )

@st.cache_data
def agrupar_stc_nao_expedida(df):
    return (
        df[(df['STC'] != '') & (df['STATUS'] != 'EXPEDIDO') & (df['STATUS'] != 'CANCELADO')]
        .groupby(['CAM', 'STC'])
        .agg({'MAPA': lambda x: ', '.join(sorted(set(x)))})
        .reset_index()
    )

# ===============================
# 🔹 Upload dos arquivos
# ===============================
with st.expander("📄 Upload de arquivos"):
    singra_file = st.file_uploader("Upload planilha do SINGRA (.csv)", type=["csv"])
    pwa_file = st.file_uploader("Upload planilha do PWA (.xlsx)", type=["xlsx"])

if singra_file and pwa_file:
    df_singra = carregar_singra(singra_file)
    df_pwa = carregar_pwa(pwa_file)

    # Carregar lotes do Google Sheets
    service_account_dict = dict(st.secrets["gcp_service_account"])
    df_lotes_user = carregar_lotes(service_account_dict)

    # ===============================
    # 🔹 Consulta rápida de RMs
    # ===============================
    st.markdown("### 🔍 Consulta rápida de RMs (via texto)")
    with st.expander("Consultar RMs colando mensagem"):
        texto_rms = st.text_area("Cole aqui a mensagem com as RMs", height=200)
        if st.button("🔎 Consultar RMs no sistema"):
            if texto_rms.strip():
                rms_extraidas = re.findall(r"\b\d{2}\.\d{3}\.\d{3}\b", texto_rms)
                if rms_extraidas:
                    rms_sem_ponto = [rm.replace(".", "") for rm in rms_extraidas]
                    df_filtro = df_pwa[df_pwa['PEDIDO_LIMPO'].isin(rms_sem_ponto)]
                    resultados = []
                    for rm, rm_limpo in zip(rms_extraidas, rms_sem_ponto):
                        dados_rm = df_filtro[df_filtro['PEDIDO_LIMPO'] == rm_limpo]
                        mapa = ', '.join(dados_rm['MAPA'].unique()) if any(dados_rm['MAPA'] != '') else "Não consta"
                        stc = ', '.join(dados_rm['STC'].unique()) if any(dados_rm['STC'] != '') else "Não consta"
                        resultados.append({
                            "RM (texto)": rm,
                            "RM (planilha)": rm_limpo,
                            "MAPA": mapa,
                            "STC": stc
                        })
                    df_resultados = pd.DataFrame(resultados)
                    st.dataframe(df_resultados.style.set_properties(**{'text-align': 'left'}))
                else:
                    st.warning("⚠️ Nenhuma RM válida encontrada no texto.")
            else:
                st.info("Cole o texto acima e clique em **Consultar**.")

    # ===============================
    # 🔹 BLOCO 1: CAPA completamente atendidas
    # ===============================
    capa_completa = []
    capa_incompleta = []
    capas_unicas = df_singra['LISTA_WMS_ID'].unique() if 'LISTA_WMS_ID' in df_singra.columns else []
    lotes_disponiveis = df_lotes_user['LOTE'].unique().tolist()

    for capa in capas_unicas:
        rms_da_capa = df_singra[df_singra['LISTA_WMS_ID'] == capa]['ID'].unique()
        rms_sem_mapa = [rm for rm in rms_da_capa if df_pwa[(df_pwa['PEDIDO'] == rm)]['MAPA'].eq('').any()]
        if not rms_sem_mapa:
            continue
        todos_lotes_capa = []
        for rm in rms_sem_mapa:
            lotes_rm = df_pwa[df_pwa['PEDIDO'] == rm]['LOTE'].unique().tolist()
            todos_lotes_capa.extend(lotes_rm)
        faltando_lotes = [l for l in todos_lotes_capa if l not in lotes_disponiveis]
        cam = df_singra.loc[df_singra['LISTA_WMS_ID'] == capa, 'OMS'].values[0] if 'OMS' in df_singra.columns else ''
        if not faltando_lotes:
            capa_completa.append({"CAM": cam, "CAPA": capa, "RMs": ', '.join(rms_sem_mapa)})
        else:
            capa_incompleta.append({"CAM": cam, "CAPA": capa, "RMs": ', '.join(rms_sem_mapa), "LOTES_FALTANDO": ', '.join(faltando_lotes)})

    st.markdown("### 📊 Resumo Conferência por CAPA")
    col1, col2 = st.columns(2)
    col1.success(f"✅ Total de CAPA completamente atendidas (sem MAPA): {len(capa_completa)}")
    col2.warning(f"⚠️ Total de CAPA parcialmente atendidas (pendências): {len(capa_incompleta)}")

    st.subheader("✅ CAPA completamente atendidas (agrupadas por CAM e CAPA)")
    df_capa_completa = pd.DataFrame(capa_completa)
    if not df_capa_completa.empty:
        st.dataframe(df_capa_completa.style.set_properties(**{'text-align': 'left'}))
    else:
        st.info("Nenhuma CAPA completamente atendida encontrada.")

    st.subheader("⚠️ CAPA parcialmente atendidas (agrupadas por CAM e CAPA)")
    df_capa_incompleta = pd.DataFrame(capa_incompleta)
    if not df_capa_incompleta.empty:
        st.dataframe(df_capa_incompleta.style.set_properties(**{'text-align': 'left'}))
    else:
        st.info("Nenhuma CAPA parcialmente atendida encontrada.")

    # ===============================
    # 🔹 BLOCO 2: RMs com MAPA sem STC
    # ===============================
    st.markdown("### 📋 Mapas de Carregamento sem STC lançada no WMS")
    if all(col in df_pwa.columns for col in ['MAPA', 'STC', 'STATUS', 'CAM', 'CAPA']):
        agrupado_mapa = agrupar_mapa_sem_stc(df_pwa)
        cams_disponiveis = ["Todos"] + sorted(agrupado_mapa['CAM'].unique().tolist())
        cam_escolhido = st.selectbox("Filtrar por CAM", cams_disponiveis)
        agrupado_filtrado = agrupado_mapa if cam_escolhido == "Todos" else agrupado_mapa[agrupado_mapa['CAM'] == cam_escolhido]
        if not agrupado_filtrado.empty:
            st.dataframe(agrupado_filtrado.style.set_properties(**{'text-align': 'left'}), use_container_width=True)
        else:
            st.info("Nenhuma RM encontrada com MAPA sem STC.")

    # ===============================
    # 🔹 BLOCO 3: RMs com STC mas não expedidas
    # ===============================
    st.markdown("### 🚚 STC não expedidas")
    if all(col in df_pwa.columns for col in ['STC', 'STATUS', 'CAM', 'CAPA']):
        agrupado_stc = agrupar_stc_nao_expedida(df_pwa)
        cams_disponiveis = ["Todos"] + sorted(agrupado_stc['CAM'].unique().tolist())
        cam_escolhido = st.selectbox("Filtrar por CAM (Bloco 3)", cams_disponiveis)
        agrupado_filtrado = agrupado_stc if cam_escolhido == "Todos" else agrupado_stc[agrupado_stc['CAM'] == cam_escolhido]
        if not agrupado_filtrado.empty:
            st.dataframe(agrupado_filtrado.style.set_properties(**{'text-align': 'left'}), use_container_width=True)
        else:
            st.info("Nenhuma RM encontrada com STC sem expedição.")

    # ===============================
    # 🔹 Exportação Excel
    # ===============================
    def to_excel(dfs, names):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for df, name in zip(dfs, names):
                df.to_excel(writer, sheet_name=name, index=False)
        return output.getvalue()

    with st.expander("📥 Exportar resultados"):
        if st.button("Baixar resultados em Excel"):
            excel_bytes = to_excel(
                [df_capa_completa, df_capa_incompleta, agrupado_mapa, agrupado_stc],
                ["Atendidas_CAM_CAPA", "Pendentes_CAM_CAPA", "MAPA_sem_STC", "STC_nao_expedido"]
            )
            st.download_button(
                label="Clique aqui para baixar o Excel",
                data=excel_bytes,
                file_name="resultado_rm_completo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
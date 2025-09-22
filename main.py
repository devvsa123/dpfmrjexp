import streamlit as st
import pandas as pd
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json

st.set_page_config(page_title="Controle de RM atendidas", layout="wide")
st.title("📦 Controle de RMs - Estocagem e Expedição")
st.markdown("O sistema organiza as RMs em três blocos: conferência por lote, MAPA sem STC e STC não expedido.")

# ===============================
# 🔹 Upload dos arquivos
# ===============================
with st.expander("📄 Upload de arquivos"):
    singra_file = st.file_uploader("Upload planilha do SINGRA (.csv)", type=["csv"])
    pwa_file = st.file_uploader("Upload planilha do PWA (.xlsx)", type=["xlsx"])

if singra_file and pwa_file:

    # ===============================
    # 🔹 Ler CSV do SINGRA
    # ===============================
    df_singra = pd.read_csv(singra_file, sep=';', encoding='latin1')
    df_singra.columns = df_singra.columns.str.replace("'", "").str.strip()
    df_singra = df_singra.fillna('')
    for col in ['ID', 'OMS', 'LISTA_WMS_ID']:
        if col in df_singra.columns:
            df_singra[col] = df_singra[col].astype(str).str.replace("'", "").str.replace('"', '').str.strip()

    # ===============================
    # 🔹 Ler planilha PWA
    # ===============================
    df_pwa = pd.read_excel(pwa_file, sheet_name=0)
    df_pwa = df_pwa.fillna('')
    for col in ['PEDIDO', 'LOTE', 'CAPA', 'MAPA', 'STC', 'STATUS', 'CAM']:
        if col in df_pwa.columns:
            df_pwa[col] = df_pwa[col].astype(str).str.replace("'", "").str.replace('"', '').str.strip()

    # Forçar MAPA como inteiro
    if 'MAPA' in df_pwa.columns:
        df_pwa['MAPA'] = df_pwa['MAPA'].apply(lambda x: str(int(float(x))) if x not in ['', None] else '')

    # ===============================
    # 🔹 Ler LOTE do Google Sheets via secrets.toml
    # ===============================
    # 🔑 Lendo credenciais do secrets
    service_account_dict = dict(st.secrets["gcp_service_account"])

    # 🔹 Escopos de acesso ao Google Sheets e Drive
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

    # 🔹 Criando credenciais a partir do dicionário
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(service_account_dict, scope)

    # 🔹 Autenticando no Google Sheets
    client = gspread.authorize(credentials)

    spreadsheet = client.open_by_url(
        "https://docs.google.com/spreadsheets/d/1naVnAlUGmeAMb_YftLGYit-1e1BcYFJgiJwSnOcgJf4/edit?gid=0"
    )
    worksheet = spreadsheet.get_worksheet(0)
    data = worksheet.get_all_records()
    df_lotes_user = pd.DataFrame(data)
    df_lotes_user['LOTE'] = df_lotes_user['LOTE'].astype(str).str.strip()

    # 🔹 Pré-processar df_pwa logo após o upload
    df_pwa['PEDIDO_LIMPO'] = df_pwa['PEDIDO'].astype(str).str.replace(".", "", regex=False)

    # ===============================
    # 🔹 Nova aba: Consulta rápida de RMs via texto
    # ===============================
    st.markdown("### 🔍 Consulta rápida de RMs (via texto)")

    with st.expander("Consultar RMs colando mensagem"):
        texto_rms = st.text_area("Cole aqui a mensagem com as RMs", height=200)

        if st.button("🔎 Consultar RMs no sistema"):
            if texto_rms.strip():
                import re

                # Regex para capturar números no formato 99.999.999
                rms_extraidas = re.findall(r"\b\d{2}\.\d{3}\.\d{3}\b", texto_rms)

                if rms_extraidas:
                    # Normalizar para comparar
                    rms_sem_ponto = [rm.replace(".", "") for rm in rms_extraidas]

                    # Filtrar de uma vez só em vez de loop
                    df_filtro = df_pwa[df_pwa['PEDIDO_LIMPO'].isin(rms_sem_ponto)]

                    resultados = []
                    for rm, rm_limpo in zip(rms_extraidas, rms_sem_ponto):
                        dados_rm = df_filtro[df_filtro['PEDIDO_LIMPO'] == rm_limpo]

                        if not dados_rm.empty:
                            mapa = ', '.join(dados_rm['MAPA'].unique()) if any(dados_rm['MAPA'] != '') else "Não consta"
                            stc = ', '.join(dados_rm['STC'].unique()) if any(dados_rm['STC'] != '') else "Não consta"
                        else:
                            mapa = "Não consta"
                            stc = "Não consta"

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
    # 🔹 BLOCO 1: CAPA completamente atendidas (considerando RM sem MAPA)
    # ===============================
    capa_completa = []
    capa_incompleta = []

    # Pegar todas as CAPA únicas do SINGRA
    capas_unicas = df_singra['LISTA_WMS_ID'].unique() if 'LISTA_WMS_ID' in df_singra.columns else []

    # Lista de LOTES disponíveis na planilha de conferência
    lotes_disponiveis = df_lotes_user['LOTE'].unique().tolist()

    for capa in capas_unicas:
        # Pegar todas as RM desta CAPA
        rms_da_capa = df_singra[df_singra['LISTA_WMS_ID'] == capa]['ID'].unique()
        
        # Filtrar apenas as RM que ainda não possuem MAPA
        rms_sem_mapa = [rm for rm in rms_da_capa if df_pwa[(df_pwa['PEDIDO'] == rm)]['MAPA'].eq('').any()]

        # Se todas RM têm MAPA, ignorar essa CAPA
        if not rms_sem_mapa:
            continue

        todos_lotes_capa = []
        for rm in rms_sem_mapa:
            lotes_rm = df_pwa[df_pwa['PEDIDO'] == rm]['LOTE'].unique().tolist()
            todos_lotes_capa.extend(lotes_rm)

        # Verificar se todos os lotes da CAPA estão presentes na conferência
        faltando_lotes = [l for l in todos_lotes_capa if l not in lotes_disponiveis]

        # Pegar CAM (considerando que todas RM da CAPA tenham o mesmo CAM)
        cam = df_singra.loc[df_singra['LISTA_WMS_ID'] == capa, 'OMS'].values[0] if 'OMS' in df_singra.columns else ''

        if not faltando_lotes:
            capa_completa.append({"CAM": cam, "CAPA": capa, "RMs": ', '.join(rms_sem_mapa)})
        else:
            capa_incompleta.append({"CAM": cam, "CAPA": capa, "RMs": ', '.join(rms_sem_mapa), "LOTES_FALTANDO": ', '.join(faltando_lotes)})

    # Resumo
    st.markdown("### 📊 Resumo Conferência por CAPA")
    col1, col2 = st.columns(2)
    col1.success(f"✅ Total de CAPA completamente atendidas (sem MAPA): {len(capa_completa)}")
    col2.warning(f"⚠️ Total de CAPA parcialmente atendidas (pendências): {len(capa_incompleta)}")

    # Mostrar blocos
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
    st.markdown("### 📋 RMs com MAPA porém sem STC")

    if all(col in df_pwa.columns for col in ['MAPA', 'STC', 'STATUS', 'CAM', 'CAPA']):
        df_mapa_sem_stc = df_pwa[
            (df_pwa['MAPA'] != '') &
            (df_pwa['STC'] == '') &
            (df_pwa['STATUS'] != 'EXPEDIDO')
        ]

        if not df_mapa_sem_stc.empty:
            # Agrupar por CAM e MAPA → trazer as CAPAs relacionadas
            agrupado_mapa = (
                df_mapa_sem_stc.groupby(['CAM', 'MAPA'])
                .agg({
                    'CAPA': lambda x: ', '.join(sorted(set(x)))
                })
                .reset_index()
            )

            # 🔹 Filtro por CAM
            cams_disponiveis = ["Todos"] + sorted(agrupado_mapa['CAM'].unique().tolist())
            cam_escolhido = st.selectbox("Filtrar por CAM", cams_disponiveis)

            if cam_escolhido != "Todos":
                agrupado_filtrado = agrupado_mapa[agrupado_mapa['CAM'] == cam_escolhido]
            else:
                agrupado_filtrado = agrupado_mapa

            # 🔹 Mostrar dataframe compacto
            st.dataframe(
                agrupado_filtrado.style.set_properties(**{'text-align': 'left'}),
                use_container_width=True
            )

        else:
            st.info("Nenhuma RM encontrada com MAPA sem STC.")

    # ===============================
    # 🔹 BLOCO 3: RMs com STC mas não expedidas
    # ===============================
    st.markdown("### 🚚 RMs com STC porém não expedidas")

    if all(col in df_pwa.columns for col in ['STC', 'STATUS', 'CAM', 'CAPA']):
        df_stc_nao_expedida = df_pwa[
            (df_pwa['STC'] != '') &
            (df_pwa['STATUS'] != 'EXPEDIDO') &
            (df_pwa['STATUS'] != 'CANCELADO')
        ]

        if not df_stc_nao_expedida.empty:
            # Agrupar por CAM e STC → trazer todos os MAPAs
            agrupado_stc = (
                df_stc_nao_expedida.groupby(['CAM', 'STC'])
                .agg({
                    'MAPA': lambda x: ', '.join(sorted(set(x)))
                })
                .reset_index()
            )

            # 🔹 Filtro por CAM
            cams_disponiveis = ["Todos"] + sorted(agrupado_stc['CAM'].unique().tolist())
            cam_escolhido = st.selectbox("Filtrar por CAM (Bloco 3)", cams_disponiveis)

            if cam_escolhido != "Todos":
                agrupado_filtrado = agrupado_stc[agrupado_stc['CAM'] == cam_escolhido]
            else:
                agrupado_filtrado = agrupado_stc

            # 🔹 Mostrar dataframe
            st.dataframe(
                agrupado_filtrado.style.set_properties(**{'text-align': 'left'}),
                use_container_width=True
            )

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
                [agrupado_att, agrupado_pend, agrupado_mapa, agrupado_stc],
                ["Atendidas_CAM_CAPA", "Pendentes_CAM_CAPA", "MAPA_sem_STC", "STC_nao_expedido"]
            )
            st.download_button(
                label="Clique aqui para baixar o Excel",
                data=excel_bytes,
                file_name="resultado_rm_completo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

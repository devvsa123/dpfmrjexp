import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Controle de RM atendidas", layout="wide")

st.title("📦 Controle de RMs - Estocagem e Expedição")
st.markdown("O sistema organiza as RMs em três blocos: conferência por lote, MAPA sem STC e STC não expedido.")

# ===============================
# 🔹 Upload dos arquivos
# ===============================
with st.expander("📄 Upload de arquivos"):
    singra_file = st.file_uploader("Upload planilha do SINGRA (.csv)", type=["csv"])
    pwa_file = st.file_uploader("Upload planilha do PWA (.xlsx)", type=["xlsx"])
    lotes_file = st.file_uploader("Upload planilha com LOTES do usuário (.xlsx)", type=["xlsx"])

if singra_file and pwa_file and lotes_file:

    # === Ler CSV do SINGRA ===
    df_singra = pd.read_csv(singra_file, sep=';', encoding='latin1')
    df_singra.columns = df_singra.columns.str.replace("'", "").str.strip()
    df_singra = df_singra.fillna('')
    
    for col in ['ID', 'OMS', 'LISTA_WMS_ID']:
        if col in df_singra.columns:
            df_singra[col] = df_singra[col].astype(str).str.replace("'", "").str.replace('"', '').str.strip()

    # === Ler planilha PWA ===
    df_pwa = pd.read_excel(pwa_file, sheet_name=0)
    df_pwa = df_pwa.fillna('')
    
    for col in ['PEDIDO', 'LOTE', 'CAPA', 'MAPA', 'STC', 'STATUS', 'CAM']:
        if col in df_pwa.columns:
            df_pwa[col] = df_pwa[col].astype(str).str.replace("'", "").str.replace('"', '').str.strip()

    # Forçar MAPA como inteiro
    if 'MAPA' in df_pwa.columns:
        df_pwa['MAPA'] = df_pwa['MAPA'].apply(lambda x: str(int(float(x))) if x not in ['', None] else '')

    # === Ler planilha de lotes do usuário ===
    df_lotes_user = pd.read_excel(lotes_file, sheet_name=0)
    df_lotes_user['LOTE'] = df_lotes_user['LOTE'].astype(str).str.strip()

    # ===============================
    # 🔹 BLOCO 1: Conferência por LOTE
    # ===============================
    atendidas = []
    pendentes = []

    rms_unicas = df_singra['ID'].unique()

    for rm in rms_unicas:
        # Ignorar RMs que já possuem MAPA
        if 'MAPA' in df_pwa.columns and not df_pwa[df_pwa['PEDIDO'] == rm]['MAPA'].eq('').all():
            continue

        lotes_rm = df_pwa[df_pwa['PEDIDO'] == rm]['LOTE'].unique().tolist()
        lotes_usuario = df_lotes_user['LOTE'].unique().tolist()
        lotes_presentes = [l for l in lotes_rm if l in lotes_usuario]

        cam = df_singra.loc[df_singra['ID'] == rm, 'OMS'].values[0] if 'OMS' in df_singra.columns else ''
        capa = df_singra.loc[df_singra['ID'] == rm, 'LISTA_WMS_ID'].values[0] if 'LISTA_WMS_ID' in df_singra.columns else ''

        if set(lotes_presentes) == set(lotes_rm):
            atendidas.append({"RM": rm, "OMS/CAM": cam, "CAPA": capa})
        else:
            faltam = list(set(lotes_rm) - set(lotes_presentes))
            pendentes.append({"RM": rm, "OMS/CAM": cam, "CAPA": capa, "LOTES_FALTANDO": ', '.join(faltam)})

    # Resumo
    st.markdown("### 📊 Resumo Conferência por Lotes")
    col1, col2 = st.columns(2)
    col1.success(f"✅ Total de RMs totalmente atendidas: {len(atendidas)}")
    col2.warning(f"⚠️ Total de RMs parcialmente atendidas: {len(pendentes)}")

    # Mostrar blocos
    st.subheader("✅ RMs totalmente atendidas (agrupadas por CAM e CAPA)")
    df_att = pd.DataFrame(atendidas)
    if not df_att.empty:
        agrupado_att = df_att.groupby(['OMS/CAM', 'CAPA'])['RM'].apply(lambda x: ', '.join(x.astype(str))).reset_index()
        st.dataframe(agrupado_att.style.set_properties(**{'text-align': 'left'}))
    else:
        st.info("Nenhuma RM totalmente atendida encontrada.")

    st.subheader("⚠️ RMs parcialmente atendidas (agrupadas por CAM e CAPA)")
    df_pend = pd.DataFrame(pendentes)
    if not df_pend.empty:
        agrupado_pend = df_pend.groupby(['OMS/CAM', 'CAPA']).apply(
            lambda x: pd.Series({
                'RMs': ', '.join(x['RM'].astype(str)),
                'LOTES_FALTANDO': '; '.join(x['LOTES_FALTANDO'])
            })
        ).reset_index()
        st.dataframe(agrupado_pend.style.set_properties(**{'text-align': 'left'}))
    else:
        st.info("Nenhuma RM parcialmente atendida encontrada.")

    # ===============================
    # 🔹 BLOCO 2: RMs com MAPA sem STC
    # ===============================
    st.markdown("### 📋 RMs com MAPA porém sem STC")
    if 'MAPA' in df_pwa.columns and 'STC' in df_pwa.columns and 'STATUS' in df_pwa.columns and 'CAM' in df_pwa.columns and 'CAPA' in df_pwa.columns:
        df_mapa_sem_stc = df_pwa[(df_pwa['MAPA'] != '') & (df_pwa['STC'] == '') & (df_pwa['STATUS'] != 'EXPEDIDO')]
        if not df_mapa_sem_stc.empty:
            agrupado_mapa = df_mapa_sem_stc.groupby(['CAM', 'CAPA']).agg({
                'PEDIDO': lambda x: ', '.join(sorted(set(x))),
                'MAPA': lambda x: ', '.join(sorted(set(x)))
            }).reset_index()
            st.dataframe(agrupado_mapa.style.set_properties(**{'text-align': 'left'}))
        else:
            st.info("Nenhuma RM encontrada com MAPA sem STC.")

    # ===============================
    # 🔹 BLOCO 3: RMs com STC mas não expedidas
    # ===============================
    st.markdown("### 🚚 RMs com STC porém não expedidas")
    if 'STC' in df_pwa.columns and 'STATUS' in df_pwa.columns and 'CAM' in df_pwa.columns and 'CAPA' in df_pwa.columns:
        df_stc_nao_expedida = df_pwa[(df_pwa['STC'] != '') & (df_pwa['STATUS'] != 'EXPEDIDO') & (df_pwa['STATUS'] != 'CANCELADO')]
        if not df_stc_nao_expedida.empty:
            agrupado_stc = df_stc_nao_expedida.groupby(['CAM', 'CAPA']).agg({
                'PEDIDO': lambda x: ', '.join(sorted(set(x))),
                'MAPA': lambda x: ', '.join(sorted(set(x))),
                'STC': lambda x: ', '.join(sorted(set(x))),
                'STATUS': lambda x: ', '.join(sorted(set(x)))
            }).reset_index()
            st.dataframe(agrupado_stc.style.set_properties(**{'text-align': 'left'}))
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
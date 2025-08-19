import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Controle de RM atendidas", layout="wide")

st.title("📦 Verificação de RMs totalmente atendidas por lote")
st.markdown("Este sistema verifica quais RMs estão totalmente ou parcialmente atendidas com base nos lotes enviados pelo usuário.")

# Upload dos arquivos
with st.expander("📄 Upload de arquivos"):
    singra_file = st.file_uploader("Upload planilha do SINGRA (.csv)", type=["csv"])
    pwa_file = st.file_uploader("Upload planilha do PWA (.xlsx)", type=["xlsx"])
    lotes_file = st.file_uploader("Upload planilha com LOTES do usuário (.xlsx)", type=["xlsx"])

if singra_file and pwa_file and lotes_file:

    # Ler CSV do SINGRA
    df_singra = pd.read_csv(singra_file, sep=';', encoding='latin1')
    df_singra.columns = df_singra.columns.str.replace("'", "").str.strip()

    # Ler arquivos xlsx
    df_pwa = pd.read_excel(pwa_file, sheet_name=0)
    df_lotes_user = pd.read_excel(lotes_file, sheet_name=0)

    # Normalizar LOTE
    df_lotes_user['LOTE'] = df_lotes_user['LOTE'].astype(str).str.strip()
    df_pwa['LOTE'] = df_pwa['LOTE'].astype(str).str.strip()

    atendidas = []
    pendentes = []

    rms_unicas = df_singra['ID'].unique()

    for rm in rms_unicas:
        lotes_rm = df_pwa[df_pwa['PEDIDO'] == rm]['LOTE'].astype(str).str.strip().unique().tolist()
        lotes_usuario = df_lotes_user['LOTE'].astype(str).str.strip().unique().tolist()
        lotes_presentes = [l for l in lotes_rm if l in lotes_usuario]

        cam = df_singra.loc[df_singra['ID'] == rm, 'OMS'].values[0]

        if set(lotes_presentes) == set(lotes_rm):
            atendidas.append({"RM": rm, "OMS/CAM": cam})
        else:
            faltam = list(set(lotes_rm) - set(lotes_presentes))
            pendentes.append({"RM": rm, "OMS/CAM": cam, "LOTES_FALTANDO": ', '.join(faltam)})

    # === INDICADORES RÁPIDOS ===
    total_att = len(atendidas)
    total_pend = len(pendentes)
    st.markdown("### 📊 Resumo")
    col1, col2 = st.columns(2)
    col1.success(f"✅ Total de RMs totalmente atendidas: {total_att}")
    col2.warning(f"⚠️ Total de RMs parcialmente atendidas: {total_pend}")

    # === ATENDIDAS AGRUPADAS POR CAM ===
    st.subheader("✅ RMs totalmente atendidas por CAM")
    df_att = pd.DataFrame(atendidas)
    if not df_att.empty:
        agrupado_att = df_att.groupby('OMS/CAM')['RM'].apply(lambda x: ', '.join(x.astype(str))).reset_index()
        st.dataframe(agrupado_att.style.set_properties(**{'text-align': 'left'}))
    else:
        st.info("Nenhuma RM totalmente atendida encontrada.")

    # === PENDENTES AGRUPADAS POR CAM ===
    st.subheader("⚠️ RMs parcialmente atendidas por CAM")
    df_pend = pd.DataFrame(pendentes)
    if not df_pend.empty:
        agrupado_pend = df_pend.groupby('OMS/CAM').apply(
            lambda x: pd.Series({
                'RMs': ', '.join(x['RM'].astype(str)),
                'LOTES_FALTANDO': '; '.join(x['LOTES_FALTANDO'])
            })
        ).reset_index()
        st.dataframe(agrupado_pend.style.set_properties(**{'text-align': 'left'}))
    else:
        st.info("Nenhuma RM parcialmente atendida encontrada.")

    # Exportação em Excel
    def to_excel(df1, df2):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df1.to_excel(writer, sheet_name='Atendidas_por_CAM', index=False)
            df2.to_excel(writer, sheet_name='Pendentes_por_CAM', index=False)
        return output.getvalue()

    with st.expander("📥 Exportar resultados"):
        if st.button("Baixar resultados em Excel"):
            excel_bytes = to_excel(agrupado_att, agrupado_pend)
            st.download_button(
                label="Clique aqui para baixar o Excel",
                data=excel_bytes,
                file_name="resultado_rm.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

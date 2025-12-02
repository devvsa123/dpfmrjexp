import streamlit as st
import pandas as pd

st.set_page_config(page_title="PWA – Controle", layout="wide")

# ============================================================
#  FUNÇÃO PARA CARREGAR ARQUIVO PWA COM TRATAMENTO CORRETO
# ============================================================

@st.cache_data
def carregar_pwa(arquivo):
    try:
        df = pd.read_excel(arquivo)

        # Lista real de colunas existentes no seu arquivo
        colunas_validas = [
            "PEDIDO", "STATUS", "STC", "MAPA", "VOLUME", "CAM",
            "CAPA", "LOTE", "PI", "NOMENCLATURA", "QTD"
        ]

        # Mantém apenas colunas existentes no arquivo
        colunas_presentes = [col for col in colunas_validas if col in df.columns]

        # Limpa colunas presentes, sem causar erro
        for col in colunas_presentes:
            df[col] = df[col].astype(str).str.strip()

        return df

    except Exception as e:
        st.error(f"Erro ao carregar o arquivo PWA: {str(e)}")
        return pd.DataFrame()


# ============================================================
#  INTERFACE
# ============================================================

st.title("📦 Sistema de Controle PWA")
st.write("Carregue o arquivo PWA para iniciar o processamento.")

arquivo_pwa = st.file_uploader("Selecione o arquivo Excel do PWA", type=["xlsx"])

if arquivo_pwa:
    df_pwa = carregar_pwa(arquivo_pwa)

    if df_pwa.empty:
        st.warning("⚠ O arquivo foi carregado, mas não contém dados válidos.")
    else:
        st.success("Arquivo carregado com sucesso!")

        st.subheader("📄 Visualização dos Dados")
        st.dataframe(df_pwa, use_container_width=True)

        st.write(f"Total de linhas carregadas: **{len(df_pwa)}**")

        # BOTÃO DE DOWNLOAD (Opcional)
        csv = df_pwa.to_csv(index=False).encode("utf-8")
        st.download_button(
            "📥 Baixar Dados Tratados (CSV)",
            data=csv,
            file_name="pwa_limpo.csv",
            mime="text/csv"
        )

else:
    st.info("Envie um arquivo Excel (.xlsx) para começar.")

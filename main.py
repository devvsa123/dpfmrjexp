# ---------------------------------------
# FUNÇÃO PARA NORMALIZAR CAMPOS NUMÉRICOS
# ---------------------------------------
def normalizar_codigo(valor):
    """
    Converte qualquer valor numérico ou texto para string de 8 dígitos.
    """
    try:
        return f"{int(str(valor).replace('.', '').replace(',', '').strip()):08d}"
    except:
        return str(valor).strip()


# -------------------------------------------------
# NORMALIZAÇÃO DAS COLUNAS IMPORTANTES DO df_pwa
# -------------------------------------------------
colunas_normalizar = ['CAM', 'CAPA', 'MAPA', 'STC', 'PEDIDO']

for col in colunas_normalizar:
    if col in df_pwa.columns:
        df_pwa[col] = df_pwa[col].astype(str).apply(normalizar_codigo)


# ============================================================
# BLOCOS DO SISTEMA
# ============================================================


# ----------------------------------------------------------------
# 📌 BLOCO 2 — RMs COM MAPA MAS SEM STC (EVITAR EXPEDIDOS)
# ----------------------------------------------------------------
st.markdown("### 📋 RMs com MAPA porém sem STC")

if all(col in df_pwa.columns for col in ['MAPA', 'STC', 'STATUS', 'CAM', 'CAPA']):

    # 🔥 AGORA NÃO MOSTRA RM EXPEDIDO
    df_mapa_sem_stc = df_pwa[
        (df_pwa['MAPA'] != '') &
        (df_pwa['STC'] == '') &
        (df_pwa['STATUS'].str.upper() != 'EXPEDIDO')
    ]

    if not df_mapa_sem_stc.empty:

        # -----------------------------
        # 🎯 FILTRO POR CAM (otimizado)
        # -----------------------------
        cams_disponiveis = sorted(df_mapa_sem_stc['CAM'].unique())
        cam_selecionado = st.selectbox("Selecione o CAM:", cams_disponiveis)

        df_filtrado = df_mapa_sem_stc[df_mapa_sem_stc['CAM'] == cam_selecionado]

        if not df_filtrado.empty:

            # ---------------------------------------------------
            # AGRUPAMENTO POR CAM E MAPA + LISTAGEM DAS CAPAS
            # ---------------------------------------------------
            agrupado = (
                df_filtrado
                .groupby(['CAM', 'MAPA'])
                .agg({
                    'CAPA': lambda x: ', '.join(sorted(set(x)))
                })
                .reset_index()
            )

            st.dataframe(
                agrupado.style.set_properties(**{'text-align': 'left'})
            )

        else:
            st.info("Nenhuma RM encontrada para o CAM selecionado.")

    else:
        st.info("Nenhuma RM encontrada com MAPA sem STC.")



# ----------------------------------------------------------------
# 📦 BLOCO 3 — RMs COM STC MAS NÃO EXPEDIDAS
# ----------------------------------------------------------------
st.markdown("### 🚚 RMs com STC porém não expedidas")

if all(col in df_pwa.columns for col in ['STC', 'STATUS', 'CAM']):

    df_stc_nao_expedida = df_pwa[
        (df_pwa['STC'] != '') &
        (df_pwa['STATUS'].str.upper() != 'EXPEDIDO') &
        (df_pwa['STATUS'].str.upper() != 'CANCELADO')
    ]

    if not df_stc_nao_expedida.empty:

        # 🔽 Filtro por CAM
        cams_disponiveis_3 = sorted(df_stc_nao_expedida['CAM'].unique())
        cam_selecionado_3 = st.selectbox("Selecione o CAM (Bloco 3):", cams_disponiveis_3)

        df_filtrado_3 = df_stc_nao_expedida[df_stc_nao_expedida['CAM'] == cam_selecionado_3]

        if not df_filtrado_3.empty:

            # ----------------------------------------------------------
            # AGRUPAMENTO POR CAM E STC + TODOS OS MAPAS DAQUELE STC
            # ----------------------------------------------------------
            agrupado_stc = (
                df_filtrado_3
                .groupby(['CAM', 'STC'])
                .agg({
                    'MAPA': lambda x: ', '.join(sorted(set(x)))
                })
                .reset_index()
            )

            st.dataframe(
                agrupado_stc.style.set_properties(**{'text-align': 'left'})
            )

        else:
            st.info("Nenhum registro para o CAM selecionado.")

    else:
        st.info("Nenhuma RM encontrada com STC sem expedição.")

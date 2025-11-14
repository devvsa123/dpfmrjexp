# ============================================================
#   BLOCO 1 – CAPAS COMPLETAMENTE ATENDIDAS / NÃO ATENDIDAS
#   (PWA = fonte de verdade | SINGRA = conferência da expedição)
# ============================================================

st.markdown("## 📦 CAPAS – Situação Geral (Base PWA como verdade)")

# Verifica se as colunas necessárias existem
col_required_pwa = ['PEDIDO','STATUS','CAM','CAPA','LOTE']
col_required_singra = ['RM','LOTE']

if not all(c in df_pwa.columns for c in col_required_pwa):
    st.error("❌ Algumas colunas essenciais estão faltando no df_pwa.")
elif not all(c in df_lotes_user.columns for c in col_required_singra):
    st.error("❌ Algumas colunas essenciais estão faltando na planilha de lotes (SINGRA).")
else:

    capas = df_pwa['CAPA'].unique()

    capas_completas = []
    capas_incompletas = []

    for capa in capas:

        df_capa = df_pwa[df_pwa['CAPA'] == capa]
        cam = df_capa['CAM'].iloc[0]
        rms_capa = df_capa['PEDIDO'].unique()

        pendencias_rm = []

        for rm in rms_capa:

            # Lotes do PWA (verdade absoluta)
            lotes_pwa = df_pwa[df_pwa['PEDIDO'] == rm]['LOTE'].unique()

            # Lotes no SINGRA (registro da chegada)
            lotes_singra = df_lotes_user[df_lotes_user['RM'] == rm]['LOTE'].unique()

            # Se RM não existe no SINGRA → pendência crítica
            if len(lotes_singra) == 0:
                pendencias_rm.append(f"{rm} (Status SINGRA não migrou)")
                continue

            # Verificar se todos os lotes da RM estão no SINGRA
            for lote in lotes_pwa:
                if lote not in lotes_singra:
                    pendencias_rm.append(f"{rm} – faltando lote {lote}")
                    break

        # DECISÃO SOBRE A CAPA
        if len(pendencias_rm) == 0:
            capas_completas.append([cam, capa, ", ".join(rms_capa)])
        else:
            capas_incompletas.append([cam, capa, ", ".join(pendencias_rm)])

    # --------------------------------------------------------------
    #   EXIBIR CAPAS COMPLETAS
    # --------------------------------------------------------------
    st.markdown("### ✅ CAPAS completamente atendidas")

    if len(capas_completas) > 0:
        df_ok = pd.DataFrame(capas_completas, columns=['CAM','CAPA','RMs'])
        st.dataframe(df_ok.style.set_properties(**{'text-align':'left'}))
    else:
        st.info("Nenhuma CAPA completamente atendida no momento.")

    # --------------------------------------------------------------
    #   EXIBIR CAPAS COM PENDÊNCIAS
    # --------------------------------------------------------------
    st.markdown("### ❌ CAPAS com pendências")

    if len(capas_incompletas) > 0:
        df_bad = pd.DataFrame(capas_incompletas, columns=['CAM','CAPA','Pendências'])
        st.dataframe(df_bad.style.set_properties(**{'text-align':'left'}))
    else:
        st.success("Nenhuma pendência encontrada.")

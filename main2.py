import streamlit as st
import pandas as pd
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re

st.set_page_config(page_title="Controle de RM atendidas", layout="wide")
st.title("📦 Controle de RMs - Estocagem e Expedição")
st.markdown("Sistema: PWA = fonte da verdade. BLOCO 1 agora verifica por VOLUME (planilha LOTE contém volumes presentes na expedição).")

# ----------------------
# Utilitários / Normalização
# ----------------------
def clean_colnames(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    newcols = []
    for c in df.columns:
        c2 = str(c).replace("'", "").replace('"', '').strip()
        newcols.append(c2)
    df.columns = newcols
    return df

def normalizar_codigo_rm(valor):
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
    if pd.isna(val) or str(val).strip() == '':
        return False
    v = str(val).strip().upper()
    return ('EXPED' in v) or ('EM EXPED' in v) or ('EXPEDIÇÃO' in v) or ('EXPEDICAO' in v)

# ----------------------
# Cache: carregamento arquivos
# ----------------------
@st.cache_data
def carregar_singra(file):
    df = pd.read_csv(file, sep=';', encoding='latin1', dtype=str, low_memory=False)
    df = clean_colnames(df)
    df = df.fillna('')
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
    # Limpar somente colunas que existem
    for col in ['PEDIDO', 'CAPA', 'MAPA', 'STC', 'CAM', 'LOTE', 'STATUS', 'VOLUME', 'PI', 'NOMENCLATURA', 'QTD']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    # PEDIDO limpo para comparar (remove pontos e espaços)
    if 'PEDIDO' in df.columns:
        df['PEDIDO_LIMPO'] = df['PEDIDO'].apply(normalizar_codigo_rm)
    else:
        df['PEDIDO'] = ''
        df['PEDIDO_LIMPO'] = ''
    # Normalize MAPA to integer-like string
    if 'MAPA' in df.columns:
        def mapa_to_intstr(x):
            x = str(x).strip()
            if x == '' or x.upper() == 'NAN':
                return ''
            try:
                if '.' in x:
                    return str(int(float(x)))
                return x
            except:
                return x
        df['MAPA'] = df['MAPA'].apply(mapa_to_intstr)
    # Upper STATUS
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
        df['LOTE'] = df['LOTE'].astype(str).str.strip()
    return df

# ----------------------
# UI: Uploads
# ----------------------
with st.expander("📄 Upload de arquivos"):
    singra_file = st.file_uploader("Upload planilha do SINGRA (.csv)", type=["csv"])
    pwa_file = st.file_uploader("Upload planilha do PWA (.xlsx)", type=["xlsx"])

if not (singra_file and pwa_file):
    st.info("Faça upload do SINGRA (.csv) e do PWA (.xlsx) para prosseguir.")
    st.stop()

# ----------------------
# Carrega dados (cached)
# ----------------------
df_singra = carregar_singra(singra_file)
df_pwa = carregar_pwa(pwa_file)

# Carregar planilha de lotes (Google Sheets)
SHEET_URL = "https://docs.google.com/spreadsheets/d/1naVnAlUGmeAMb_YftLGYit-1e1BcYFJgiJwSnOcgJf4/edit?gid=0"
service_account_dict = dict(st.secrets["gcp_service_account"])
df_lotes_user = carregar_lotes_google(service_account_dict, SHEET_URL)

# ----------------------
# PREP: volumes presentes na expedição (planilha LOTE)
# ----------------------
# Nota: a coluna 'LOTE' na planilha Google contém os números de VOLUME (um por linha)
volumes_expedicao = set()
if 'LOTE' in df_lotes_user.columns:
    volumes_expedicao = set(df_lotes_user['LOTE'].astype(str).str.strip().tolist())
else:
    volumes_expedicao = set()

# ----------------------
# Map SINGRA: RM -> {SITUACAO, OMS}
# ----------------------
singra_map = {}
if 'ID' in df_singra.columns:
    for rm, grp in df_singra.groupby('ID'):
        situ = grp['SITUACAO'].iloc[0] if 'SITUACAO' in df_singra.columns else ''
        oms = grp['OMS'].iloc[0] if 'OMS' in df_singra.columns else ''
        singra_map[rm] = {'SITUACAO': situ, 'OMS': oms}

# ----------------------
# Precompute PWA maps for performance
# ----------------------
# 1) pwa_lotes_map: RM -> list of LOTES (strings)
pwa_lotes_map = {pedido: sorted(set(g['LOTE'].astype(str).tolist())) for pedido, g in df_pwa.groupby('PEDIDO_LIMPO')}

# 2) capa_to_rms: CAPA -> list of RMs
capa_to_rms = {capa: sorted(df_pwa[df_pwa['CAPA'] == capa]['PEDIDO_LIMPO'].unique().tolist()) for capa in sorted(df_pwa['CAPA'].unique().tolist())}

# 3) lote -> volumes previstos (set)
lote_to_volumes_previstos = {}
if 'LOTE' in df_pwa.columns and 'VOLUME' in df_pwa.columns:
    for lote, grp in df_pwa.groupby('LOTE'):
        vols = set([str(v).strip() for v in grp['VOLUME'].astype(str).tolist() if str(v).strip() != '' and str(v).strip().upper() != 'NAN'])
        # if empty, we mark as unknown (empty set) and will treat as pendência (precisa definir volumes)
        lote_to_volumes_previstos[lote] = vols
else:
    st.error("PWA precisa ter as colunas 'LOTE' e 'VOLUME'.")
    st.stop()

# Quick metrics
c1, c2, c3 = st.columns(3)
c1.metric("RMs únicas (PWA)", df_pwa['PEDIDO_LIMPO'].nunique() if 'PEDIDO_LIMPO' in df_pwa.columns else 0)
c2.metric("Linhas PWA", len(df_pwa))
c3.metric("Volumes na planilha (Google)", len(volumes_expedicao))

# ----------------------
# Consulta rápida via texto (mantive)
# ----------------------
st.markdown("### 🔍 Consulta rápida de RMs (via texto)")
with st.expander("Consultar RMs colando mensagem"):
    texto_rms = st.text_area("Cole aqui a mensagem com as RMs", height=180)
    if st.button("🔎 Consultar RMs no sistema"):
        if texto_rms.strip():
            rms_extraidas = re.findall(r"\b\d{2}\.\d{3}\.\d{3}\b", texto_rms)
            if rms_extraidas:
                rms_sem_ponto = [rm.replace(".", "") for rm in rms_extraidas]
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
# BLOCO 1 – NOVA LÓGICA (somente RMs sem MAPA) -> agora analisando VOLUME
# ----------------------
st.markdown("## 🔵 BLOCO 1 — CAPA: verificação (somente RMs sem MAPA) — conferência por VOLUME")

required_pwa_cols = ['PEDIDO_LIMPO', 'LOTE', 'CAPA', 'CAM', 'STATUS']
if not all(c in df_pwa.columns for c in required_pwa_cols):
    st.error("Colunas essenciais faltando no PWA: preciso de PEDIDO/LOTE/CAPA/CAM/STATUS.")
else:
    capas = sorted(df_pwa['CAPA'].unique().tolist())
    capa_completa_rows = []
    capa_incompleta_rows = []
    migration_errors = []  # tabela separada para RMs que não existem no SINGRA (quando houver volumes na expedição)

    # função que retorna volumes faltantes por lote para a RM (comportamento seguro)
    def volumes_faltantes_para_rm(rm):
        faltantes_por_lote = {}  # lote -> list(missing volumes) or ["UNKNOWN"] if PWA não indica volumes
        lotes_pwa = pwa_lotes_map.get(rm, [])
        for lote in lotes_pwa:
            previstos = lote_to_volumes_previstos.get(lote, set())
            recebidos = set([v for v in volumes_expedicao if v != '' and v is not None])

            # caso PWA não traga volumes previstos (conjunto vazio) -> sinalizar como pendência (UNKNOWN)
            if not previstos:
                # Se PWA não tem volumes para o lote, não podemos checar — marcamos como pendência
                # alternativa: inferir por QTD, mas aqui optamos por considerar pendência para revisão manual
                faltantes_por_lote[lote] = ["UNKNOWN"]
                continue

            # verificar volumes previstos que estão ausentes na lista de volumes na expedição
            missing = [v for v in sorted(previstos) if v not in recebidos]
            if missing:
                faltantes_por_lote[lote] = missing
        return faltantes_por_lote

    # iteração por CAPA
    for capa in capas:
        rms_da_capa = capa_to_rms.get(capa, [])
        cam = df_pwa[df_pwa['CAPA'] == capa]['CAM'].iloc[0] if len(df_pwa[df_pwa['CAPA'] == capa]) > 0 else ''
        pendencias = []
        # track flags to decide full completion
        all_rms_with_volumes_ok = True
        all_rms_migrated_in_singra = True

        # Select only RMs that do NOT have MAPA (same as original BLOCO 1)
        rms_considered = []
        for rm in rms_da_capa:
            df_rm_rows = df_pwa[df_pwa['PEDIDO_LIMPO'] == rm]
            rm_has_mapa = df_rm_rows['MAPA'].apply(lambda x: str(x).strip() != '').any()
            if not rm_has_mapa:
                rms_considered.append(rm)

        if len(rms_considered) == 0:
            # nothing to consider for this capa (all RMs have MAPA)
            continue

        for rm in rms_considered:
            # verificar se rm está no singra
            singra_info = singra_map.get(rm)
            rm_migrated = bool(singra_info and singra_indica_em_expedicao(singra_info.get('SITUACAO', '')))
            if not rm_migrated:
                all_rms_migrated_in_singra = False

            # checar volumes faltantes
            faltantes = volumes_faltantes_para_rm(rm)

            # Se PWA não tiver volumes definidos para algum lote -> considerar pendência
            if faltantes:
                # montar mensagem de pendência por RM
                partes = []
                for lote_key, vols in faltantes.items():
                    if vols == ["UNKNOWN"]:
                        partes.append(f"{lote_key} (volumes não informados no PWA)")
                    else:
                        partes.append(f"{lote_key} – faltando volumes: {', '.join(vols)}")
                pendencias.append(f"{rm} – " + '; '.join(partes))
                all_rms_with_volumes_ok = False
            else:
                # sem faltantes por volume para essa RM
                pass

            # Agora cenários especiais envolvendo volumes presentes mas RM não migrada:
            # - se todos volumes dessa RM estão na expedição (ou seja, faltantes == {}), mas rm_migrated == False:
            #   devemos sinalizar isso (material presente mas rm nao migrou)
            if not rm_migrated:
                # verificar se PWA previa volumes e se todos eles estão na expedição
                lotes_do_rm = pwa_lotes_map.get(rm, [])
                rm_all_volumes_present = True
                for lote_key in lotes_do_rm:
                    previstos = lote_to_volumes_previstos.get(lote_key, set())
                    if not previstos:
                        rm_all_volumes_present = False
                        break
                    # se algum previsto não está na volumes_expedicao -> não está tudo presente
                    for vol in previstos:
                        if vol not in volumes_expedicao:
                            rm_all_volumes_present = False
                            break
                    if not rm_all_volumes_present:
                        break

                # se rm_all_volumes_present == True mas rm_migrated == False -> sinal especial
                if rm_all_volumes_present:
                    # marcar migração ausente como pendência mas com mensagem específica
                    msg = f"{rm} – Todos os volumes de suas remessas estão na Expedição mas RM não migrou no SINGRA"
                    pendencias.append(msg)
                    # também registrar em migration_errors para relatório separado
                    migration_errors.append({
                        "RM": rm,
                        "CAPA": capa,
                        "CAM": cam,
                        "Erro": "Todos volumes na expedição, porém RM não migrou no SINGRA"
                    })
                    all_rms_with_volumes_ok = False  # considera-se pendência até migração

                else:
                    # se volumes parcialmente presentes mas RM não migrou -> sinalizar também
                    # se já houve pendências por volumes, a mensagem já existe; caso contrário, podemos
                    # acrescentar uma mensagem indicando material na expedição mas RM ausente
                    # aqui detectamos se existe ao menos um volume do RM na expedição
                    lotes_do_rm = pwa_lotes_map.get(rm, [])
                    any_volume_present = False
                    for lote_key in lotes_do_rm:
                        previstos = lote_to_volumes_previstos.get(lote_key, set())
                        for vol in previstos:
                            if vol in volumes_expedicao:
                                any_volume_present = True
                                break
                        if any_volume_present:
                            break
                    if any_volume_present:
                        pendencias.append(f"{rm} – Material na Expedição porém RM não consta como Em Expedição no SINGRA")
                        # register in migration_errors as well
                        migration_errors.append({
                            "RM": rm,
                            "CAPA": capa,
                            "CAM": cam,
                            "Erro": "Material na Expedição porém RM não consta em Expedição no SINGRA"
                        })
                        all_rms_with_volumes_ok = False

        # Decide CAPA status: somente considerada completa se não houver pendências e todas RMs migradas
        if len(pendencias) == 0 and all_rms_migrated_in_singra:
            capa_completa_rows.append({"CAM": cam, "CAPA": capa, "RMs": ', '.join(rms_considered)})
        else:
            # se houver pendencias (incluindo migração ausente) -> capa incompleta
            capa_incompleta_rows.append({"CAM": cam, "CAPA": capa, "Pendências": '; '.join(pendencias)})

    df_capa_completa = pd.DataFrame(capa_completa_rows)
    df_capa_incompleta = pd.DataFrame(capa_incompleta_rows)
    df_migration_errors = pd.DataFrame(migration_errors)

    # Resumo
    ca, cb = st.columns(2)
    ca.success(f"CAPAs totalmente atendidas (apenas RMs sem MAPA): {len(df_capa_completa)}")
    cb.warning(f"CAPAs com pendências (considerando RMs sem MAPA): {len(df_capa_incompleta)}")

    st.subheader("✅ CAPAs completamente atendidas (somente RMs sem MAPA)")
    if not df_capa_completa.empty:
        st.dataframe(df_capa_completa.style.set_properties(**{'text-align':'left'}), use_container_width=True)
    else:
        st.info("Nenhuma CAPA completamente atendida (considerando somente RMs sem MAPA).")

    st.subheader("⚠️ CAPAs parcialmente atendidas (detalhes)")
    if not df_capa_incompleta.empty:
        st.dataframe(df_capa_incompleta.style.set_properties(**{'text-align':'left'}), use_container_width=True)
    else:
        st.info("Nenhuma CAPA parcialmente atendida encontrada (para RMs sem MAPA).")

    st.subheader("🚨 RMs do PWA que não constam no SINGRA (migração)")
    if not df_migration_errors.empty:
        st.dataframe(df_migration_errors.style.set_properties(**{'text-align':'left'}), use_container_width=True)
    else:
        st.info("Nenhuma RM do PWA ausente no SINGRA encontrada.")

# ----------------------
# (O resto dos BLOCOS 2-5 e exportação seguem iguais ao seu código original)
# ----------------------

# ----------------------
# BLOCO 2: MAPA sem STC (agrupar por CAM e MAPA) — excluir STATUS EXPEDIDO
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
# BLOCO 3: MAPA sem STC + LOTE confirmado na expedição
# ----------------------
st.markdown("## 🔷 BLOCO 3 — MAPA sem STC com LOTE confirmado na expedição (agrupar por CAM e MAPA)")

# Verificar colunas necessárias
if all(c in df_pwa.columns for c in ['MAPA','STC','STATUS','CAM','CAPA','LOTE']) and \
   'LOTE' in df_lotes_user.columns:

    # MAPA sem STC e não expedido
    df_mapa_sem_stc = df_pwa[
        (df_pwa['MAPA'] != '') &
        (df_pwa['STC'] == '') &
        (df_pwa['STATUS'] != 'EXPEDIDO')
    ]

    # Se nada encontrado, parar
    if df_mapa_sem_stc.empty:
        st.info("Nenhuma MAPA sem STC encontrada para este filtro.")
    else:

        # Padronizar LOTE para comparação
        df_mapa_sem_stc['LOTE'] = df_mapa_sem_stc['LOTE'].astype(str).str.strip()
        df_lotes_user['LOTE']   = df_lotes_user['LOTE'].astype(str).str.strip()

        # Obter lotes realmente confirmados
        lotes_validos = set(df_lotes_user['LOTE'].unique())

        # Filtrar somente linhas cujos lotes constam na planilha de LOTE
        df_mapa_com_lote_real = df_mapa_sem_stc[df_mapa_sem_stc['LOTE'].isin(lotes_validos)]

        if df_mapa_com_lote_real.empty:
            st.info("Nenhuma MAPA sem STC possui lote confirmado na expedição.")
        else:
            # Agrupar igual ao bloco 3
            agrupado_mapa5 = (
                df_mapa_com_lote_real.groupby(['CAM', 'MAPA'])
                .agg({'CAPA': lambda x: ', '.join(sorted(set(x))),
                      'LOTE': lambda x: ', '.join(sorted(set(x)))})
                .reset_index()
            )

            # Filtro por CAM
            cams5 = ["Todos"] + sorted(agrupado_mapa5['CAM'].unique().tolist())
            cam_sel5 = st.selectbox("Filtrar por CAM (Bloco 3)", cams5)

            display5 = agrupado_mapa5 if cam_sel5 == "Todos" else agrupado_mapa5[agrupado_mapa5['CAM'] == cam_sel5]

            # Exibir
            st.dataframe(
                display5.style.set_properties(**{'text-align':'left'}),
                use_container_width=True
            )

else:
    st.info("Colunas necessárias para Bloco 3 ausentes no PWA ou no arquivo de LOTE.")

# ----------------------
# BLOCO 4: STC não expedidas (agrupar por CAM e STC)
# ----------------------
st.markdown("## 🔶 BLOCO 4 — STC não expedidas (agrupar por CAM e STC)")
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
        cam_sel3 = st.selectbox("Filtrar por CAM (Bloco 4)", cams3)
        display3 = agrupado_stc if cam_sel3 == "Todos" else agrupado_stc[agrupado_stc['CAM'] == cam_sel3]
        st.dataframe(display3.style.set_properties(**{'text-align':'left'}), use_container_width=True)
else:
    st.info("Colunas necessárias para Bloco 4 ausentes no PWA.")

# ============================
# 🔷 BLOCO 5 — STC não expedidas (agrupar por CAM e STC) com LOTE confirmado na Expedição
# ============================

st.markdown("## 🔷 BLOCO 5 — STC com lote confirmado na expedição (agrupar por CAM e STC)")

# Verificar se todas as colunas necessárias existem
if all(c in df_pwa.columns for c in ['STC','STATUS','CAM','MAPA','LOTE']) and \
   'LOTE' in df_lotes_user.columns:

    # Selecionar apenas STC válidas
    df_stc_validas = df_pwa[
        (df_pwa['STC'] != '') &
        (df_pwa['STATUS'] != 'EXPEDIDO') &
        (df_pwa['STATUS'] != 'CANCELADO')
    ]

    # Garantir que o LOTE seja comparável (remover espaços e converter para string)
    df_stc_validas['LOTE'] = df_stc_validas['LOTE'].astype(str).str.strip()
    df_lotes_user['LOTE'] = df_lotes_user['LOTE'].astype(str).str.strip()

    # Filtrar apenas LOTE realmente existente na planilha LOTE (Google Sheets)
    lotes_validos = set(df_lotes_user['LOTE'].unique())
    df_stc_com_lote_real = df_stc_validas[df_stc_validas['LOTE'].isin(lotes_validos)]

    if df_stc_com_lote_real.empty:
        st.info("Nenhuma STC encontrada com lote confirmado na expedição.")
    else:
        # Agrupar (mesmo modelo do bloco 5)
        agrupado_stc4 = (
            df_stc_com_lote_real.groupby(['CAM','STC'])
            .agg({
                'MAPA': lambda x: ', '.join(sorted(set([m for m in x if m and m != '']))),
                'LOTE': lambda x: ', '.join(sorted(set(x)))
            })
            .reset_index()
        )

        # Filtro por CAM
        cams4 = ["Todos"] + sorted(agrupado_stc4['CAM'].unique().tolist())
        cam_sel4 = st.selectbox("Filtrar por CAM (Bloco 5)", cams4)

        display4 = agrupado_stc4 if cam_sel4 == "Todos" else agrupado_stc4[agrupado_stc4['CAM'] == cam_sel4]

        # Exibir tabela
        st.dataframe(
            display4.style.set_properties(**{'text-align':'left'}),
            use_container_width=True
        )

# ----------------------
# Exportação Excel (inclui debug tables)
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
        export_dfs = [
            df_capa_completa if 'df_capa_completa' in locals() else pd.DataFrame(),
            df_capa_incompleta if 'df_capa_incompleta' in locals() else pd.DataFrame(),
            # manter outputs originais do bloco 2/4 caso existam
            agrupado_mapa if 'agrupado_mapa' in locals() else pd.DataFrame(),
            agrupado_stc if 'agrupado_stc' in locals() else pd.DataFrame(),
            df_singra if 'df_singra' in locals() else pd.DataFrame(),
            df_pwa if 'df_pwa' in locals() else pd.DataFrame(),
            df_lotes_user if 'df_lotes_user' in locals() else pd.DataFrame(),
            df_migration_errors if 'df_migration_errors' in locals() else pd.DataFrame()
        ]
        names = ["CAPA_Atendidas", "CAPA_Pendentes", "MAPA_sem_STC", "STC_nao_expedida", "SINGRA_RAW", "PWA_RAW", "LOTES_CONFERENCIA", "MIGRATION_ERRORS"]
        excel_bytes = to_excel(export_dfs, names)
        st.download_button(
            label="📥 Baixar Excel completo",
            data=excel_bytes,
            file_name="resultado_controle_rm_completo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ============================================================
# 🚀 NOVO MÓDULO — ANÁLISE DE LOTE E CAPA COMPLETAMENTE ATENDIDOS
# ============================================================

st.markdown("---")
st.header("📦 Análise de Lotes e Capas Completamente Atendidos")

# 1. Conjunto de VOLUMES que estão fisicamente na expedição (planilha LOTE)
volumes_exp = set(df_lotes_user["LOTE"].astype(str).tolist())

# 2. Garantir tipagem correta no PWA
df_pwa["VOLUME"] = df_pwa["VOLUME"].astype(str).str.strip()
df_pwa["LOTE"] = df_pwa["LOTE"].astype(str).str.strip()
df_pwa["CAPA"] = df_pwa["CAPA"].astype(str).str.strip()

# 3. Agrupamentos
# LOTES → lista de volumes de cada lote
lote_to_volumes = {
    lote: set(grupo["VOLUME"].tolist())
    for lote, grupo in df_pwa.groupby("LOTE")
}

# CAPA → LOTES associados
capa_to_lotes = {
    capa: set(grupo["LOTE"].unique().tolist())
    for capa, grupo in df_pwa.groupby("CAPA")
}

# ============================================================
# 4. Identificar LOTES completamente atendidos
# ============================================================

lotes_completos = []
lotes_incompletos = []

for lote, volumes_lote in lote_to_volumes.items():

    # volumes faltantes = volumes do lote que não estão na planilha LOTE
    volumes_faltando = volumes_lote - volumes_exp

    if len(volumes_faltando) == 0:
        lotes_completos.append({
            "LOTE": lote,
            "TOTAL VOLUMES": len(volumes_lote),
            "STATUS": "COMPLETO"
        })
    else:
        lotes_incompletos.append({
            "LOTE": lote,
            "TOTAL VOLUMES": len(volumes_lote),
            "VOLUMES FALTANTES": ", ".join(sorted(volumes_faltando)),
            "STATUS": "INCOMPLETO"
        })

df_lotes_completos = pd.DataFrame(lotes_completos)
df_lotes_incompletos = pd.DataFrame(lotes_incompletos)

# ============================================================
# 5. Identificar CAPAS completamente atendidas
# ============================================================

capas_completas = []
capas_incompletas = []

# transforma lotes completos em set para performance
lotes_completos_set = set(df_lotes_completos["LOTE"].tolist()) if not df_lotes_completos.empty else set()

for capa, lotes_da_capa in capa_to_lotes.items():

    # Se todos os LOTES dessa CAPA estão completos → CAPA completa
    if lotes_da_capa.issubset(lotes_completos_set):
        capas_completas.append({"CAPA": capa, "TOTAL LOTES": len(lotes_da_capa), "STATUS": "COMPLETA"})
    else:
        lotes_faltantes = lotes_da_capa - lotes_completos_set
        capas_incompletas.append({
            "CAPA": capa,
            "TOTAL LOTES": len(lotes_da_capa),
            "LOTES NÃO ATENDIDOS": ", ".join(sorted(lotes_faltantes)),
            "STATUS": "INCOMPLETA"
        })

df_capas_completas = pd.DataFrame(capas_completas)
df_capas_incompletas = pd.DataFrame(capas_incompletas)

# ============================================================
# 6. EXIBIÇÃO
# ============================================================

st.subheader("✅ LOTES Completamente Atendidos")
if df_lotes_completos.empty:
    st.info("Nenhum LOTE completamente atendido ainda.")
else:
    st.dataframe(df_lotes_completos, use_container_width=True)

st.subheader("⚠️ LOTES Incompletos")
if df_lotes_incompletos.empty:
    st.success("Todos os LOTES estão completos!")
else:
    st.dataframe(df_lotes_incompletos, use_container_width=True)

st.subheader("🏁 CAPAS Completamente Atendidas")
if df_capas_completas.empty:
    st.info("Nenhuma CAPA completamente atendida ainda.")
else:
    st.dataframe(df_capas_completas, use_container_width=True)

st.subheader("📍 CAPAS Incompletas")
if df_capas_incompletas.empty:
    st.success("Todas as CAPAS estão completas!")
else:
    st.dataframe(df_capas_incompletas, use_container_width=True)



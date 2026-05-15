import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re

st.set_page_config(page_title="Controle de RM atendidas", layout="wide")
st.title("📦 Controle de RMs - Estocagem e Expedição")
st.markdown("Sistema: PWA = fonte da verdade. Bloco 1 com validação rigorosa de CAPAS prontas, parciais e pendentes.")

# ----------------------
# Utilitários / Normalização
# ----------------------
def clean_colnames(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).replace('\ufeff', '').replace("'", "").replace('"', '').strip().upper() for c in df.columns]
    return df

def normalizar_codigo_rm(valor):
    if pd.isna(valor) or str(valor).strip() == '':
        return ''
    s = str(valor).replace('\ufeff', '').strip().replace("'", "").replace('"', "")
    s = s.replace(".", "").replace(",", "").replace(" ", "")
    if s.endswith('.0'):
        s = s[:-2]
    return s

def normalizar_lote(valor):
    if pd.isna(valor):
        return ''
    v_str = str(valor).replace('\ufeff', '').strip().replace("'", "").replace('"', '')
    if v_str.endswith('.0'):
        v_str = v_str[:-2]
    return v_str

# ----------------------
# Cache: carregamento de dados
# ----------------------
@st.cache_data
def carregar_singra(file):
    try:
        # 1ª Tentativa: Lê com utf-8-sig
        df = pd.read_csv(file, sep=';', encoding='utf-8-sig', dtype=str, low_memory=False)
    except Exception:
        # REBOBINA o arquivo para a posição 0 antes de tentar de novo
        file.seek(0)
        # 2ª Tentativa: Lê com latin1
        df = pd.read_csv(file, sep=';', encoding='latin1', dtype=str, low_memory=False)
    
    df = clean_colnames(df)
    df = df.fillna('')
    
    # Busca inteligente da coluna ID caso venha com sujeira
    if 'ID' not in df.columns:
        for col in df.columns:
            if 'ID' in col:
                df.rename(columns={col: 'ID'}, inplace=True)
                break
                
    if 'ID' in df.columns:
        df['ID'] = df['ID'].apply(normalizar_codigo_rm)
    return df

@st.cache_data
def carregar_pwa(file):
    df = pd.read_excel(file, sheet_name=0, dtype=str)
    df = clean_colnames(df)
    df = df.fillna('')
    
    cols_to_clean = ['PEDIDO', 'CAPA', 'MAPA', 'STC', 'CAM', 'LOTE', 'STATUS']
    for col in cols_to_clean:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            
    if 'PEDIDO' in df.columns:
        df['PEDIDO_LIMPO'] = df['PEDIDO'].apply(normalizar_codigo_rm)
    else:
        df['PEDIDO_LIMPO'] = ''
        
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
# UI: Uploads
# ----------------------
with st.expander("📄 Upload de arquivos", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        singra_file = st.file_uploader("Upload planilha do SINGRA (.csv)", type=["csv"])
    with col2:
        pwa_file = st.file_uploader("Upload planilha do PWA (.xlsx)", type=["xlsx"])

if not (singra_file and pwa_file):
    st.info("Faça upload do SINGRA (.csv) e do PWA (.xlsx) para prosseguir.")
    st.stop()

# Carregamento
df_singra = carregar_singra(singra_file)
df_pwa = carregar_pwa(pwa_file)

# Carregar Lotes (Google Sheets)
try:
    SHEET_URL = "https://docs.google.com/spreadsheets/d/1naVnAlUGmeAMb_YftLGYit-1e1BcYFJgiJwSnOcgJf4/edit?gid=0"
    service_account_dict = dict(st.secrets["gcp_service_account"])
    df_lotes_user = carregar_lotes_google(service_account_dict, SHEET_URL)
except Exception as e:
    st.error(f"Erro ao conectar com o Google Sheets: {e}")
    st.stop()

# ----------------------
# Preparação dos Conjuntos (Sets) para Validação Rápida
# ----------------------
lotes_disponiveis = set(df_lotes_user['LOTE'].apply(normalizar_lote)) if 'LOTE' in df_lotes_user.columns else set()
# Remove lotes vazios do set para não dar falso positivo
lotes_disponiveis.discard('') 

pedidos_singra = set()
if 'ID' in df_singra.columns:
    pedidos_singra = set(df_singra['ID'].dropna().tolist())
pedidos_singra.discard('')

c1, c2, c3 = st.columns(3)
c1.metric("RMs únicas (PWA)", df_pwa['PEDIDO_LIMPO'].nunique())
c2.metric("RMs no SINGRA", len(pedidos_singra))
c3.metric("Lotes conferidos (Google)", len(lotes_disponiveis))

st.divider()

# ----------------------
# BLOCO 1 – VISÃO POR CAPA E VISÃO POR RM (COM MÉTRICAS DE RESUMO)
# ----------------------
st.markdown("## 🔵 BLOCO 1 — Status de Processamento e Expedição")

required_pwa_cols = ['PEDIDO_LIMPO', 'LOTE', 'CAPA', 'CAM', 'STATUS', 'MAPA']
if not all(c in df_pwa.columns for c in required_pwa_cols):
    st.error(f"Colunas essenciais faltando no PWA. Necessário: {required_pwa_cols}")
else:
    # --- PROCESSAMENTO DOS DADOS ---
    lista_rm_final = []
    capas_prontas = []
    capas_parciais = []     
    capas_pendentes = []    
    capas_quebradas_prontas = []   
    capas_quebradas_pendentes = [] 
    capas_finalizadas = []  

    # 1. Processamento por RM (Individual)
    for rm, grupo_rm in df_pwa.groupby('PEDIDO_LIMPO'):
        if rm == '': continue
        
        cam_rm = str(grupo_rm['CAM'].iloc[0])
        capa_rm = str(grupo_rm['CAPA'].iloc[0])
        status_pwa = str(grupo_rm['STATUS'].iloc[0]).upper()
        
        tem_mapa_rm = grupo_rm['MAPA'].replace(r'^\s*$', np.nan, regex=True).notna().any()
        mapa_val = ", ".join(set(grupo_rm['MAPA'].dropna().astype(str).str.strip())) if tem_mapa_rm else ""
        
        if status_pwa == 'CANCELADO':
            categoria = "CANCELADA"
            pendencia = "Item cancelado no sistema"
        elif tem_mapa_rm:
            categoria = "COM MAPA"
            pendencia = f"MAPA gerado: {mapa_val}"
        else:
            lotes_rm = set(grupo_rm['LOTE'].apply(normalizar_lote)) - {''}
            lotes_faltantes = lotes_rm - lotes_disponiveis
            no_singra = rm in pedidos_singra
            
            if not lotes_faltantes and no_singra:
                categoria = "PRONTA"
                pendencia = "Apta para gerar MAPA (Em Expedição)"
            else:
                categoria = "PENDENTE"
                erros = []
                if lotes_faltantes: erros.append(f"Lotes não bipados na exp.: {', '.join(sorted(lotes_faltantes))}")
                if not no_singra: erros.append("Não consta 'Em Expedição' no SINGRA")
                pendencia = " | ".join(erros)
        
        lista_rm_final.append({
            "RM": rm, "CAPA": capa_rm, "CAM": cam_rm, 
            "STATUS PWA": status_pwa, "SITUAÇÃO": categoria, "DETALHE": pendencia
        })

    df_rm_visao = pd.DataFrame(lista_rm_final)

    # --- CÁLCULO DAS MÉTRICAS DE RESUMO ---
    # Contamos apenas as RMs que não estão canceladas nem já possuem mapa
    total_prontas = len(df_rm_visao[df_rm_visao['SITUAÇÃO'] == "PRONTA"])
    total_pendentes = len(df_rm_visao[df_rm_visao['SITUAÇÃO'] == "PENDENTE"])
    total_com_mapa = len(df_rm_visao[df_rm_visao['SITUAÇÃO'] == "COM MAPA"])

    # Exibição das métricas em destaque
    m1, m2, m3 = st.columns(3)
    m1.metric("✅ RMs Prontas p/ MAPA", total_prontas)
    m2.metric("⚠️ RMs Pendentes", total_pendentes)
    m3.metric("🏁 RMs com MAPA (Finalizadas)", total_com_mapa)
    st.divider()

    # 2. Processamento por CAPA (Agrupado)
    for capa, grupo_capa in df_pwa.groupby('CAPA'):
        if capa == '': continue
        cam_capa = str(grupo_capa['CAM'].iloc[0])
        
        mascara_cancelado = grupo_capa['STATUS'].astype(str).str.upper() == 'CANCELADO'
        tem_cancelado = mascara_cancelado.any()
        grupo_ativo = grupo_capa[~mascara_cancelado]
        
        if grupo_ativo.empty: continue 
        
        mascara_com_mapa = grupo_ativo['MAPA'].replace(r'^\s*$', np.nan, regex=True).notna()
        qtd_com_mapa = mascara_com_mapa.sum()
        total_ativos = len(grupo_ativo)
        
        pedidos_ativos = set(grupo_ativo['PEDIDO_LIMPO'].apply(normalizar_codigo_rm)) - {''}
        mapas_existentes = set(grupo_ativo['MAPA'].dropna().astype(str).str.strip()) - {''}
        
        if qtd_com_mapa == total_ativos:
            capas_finalizadas.append({
                "CAPA": capa, "CAM": cam_capa, 
                "RMs": ", ".join(sorted(pedidos_ativos)), 
                "MAPAs": ", ".join(sorted(mapas_existentes))
            })
        elif 0 < qtd_com_mapa < total_ativos:
            rms_com = set(grupo_ativo[mascara_com_mapa]['PEDIDO_LIMPO'].apply(normalizar_codigo_rm)) - {''}
            rms_sem = pedidos_ativos - rms_com
            detalhe_geral = f"MAPAs existentes: {', '.join(sorted(mapas_existentes))}\n"
            detalhe_geral += f"RMs já com MAPA: {', '.join(sorted(rms_com))}\n"
            
            grupo_restante = grupo_ativo[grupo_ativo['PEDIDO_LIMPO'].isin(rms_sem)]
            lotes_restantes = set(grupo_restante['LOTE'].apply(normalizar_lote)) - {''}
            faltantes_lote_rest = lotes_restantes - lotes_disponiveis
            faltantes_singra_rest = rms_sem - pedidos_singra
            
            if not faltantes_lote_rest and not faltantes_singra_rest:
                capas_quebradas_prontas.append({
                    "CAPA": capa, "CAM": cam_capa,
                    "Qtd RM": len(rms_sem),
                    "RMs Pendentes (Prontas)": ", ".join(sorted(rms_sem)),
                    "Histórico": detalhe_geral
                })
            else:
                razão_quebra = [detalhe_geral]
                if faltantes_lote_rest: razão_quebra.append(f"Lotes Restantes ausentes: {', '.join(sorted(faltantes_lote_rest))}")
                if faltantes_singra_rest:
                    status_dict_rest = {}
                    for r in faltantes_singra_rest:
                        st_wms = str(grupo_restante[grupo_restante['PEDIDO_LIMPO'] == r]['STATUS'].iloc[0]).upper()
                        status_dict_rest.setdefault(st_wms, []).append(r)
                    msg_s = "RMs Restantes fora Singra:\n" + "\n".join([f"- {s}: {', '.join(sorted(rs))}" for s, rs in status_dict_rest.items()])
                    razão_quebra.append(msg_s)

                capas_quebradas_pendentes.append({
                    "CAPA": capa, "CAM": cam_capa,
                    "Qtd RM": len(rms_sem),
                    "RMs s/ MAPA": ", ".join(sorted(rms_sem)),
                    "Pendência do Restante": "\n\n".join(razão_quebra)
                })
        else:
            lotes_ativos = set(grupo_ativo['LOTE'].apply(normalizar_lote)) - {''}
            faltantes_lote = lotes_ativos - lotes_disponiveis
            faltantes_singra = pedidos_ativos - pedidos_singra

            if not faltantes_lote and not faltantes_singra:
                if tem_cancelado:
                    capas_parciais.append({"CAPA": capa, "CAM": cam_capa, "RMs Ativas": ", ".join(sorted(pedidos_ativos))})
                else:
                    capas_prontas.append({"CAPA": capa, "CAM": cam_capa, "Qtd RM": len(pedidos_ativos), "RMs (100% Prontas)": ", ".join(sorted(pedidos_ativos))})
            else:
                razão = []
                if faltantes_lote: razão.append(f"Lotes que não estão na Expedição: {', '.join(sorted(faltantes_lote))}")
                if faltantes_singra: 
                    status_dict = {}
                    for rm_f in faltantes_singra:
                        st_wms = str(grupo_ativo[grupo_ativo['PEDIDO_LIMPO'] == rm_f]['STATUS'].iloc[0]).upper()
                        status_dict.setdefault(st_wms or "SEM STATUS", []).append(rm_f)
                    texto_s = "RMs fora Singra:\n" + "\n".join([f"- {s}: {', '.join(sorted(rs))}" for s, rs in sorted(status_dict.items())])
                    razão.append(texto_s)
                
                capas_pendentes.append({
                    "CAPA": capa, "CAM": cam_capa, 
                    "Qtd RM": len(pedidos_ativos),
                    "RMs da CAPA": ", ".join(sorted(pedidos_ativos)), 
                    "O que falta?": "\n\n".join(razão)
                })

    # --- INTERFACE ---
    aba_capa, aba_rm = st.tabs(["📋 Visão por CAPA", "📄 Visão por RM (Individual)"])

    with aba_capa:
        t1, t2, t3, t4, t5, t6 = st.tabs([
            f"✅ Prontas ({len(capas_prontas)})", 
            f"🧩 Quebradas Prontas ({len(capas_quebradas_prontas)})", 
            f"⚠️ Pendentes ({len(capas_pendentes)})", 
            f"🧩 Quebradas Pendentes ({len(capas_quebradas_pendentes)})", 
            f"🏁 Finalizadas ({len(capas_finalizadas)})", 
            f"🔶 C/ Cancelamento ({len(capas_parciais)})"
        ])
        
        with t1: 
            st.dataframe(pd.DataFrame(capas_prontas), use_container_width=True)
        with t2:
            st.dataframe(pd.DataFrame(capas_quebradas_prontas).style.set_properties(**{'white-space': 'pre-wrap'}), use_container_width=True)
        with t3: 
            st.dataframe(pd.DataFrame(capas_pendentes).style.set_properties(**{'white-space': 'pre-wrap'}), use_container_width=True)
        with t4:
            st.dataframe(pd.DataFrame(capas_quebradas_pendentes).style.set_properties(**{'white-space': 'pre-wrap'}), use_container_width=True)
        with t5: 
            st.dataframe(pd.DataFrame(capas_finalizadas), use_container_width=True)
        with t6: 
            st.dataframe(pd.DataFrame(capas_parciais), use_container_width=True)

    with aba_rm:
        st.subheader("Rastreio Individual de RMs")
        # Filtros e Tabela de RM permanecem iguais...
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            cam_list = ["TODOS"] + sorted(df_rm_visao['CAM'].unique().tolist())
            filtro_cam = st.selectbox("Filtrar por CAM", cam_list)
        with col_f2:
            sit_list = ["TODAS", "PRONTA", "PENDENTE", "COM MAPA", "CANCELADA"]
            filtro_sit = st.selectbox("Filtrar por Situação", sit_list)

        df_filtrado = df_rm_visao.copy()
        if filtro_cam != "TODOS": df_filtrado = df_filtrado[df_filtrado['CAM'] == filtro_cam]
        if filtro_sit != "TODAS": df_filtrado = df_filtrado[df_filtrado['SITUAÇÃO'] == filtro_sit]

        st.write(f"Exibindo {len(df_filtrado)} RMs")
        st.dataframe(df_filtrado, use_container_width=True, hide_index=True)

st.divider()

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


st.divider()

# ----------------------
# BLOCO 3: STC não expedidas (agrupar por CAM e STC)
# ----------------------
st.markdown("## 🔷 BLOCO 3 — STC não expedidas (agrupar por CAM e STC)")

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
        cam_sel3 = st.selectbox("Filtrar por CAM (BLOCO 3)", cams3)
        display3 = agrupado_stc if cam_sel3 == "Todos" else agrupado_stc[agrupado_stc['CAM'] == cam_sel3]
        st.dataframe(display3.style.set_properties(**{'text-align':'left'}), use_container_width=True)
else:
    st.info("Colunas necessárias para BLOCO 3 ausentes no PWA.")

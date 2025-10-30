import sys, os, platform, traceback, re
from pathlib import Path
import pandas as pd
import streamlit as st
from csv import Sniffer

st.set_page_config(page_title="Resumo de Entregadores", layout="wide")

PT_BR_MONTHS = {"jan":"01","fev":"02","mar":"03","abr":"04","mai":"05","jun":"06",
                "jul":"07","ago":"08","set":"09","out":"10","nov":"11","dez":"12"}

def _parse_brl_to_float(s):
    if pd.isna(s): return pd.NA
    s = str(s).strip()
    if not s: return pd.NA
    s = s.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
    try: return float(s)
    except Exception: return pd.NA

def _normalize_ptbr_datetime(s):
    if pd.isna(s): return pd.NaT
    s = str(s)
    s = re.sub(r'(\d{1,2})h(\d{2})', r'\1:\2', s)
    for k,v in PT_BR_MONTHS.items():
        s = re.sub(fr'/{k}/', f'/{v}/', s, flags=re.IGNORECASE)
    return pd.to_datetime(s, dayfirst=True, errors="coerce")

def _best_read_csv(path: Path):
    encodings = ["utf-8-sig", "utf-8", "latin-1"]
    seps = [",",";","\t","|"]
    # 1) Sniffer
    for enc in encodings:
        try:
            sample = path.read_text(encoding=enc, errors="ignore")[:8192]
            dialect = Sniffer().sniff(sample, delimiters=";,|\t")
            df = pd.read_csv(path, encoding=enc, sep=dialect.delimiter)
            if df.shape[1] > 1: return df
        except Exception:
            pass
    # 2) Heurística: maior nº de colunas
    best, best_cols = None, 0
    for enc in encodings:
        for sep in seps:
            try:
                df = pd.read_csv(path, encoding=enc, sep=sep)
                if df.shape[1] > best_cols:
                    best, best_cols = df, df.shape[1]
            except Exception:
                continue
    if best is not None: return best
    # 3) Última tentativa
    return pd.read_csv(path)

def carregar_dados_failproof():
    app_dir = Path(__file__).parent.resolve()
    candidatos = [app_dir / "base", Path.cwd() / "base"]
    base_path = next((p for p in candidatos if p.exists()), None)
    if base_path is None:
        raise FileNotFoundError("Pasta 'base' não encontrada. Procurei em: " + ", ".join(map(str, candidatos)))
    arquivos = sorted(base_path.glob("resumo_entregadores_FINAL*.csv"))
    if not arquivos:
        raise FileNotFoundError(f"Nenhum arquivo 'resumo_entregadores_FINAL*.csv' em {base_path}")

    dfs = []
    for arq in arquivos:
        df = _best_read_csv(arq)
        df["__arquivo_origem"] = arq.name
        dfs.append(df)

    full = pd.concat(dfs, ignore_index=True)
    full.columns = (full.columns.astype(str).str.strip().str.replace(r"\s+","_",regex=True).str.lower())

    aliases = {
        "entregador": ["entregador","nome_entregador"],
        "chave_pix": ["chave_pix","pix","email_pix"],
        "data_filtro": ["data_filtro","data","data_base"],
        "hora_inicio_filtro": ["hora_início_filtro","hora_inicio_filtro","hora_inicio"],
        "hora_fim_filtro": ["hora_fim_filtro","hora_fim"],
        "loja": ["loja","unidade","restaurante"],
        "status": ["status"],
        "valor_do_pedido": ["valor_do_pedido","valor_pedido","pedido_valor"],
        "taxa_de_entrega": ["taxa_de_entrega","taxa_entrega"],
        "taxa_do_entregador": ["taxa_do_entregador","taxa_entregador"],
        "data_de_criacao": ["data_de_criação","data_de_criacao","criado_em"],
        "pagamento": ["pagamento","forma_pagamento"],
        "classificacao_do_turno": ["classificação_do_turno","classificacao_do_turno","turno"],
        "classificacao_do_dia": ["classificação_do_dia","classificacao_do_dia","dia_semana"],
        "pagamento_por_turno": ["pagamento_por_turno"],
    }
    cols = set(full.columns)
    canon_map = {}
    for canon, opts in aliases.items():
        for opt in opts:
            if opt in cols:
                canon_map[canon] = opt
                break

    if "data_filtro" in canon_map:
        full[canon_map["data_filtro"]] = pd.to_datetime(full[canon_map["data_filtro"]], errors="coerce", dayfirst=True)
    if "data_de_criacao" in canon_map:
        full[canon_map["data_de_criacao"]] = full[canon_map["data_de_criacao"]].apply(_normalize_ptbr_datetime)

    for money in ["valor_do_pedido","taxa_de_entrega","taxa_do_entregador","pagamento_por_turno"]:
        if money in canon_map:
            col = canon_map[money]
            full[col] = full[col].apply(_parse_brl_to_float)

    return full, [p.name for p in arquivos], str(base_path)

# ===================== UI =====================
st.title("📦 Resumo de Entregadores — Diagnóstico")

with st.sidebar:
    st.subheader("Diagnóstico rápido")
    st.write("**Arquivo**:", Path(__file__).resolve())
    st.write("**CWD**:", Path.cwd())
    st.write("**Python**:", sys.version.split()[0], platform.platform())
    app_dir = Path(__file__).parent.resolve()
    base_local = app_dir / "base"
    st.write("**auto/base existe?**", base_local.exists())
    if base_local.exists():
        st.write("**Conteúdo de auto/base**:", [p.name for p in sorted(base_local.iterdir())])

    raiz_base = Path.cwd() / "base"
    st.write("**raiz/base existe?**", raiz_base.exists())
    if raiz_base.exists():
        st.write("**Conteúdo de raiz/base**:", [p.name for p in sorted(raiz_base.iterdir())])

# Tenta carregar dados e sempre mostra algo na página
try:
    df, arquivos, caminho_base = carregar_dados_failproof()
    st.success(f"✅ Dados carregados de: `{caminho_base}`")
    colA, colB = st.columns(2)
    with colA: st.metric("Arquivos lidos", len(arquivos))
    with colB: st.metric("Linhas (total)", len(df))
    with st.expander("Arquivos processados"):
        st.write(arquivos)

    # Coluna de valor preferida
    prefer = ["taxa_do_entregador","valor_do_pedido","taxa_de_entrega","pagamento_por_turno"]
    col_valor = next((c for c in prefer if c in df.columns), None)
    st.write("**Preview (5 linhas):**")
    st.dataframe(df.head(5), use_container_width=True)

    st.subheader("Tabela consolidada")
    st.dataframe(df, use_container_width=True)

    # KPI total se achar coluna numérica
    if col_valor is not None:
        vals = pd.to_numeric(df[col_valor], errors="coerce")
        total_valor = vals.sum(skipna=True)
        st.metric(f"Soma de `{col_valor}`", f"{total_valor:,.2f}".replace(",", "X").replace(".", ",").replace("X","."))

    # Gráfico (opcional)
    try:
        import plotly.express as px
        possiveis_entregadores = [c for c in df.columns if any(x in c for x in ["entregador","motoboy","courier"])]
        if possiveis_entregadores and col_valor:
            tmp = df.copy()
            tmp[col_valor] = pd.to_numeric(tmp[col_valor], errors="coerce")
            col_ent = possiveis_entregadores[0]
            g = tmp.groupby(col_ent, dropna=True, as_index=False)[col_valor].sum().sort_values(col_valor, ascending=False).head(25)
            st.subheader(f"Top 25 por `{col_valor}`")
            fig = px.bar(g, x=col_ent, y=col_valor)
            st.plotly_chart(fig, use_container_width=True)
    except Exception as e:
        st.info(f"Gráfico desativado: {e}")

except Exception as e:
    st.error("❌ Ocorreu um erro ao carregar dados.")
    st.exception(e)
    st.stop()

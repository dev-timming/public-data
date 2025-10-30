# app.py
# Painel Streamlit — Resumo de Entregadores
# Estrutura esperada no repositório:
#   auto/
#     ├─ app.py  (este arquivo)
#     └─ base/   (CSV(s): resumo_entregadores_FINAL*.csv)

import os
import re
import sys
import platform
from pathlib import Path
from csv import Sniffer

import pandas as pd
import streamlit as st

# Opcional, mas recomendado no requirements.txt:
# plotly>=5.22
try:
    import plotly.express as px
    HAS_PLOTLY = True
    PLOTLY_TEMPLATE = "plotly_white"
except Exception:
    HAS_PLOTLY = False
    PLOTLY_TEMPLATE = None

# -----------------------------------------------------------------------------
# CONFIG GERAL + CSS LEVE
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Painel de Entregadores",
    page_icon="🚚",
    layout="wide",
)

# CSS sutil para “turbinar” o visual sem quebrar com updates do Streamlit
st.markdown("""
<style>
/* Títulos mais encorpados */
h1, h2, h3 { font-weight: 700 !important; }

/* Cards das métricas com sombra suave */
section[data-testid="stMetric"] {
  background: var(--secondary-background-color);
  padding: 12px; border-radius: 12px;
  box-shadow: 0 2px 8px rgba(0,0,0,0.05);
}

/* Tabela com borda leve */
[data-testid="stDataFrame"] {
  border-radius: 12px;
  border: 1px solid rgba(0,0,0,0.06);
}

/* Containers com leve respiro */
.block-container { padding-top: 1.0rem; padding-bottom: 2rem; }
</style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# ALIASES & HELPERS
# -----------------------------------------------------------------------------
PT_BR_MONTHS = {
    "jan": "01","fev": "02","mar": "03","abr": "04","mai": "05","jun": "06",
    "jul": "07","ago": "08","set": "09","out": "10","nov": "11","dez": "12",
}

ALIASES = {
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

def get_col(df: pd.DataFrame, canon_key: str):
    opts = ALIASES.get(canon_key, [])
    for opt in opts:
        if opt in df.columns:
            return opt
    return None

def parse_brl_to_float(s):
    """Converte 'R$ 1.234,56' -> 1234.56; vazio -> NaN."""
    if pd.isna(s):
        return pd.NA
    s = str(s).strip()
    if not s:
        return pd.NA
    s = s.replace("R$", "").replace(" ", "")
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return pd.NA

def normalize_ptbr_datetime(s):
    """Converte '29/out/2025 22h42' -> to_datetime(...)."""
    if pd.isna(s):
        return pd.NaT
    s = str(s)
    s = re.sub(r'(\d{1,2})h(\d{2})', r'\1:\2', s)  # '22h42' -> '22:42'
    for k, v in PT_BR_MONTHS.items():
        s = re.sub(fr'/{k}/', f'/{v}/', s, flags=re.IGNORECASE)
    return pd.to_datetime(s, dayfirst=True, errors="coerce")

def best_read_csv(path: Path) -> pd.DataFrame:
    """Detecta separador/encoding (csv.Sniffer + heurística)."""
    encodings = ["utf-8-sig", "utf-8", "latin-1"]
    seps = [",",";","\t","|"]

    # 1) Sniffer
    for enc in encodings:
        try:
            sample = path.read_text(encoding=enc, errors="ignore")[:8192]
            dialect = Sniffer().sniff(sample, delimiters=";,|\t")
            df = pd.read_csv(path, encoding=enc, sep=dialect.delimiter)
            if df.shape[1] > 1:
                return df
        except Exception:
            pass

    # 2) Heurística: maior nº de colunas
    best_df, best_cols = None, 0
    for enc in encodings:
        for sep in seps:
            try:
                df = pd.read_csv(path, encoding=enc, sep=sep)
                if df.shape[1] > best_cols:
                    best_df, best_cols = df, df.shape[1]
            except Exception:
                continue
    if best_df is not None:
        return best_df

    # 3) Último recurso: default do pandas
    return pd.read_csv(path)

def load_data():
    """Carrega e limpa os CSVs da pasta 'base' (ao lado de app.py ou na raiz)."""
    app_dir = Path(__file__).parent.resolve()
    candidatos = [app_dir / "base", Path.cwd() / "base"]
    base_path = next((p for p in candidatos if p.exists()), None)
    if base_path is None:
        raise FileNotFoundError("Pasta 'base' não encontrada. Procurei em: "
                                + ", ".join(map(str, candidatos)))
    arquivos = sorted(base_path.glob("resumo_entregadores_FINAL*.csv"))
    if not arquivos:
        raise FileNotFoundError(f"Nenhum arquivo 'resumo_entregadores_FINAL*.csv' em {base_path}")

    dfs = []
    for arq in arquivos:
        df = best_read_csv(arq)
        df["__arquivo_origem"] = arq.name
        dfs.append(df)
    full = pd.concat(dfs, ignore_index=True)

    # normaliza nomes
    full.columns = (full.columns.astype(str)
                    .str.strip()
                    .str.replace(r"\s+","_",regex=True)
                    .str.lower())

    # datas
    c_data = get_col(full, "data_filtro")
    c_criacao = get_col(full, "data_de_criacao")
    if c_data:
        if not pd.api.types.is_datetime64_any_dtype(full[c_data]):
            full[c_data] = pd.to_datetime(full[c_data], errors="coerce", dayfirst=True)
    if c_criacao:
        full[c_criacao] = full[c_criacao].apply(normalize_ptbr_datetime)

    # valores BRL
    for money in ["valor_do_pedido","taxa_de_entrega","taxa_do_entregador","pagamento_por_turno"]:
        cm = get_col(full, money)
        if cm:
            full[cm] = full[cm].apply(parse_brl_to_float)

    return full, [p.name for p in arquivos], str(base_path)

# -----------------------------------------------------------------------------
# UI — TITULO
# -----------------------------------------------------------------------------
st.title("📦 Painel — Resumo de Entregadores")

# Sidebar — modo diagnóstico + filtros
with st.sidebar:
    st.header("⚙️ Opções")
    show_diag = st.checkbox("Mostrar diagnóstico", value=False)

# -----------------------------------------------------------------------------
# CARREGAMENTO COM TRATAMENTO DE ERROS
# -----------------------------------------------------------------------------
try:
    df, arquivos, caminho_base = load_data()

    if show_diag:
        st.sidebar.subheader("🩺 Diagnóstico")
        st.sidebar.write("**Arquivo**:", Path(__file__).resolve())
        st.sidebar.write("**CWD**:", Path.cwd())
        st.sidebar.write("**Python**:", sys.version.split()[0], platform.platform())
        st.sidebar.write("**Pasta usada**:", caminho_base)
        st.sidebar.write("**Arquivos**:", arquivos)

    # KPIs
    colA, colB, colC = st.columns(3)
    with colA: st.metric("Arquivos lidos", len(arquivos))
    with colB: st.metric("Linhas (total)", len(df))

    # KPI total taxa do entregador (se existir)
    c_taxa_ent = get_col(df, "taxa_do_entregador")
    if c_taxa_ent is not None:
        total_taxa = pd.to_numeric(df[c_taxa_ent], errors="coerce").sum(skipna=True)
        with colC:
            st.metric("Total Taxa do Entregador (R$)",
                      f"{total_taxa:,.2f}".replace(",", "X").replace(".", ",").replace("X","."))
    else:
        with colC:
            st.metric("Total Taxa do Entregador (R$)", "—")

    # Filtros (sidebar)
    with st.sidebar:
        st.header("Filtros")
        df_filtrado = df.copy()

        # filtro por entregador (se existir)
        c_entregador = get_col(df, "entregador")
        if c_entregador:
            entregadores = sorted(df[c_entregador].dropna().astype(str).unique().tolist())[:5000]
            sel_ent = st.multiselect("Entregador", entregadores)
            if sel_ent:
                df_filtrado = df_filtrado[df_filtrado[c_entregador].astype(str).isin(sel_ent)]

        # filtro por período (data_filtro)
        c_data = get_col(df, "data_filtro")
        if c_data and pd.api.types.is_datetime64_any_dtype(df_filtrado[c_data]):
            min_d = pd.to_datetime(df_filtrado[c_data]).min()
            max_d = pd.to_datetime(df_filtrado[c_data]).max()
            if pd.notna(min_d) and pd.notna(max_d):
                periodo = st.date_input("Período (data_filtro)", value=(min_d.date(), max_d.date()))
                if isinstance(periodo, tuple) and len(periodo) == 2:
                    ini, fim = periodo
                    mask = (df_filtrado[c_data] >= pd.to_datetime(ini)) & (df_filtrado[c_data] <= pd.to_datetime(fim))
                    df_filtrado = df_filtrado[mask]

    st.divider()
    st.markdown("### 🔎 Tabela consolidada (após filtros)")
    st.dataframe(df_filtrado, use_container_width=True)

    # -----------------------------------------------------------------------------
    # VISÕES
    # -----------------------------------------------------------------------------
    st.divider()
    st.markdown("## 📊 Visões")

    aba1, aba2 = st.tabs([
        "Entregas por dia (data_filtro)",
        "Total de taxa do entregador por turno",
    ])

    # ---- ABA 1: COUNT por dia (data_filtro)
    with aba1:
        c_data = get_col(df_filtrado, "data_filtro")
        if not c_data:
            st.info("Não encontrei a coluna de data (ex.: `Data Filtro`).")
        else:
            if not pd.api.types.is_datetime64_any_dtype(df_filtrado[c_data]):
                df_filtrado[c_data] = pd.to_datetime(df_filtrado[c_data], errors="coerce", dayfirst=True)

            aux = df_filtrado.copy()
            aux["dia"] = aux[c_data].dt.date
            grp = aux.groupby("dia", dropna=True).size().reset_index(name="total_entregas")

            if grp.empty:
                st.warning("Sem linhas com `data_filtro` válida para agrupar.")
            else:
                if HAS_PLOTLY:
                    fig1 = px.bar(
                        grp.sort_values("dia"),
                        x="dia",
                        y="total_entregas",
                        title="Total de entregas por dia (data_filtro)",
                        labels={"dia": "Dia", "total_entregas": "Total de entregas"},
                        template=PLOTLY_TEMPLATE,
                    )
                    fig1.update_yaxes(tickformat="d")
                    st.plotly_chart(fig1, use_container_width=True)
                else:
                    st.info("Plotly não está disponível. Adicione `plotly>=5.22` ao requirements.txt.")
                st.dataframe(grp.sort_values("dia"), use_container_width=True)

    # ---- ABA 2: Soma Taxa do Entregador por Turno
    with aba2:
        c_turno = get_col(df_filtrado, "classificacao_do_turno")
        c_taxa_ent = get_col(df_filtrado, "taxa_do_entregador")
        if not c_turno or not c_taxa_ent:
            st.info("Não encontrei `Classificação do Turno` e/ou `Taxa do entregador`.")
        else:
            aux = df_filtrado.copy()
            aux[c_taxa_ent] = pd.to_numeric(aux[c_taxa_ent], errors="coerce")
            grp2 = (aux.groupby(c_turno, dropna=True, as_index=False)[c_taxa_ent]
                        .sum()
                        .rename(columns={c_taxa_ent: "total_taxa_entregador"})
                        .sort_values("total_taxa_entregador", ascending=False))
            if grp2.empty:
                st.warning("Sem valores numéricos em `Taxa do entregador` para somar.")
            else:
                if HAS_PLOTLY:
                    fig2 = px.bar(
                        grp2,
                        x="total_taxa_entregador",
                        y=c_turno,
                        orientation="h",
                        title="Total de 'Taxa do entregador' por turno",
                        labels={"total_taxa_entregador": "Total (R$)", c_turno: "Turno"},
                        template=PLOTLY_TEMPLATE,
                    )
                    fig2.update_traces(hovertemplate="%{y}: R$ %{x:,.2f}")
                    st.plotly_chart(fig2, use_container_width=True)
                else:
                    st.info("Plotly não está disponível. Adicione `plotly>=5.22` ao requirements.txt.")
                st.dataframe(grp2, use_container_width=True)

    st.success(f"Pronto! Dados carregados de `{caminho_base}`.")

except Exception as e:
    st.error("❌ Ocorreu um erro ao carregar/processar os dados.")
    st.exception(e)

import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Painel Profissional de Ligações", page_icon="📞", layout="wide")

st.title("📞 Painel Profissional de Ligações por Ramal")
st.caption("Meta diária • Gráficos • Ranking • Telefones únicos (sem duplicar) • Filtro por período")

uploaded = st.file_uploader("📤 Envie sua planilha", type=["xlsx", "csv"])
if not uploaded:
    st.stop()

# =========================
# LEITURA INTELIGENTE (Excel com cabeçalho na linha certa)
# =========================
if uploaded.name.lower().endswith(".csv"):
    df = pd.read_csv(uploaded, dtype=str)
else:
    tmp = pd.read_excel(uploaded, header=None, dtype=str)

    header_row = None
    for i in range(min(40, len(tmp))):
        row = tmp.iloc[i].astype(str).str.strip().str.lower().tolist()
        if "data" in row and "origem" in row:
            header_row = i
            break

    if header_row is None:
        st.error("Não encontrei a linha de cabeçalho (precisa ter 'Data' e 'Origem').")
        st.stop()

    df = tmp.iloc[header_row + 1 :].copy()
    df.columns = tmp.iloc[header_row].astype(str).str.strip()

# =========================
# COLUNAS (da sua planilha)
# =========================
COL_ORIGEM = "Origem"   # ramal
COL_TELEFONE = "Destino"  # número discado
COL_DATA = "Data"       # data/hora
COL_ESTADO = "Estado"   # opcional

RAMAL_PARA_NOME = {
    "41": "Gabriela",
    "30": "Isadora",
    "33": "Cleber",
    "31": "Daniel",
    "34": "Daniel",
    "32": "Natália",
    "40": "Guilherme",
}
required = [COL_ORIGEM, COL_TELEFONE, COL_DATA]
missing = [c for c in required if c not in df.columns]
if missing:
    st.error(f"Faltam colunas: {missing}.")
    st.stop()

# =========================
# NORMALIZAÇÃO
# =========================
df[COL_ORIGEM] = df[COL_ORIGEM].astype(str).str.replace(r"\D+", "", regex=True)
df[COL_TELEFONE] = df[COL_TELEFONE].astype(str).str.replace(r"\D+", "", regex=True)
df[COL_DATA] = pd.to_datetime(df[COL_DATA], errors="coerce", dayfirst=True)
df = df[df[COL_DATA].notna()].copy()

# Coluna de dia para agregações
df["_dia"] = df[COL_DATA].dt.date

# =========================
# SIDEBAR: FILTROS
# =========================
st.sidebar.header("⚙️ Filtros")

min_date = df[COL_DATA].min().date()
max_date = df[COL_DATA].max().date()
date_range = st.sidebar.date_input(
    "📅 Período",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date
)

if isinstance(date_range, tuple) and len(date_range) == 2:
    start_date, end_date = date_range
    df = df[(df[COL_DATA].dt.date >= start_date) & (df[COL_DATA].dt.date <= end_date)]
else:
    df = df[df[COL_DATA].dt.date == date_range]

# Telefones válidos
st.sidebar.subheader("📞 Validação de telefone")
min_digits = st.sidebar.number_input("Mínimo de dígitos do Destino", min_value=6, max_value=15, value=8, step=1)
df = df[df[COL_TELEFONE].str.len() >= int(min_digits)].copy()

# Filtro por estado (se existir)
if COL_ESTADO in df.columns:
    st.sidebar.subheader("☎️ Estado (opcional)")
    estados = sorted([e for e in df[COL_ESTADO].dropna().astype(str).unique().tolist() if e.strip() != ""])
    selected_estados = st.sidebar.multiselect("Selecione estados", options=estados, default=estados)
    if selected_estados:
        df = df[df[COL_ESTADO].astype(str).isin(selected_estados)]

# =========================
# MEU RAMAL + META
# =========================
st.sidebar.header("👤 Ramal (Origem)")

ramais = sorted([r for r in df[COL_ORIGEM].dropna().astype(str).unique().tolist() if r.strip() != ""])
if not ramais:
    st.error("Não encontrei valores na coluna Origem após os filtros.")
    st.stop()

opcoes = []
for r in ramais:
    nome = RAMAL_PARA_NOME.get(str(r), "Sem nome")
    opcoes.append(f"{r} - {nome}")

selecionado = st.sidebar.selectbox("Escolha o ramal", options=opcoes)
meu_ramal = selecionado.split(" - ")[0].strip()
nome_auto = RAMAL_PARA_NOME.get(meu_ramal, "Sem nome")
meu_nome = st.sidebar.text_input("Nome para exibir", value=nome_auto)

st.sidebar.header("🎯 Metas")
meta_diaria = st.sidebar.number_input("Meta diária (números únicos)", min_value=0, value=30, step=5)
meta_periodo = st.sidebar.number_input("Meta do período (números únicos)", min_value=0, value=0, step=10)
st.sidebar.caption("Se meta do período = 0, o painel calcula automaticamente: meta_diaria × dias no período.")

# =========================
# RESUMO GERAL
# =========================
st.subheader("📌 Visão Geral (período filtrado)")

dias_no_periodo = df["_dia"].nunique()
meta_periodo_calc = meta_periodo if meta_periodo > 0 else int(meta_diaria) * int(dias_no_periodo)

colA, colB, colC, colD = st.columns(4)
colA.metric("Registros (válidos)", len(df))
colB.metric("Ramais ativos", df[COL_ORIGEM].nunique())
colC.metric("Números únicos (geral)", df[COL_TELEFONE].nunique())
colD.metric("Dias no período", dias_no_periodo)

# =========================
# MEU RAMAL: KPIs
# =========================
df_meu = df[df[COL_ORIGEM].astype(str) == str(meu_ramal)].copy()

meu_unicos = df_meu[COL_TELEFONE].nunique()
meu_total = len(df_meu)
meu_repetidos = max(0, meu_total - meu_unicos)

st.subheader("✅ Resultado do ramal selecionado")

c1, c2, c3, c4 = st.columns(4)
c1.metric(f"{meu_nome} ({meu_ramal}) - únicos", meu_unicos)
c2.metric("Total de ligações", meu_total)
c3.metric("Repetições ignoradas", meu_repetidos)
c4.metric("Meta do período", meta_periodo_calc)

progress = 0 if meta_periodo_calc == 0 else min(1.0, meu_unicos / meta_periodo_calc)
st.progress(progress)

st.success(f"{meu_nome} ({meu_ramal}) ligou para {meu_unicos} números únicos no período selecionado.")

# =========================
# GRÁFICO: EVOLUÇÃO DIÁRIA (RAMAL)
# =========================
st.subheader("📈 Evolução diária (números únicos por dia)")

daily_me = (
    df_meu.groupby("_dia")[COL_TELEFONE]
    .nunique()
    .reset_index(name="unicos_no_dia")
    .sort_values("_dia")
)
daily_me["meta_diaria"] = int(meta_diaria)

left, right = st.columns([2, 1])
with left:
    st.line_chart(daily_me.set_index("_dia")[["unicos_no_dia", "meta_diaria"]], height=260)
with right:
    st.dataframe(daily_me, use_container_width=True, height=260)

# =========================
# RANKING: POR RAMAL
# =========================
st.subheader("🏆 Ranking por ramal (números únicos no período)")

ranking = (
    df.groupby(COL_ORIGEM)
    .agg(
        ligacoes=(COL_TELEFONE, "size"),
        unicos=(COL_TELEFONE, pd.Series.nunique),
    )
    .reset_index()
)

ranking["repeticoes"] = ranking["ligacoes"] - ranking["unicos"]
ranking["Nome"] = ranking[COL_ORIGEM].astype(str).map(RAMAL_PARA_NOME).fillna("Sem nome")
ranking = ranking[["Nome", COL_ORIGEM, "ligacoes", "unicos", "repeticoes"]]
ranking = ranking.sort_values("unicos", ascending=False)

st.dataframe(ranking, use_container_width=True, height=320)

# =========================
# GRÁFICO: TOP 10 RAMAIS
# =========================
st.subheader("📊 Top 10 ramais (únicos)")

top10 = ranking.head(10).copy()
top10 = top10.set_index("Nome")[["unicos"]]
st.bar_chart(top10, height=260)

# =========================
# NÚMEROS MAIS REPETIDOS (RAMAL)
# =========================
st.subheader("🔁 Números mais repetidos do ramal selecionado")

top_rep = (
    df_meu.groupby(COL_TELEFONE)
    .size()
    .reset_index(name="tentativas")
    .sort_values("tentativas", ascending=False)
)

st.dataframe(top_rep.head(200), use_container_width=True, height=320)
st.caption("Se um número aparece com muitas tentativas, é retrabalho (ligações repetidas).")

# =========================
# DETALHES (opcional)
# =========================
with st.expander("🔎 Ver dados filtrados (para conferência)"):
    st.dataframe(df, use_container_width=True, height=380)

# =========================
# EXPORT
# =========================
def to_excel_bytes(df_to_save: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_to_save.to_excel(writer, index=False, sheet_name="DADOS_FILTRADOS")
    return output.getvalue()

st.subheader("📦 Exportar")
exp1, exp2 = st.columns(2)

with exp1:
    st.download_button(
        "⬇️ Baixar DADOS filtrados (Excel)",
        data=to_excel_bytes(df),
        file_name="ligacoes_filtradas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

with exp2:
    st.download_button(
        "⬇️ Baixar RANKING por ramal (CSV)",
        data=ranking.to_csv(index=False).encode("utf-8-sig"),
        file_name="ranking_ramais.csv",
        mime="text/csv",
        use_container_width=True
    )
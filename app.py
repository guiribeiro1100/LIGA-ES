import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Painel Profissional de Ligações", page_icon="📞", layout="wide")

st.title("📞 Painel Profissional de Ligações por Ramal")
st.caption("Ranking • Contatos únicos • Atendidas com mínimo de 2 minutos • Sem filtro por data")

uploaded = st.file_uploader("📤 Envie sua planilha", type=["xlsx", "csv"])
if not uploaded:
    st.stop()

# =========================
# LEITURA INTELIGENTE
# =========================
if uploaded.name.lower().endswith(".csv"):
    df = pd.read_csv(uploaded)
else:
    tmp = pd.read_excel(uploaded, header=None)

    header_row = None
    for i in range(min(40, len(tmp))):
        row = tmp.iloc[i].astype(str).str.strip().str.lower().tolist()
        if "data" in row and "origem" in row:
            header_row = i
            break

    if header_row is None:
        st.error("Não encontrei a linha de cabeçalho (precisa ter 'Data' e 'Origem').")
        st.stop()

    df = pd.read_excel(uploaded, header=header_row)
    df.columns = df.columns.astype(str).str.strip()

# =========================
# COLUNAS PRINCIPAIS
# =========================
COL_ORIGEM = "Origem"
COL_DESTINO = "Destino"
COL_ESTADO = "Estado"
COL_DATA = "Data"  # opcional, só para referência

RAMAL_PARA_NOME = {
    "41": "Gabriela",
    "30": "Isadora",
    "33": "Cleber",
    "31": "Daniel",
    "34": "Daniel",
    "32": "Natália",
    "40": "Guilherme",
}

required = [COL_ORIGEM, COL_DESTINO]
missing = [c for c in required if c not in df.columns]
if missing:
    st.error(f"Faltam colunas obrigatórias: {missing}")
    st.stop()

# =========================
# DETECTAR COLUNA DE DURAÇÃO AUTOMATICAMENTE
# =========================
COL_DURACAO = None
for col in df.columns:
    nome_col = str(col).strip().lower()
    if "dura" in nome_col or "tempo" in nome_col or "duration" in nome_col:
        COL_DURACAO = col
        break

# =========================
# NORMALIZAÇÃO
# =========================
df[COL_ORIGEM] = df[COL_ORIGEM].astype(str).str.replace(r"\D+", "", regex=True)
df[COL_DESTINO] = df[COL_DESTINO].astype(str).str.strip()

# remove destino vazio
df = df[df[COL_DESTINO].notna() & (df[COL_DESTINO] != "")].copy()

# coluna auxiliar só números para contagem de únicos
df["_destino_limpo"] = df[COL_DESTINO].astype(str).str.replace(r"\D+", "", regex=True)
df["_destino_final"] = df["_destino_limpo"].where(df["_destino_limpo"] != "", df[COL_DESTINO].astype(str))

# normaliza duração em segundos
if COL_DURACAO is not None:
    df[COL_DURACAO] = (
        df[COL_DURACAO]
        .astype(str)
        .str.replace(r"[^\d]", "", regex=True)
    )
    df[COL_DURACAO] = pd.to_numeric(df[COL_DURACAO], errors="coerce").fillna(0)
else:
    df["_duracao_segundos"] = 0
    COL_DURACAO = "_duracao_segundos"

# =========================
# SIDEBAR
# =========================
st.sidebar.header("⚙️ Filtros")

st.sidebar.markdown("### 🔎 Conferência")
st.sidebar.write("Linhas no arquivo:", len(df))
st.sidebar.write("Coluna de duração encontrada:", COL_DURACAO)

# filtro por estado opcional
if COL_ESTADO in df.columns:
    filtrar_estado = st.sidebar.checkbox("Filtrar por Estado", value=False)
    if filtrar_estado:
        estados = sorted([e for e in df[COL_ESTADO].dropna().astype(str).unique().tolist() if e.strip() != ""])
        selected_estados = st.sidebar.multiselect("Selecione estados", options=estados, default=estados)
        if selected_estados:
            df = df[df[COL_ESTADO].astype(str).isin(selected_estados)]

# =========================
# RAMAL
# =========================
st.sidebar.header("👤 Ramal (Origem)")

ramais_dados = df[COL_ORIGEM].dropna().astype(str).unique().tolist()
ramais = sorted(set(list(RAMAL_PARA_NOME.keys()) + ramais_dados))

opcoes = []
for r in ramais:
    nome = RAMAL_PARA_NOME.get(str(r), "Sem nome")
    opcoes.append(f"{r} - {nome}")

selecionado = st.sidebar.selectbox("Escolha o ramal", options=opcoes)
meu_ramal = selecionado.split(" - ")[0].strip()
nome_auto = RAMAL_PARA_NOME.get(meu_ramal, "Sem nome")
meu_nome = st.sidebar.text_input("Nome para exibir", value=nome_auto)

# =========================
# FUNÇÕES AUXILIARES
# =========================
def contar_atendidas(df_base: pd.DataFrame, col_estado: str, col_duracao: str) -> int:
    if col_estado not in df_base.columns:
        return 0
    estado_upper = df_base[col_estado].astype(str).str.upper().str.strip()
    return int(((estado_upper == "ATENDIDA") & (df_base[col_duracao] >= 120)).sum())

def contar_nao_atendidas(df_base: pd.DataFrame, col_estado: str) -> int:
    if col_estado not in df_base.columns:
        return 0
    estado_upper = df_base[col_estado].astype(str).str.upper().str.strip()
    return int(((estado_upper == "NÃO ATENDIDA") | (estado_upper == "NAO ATENDIDA")).sum())

def contar_falhou(df_base: pd.DataFrame, col_estado: str) -> int:
    if col_estado not in df_base.columns:
        return 0
    estado_upper = df_base[col_estado].astype(str).str.upper().str.strip()
    return int((estado_upper == "FALHOU").sum())

def contar_congestion(df_base: pd.DataFrame, col_estado: str) -> int:
    if col_estado not in df_base.columns:
        return 0
    estado_upper = df_base[col_estado].astype(str).str.upper().str.strip()
    return int((estado_upper == "CONGESTION").sum())

# =========================
# VISÃO GERAL
# =========================
st.subheader("📌 Visão Geral (todos os registros)")

total_ligacoes = len(df)
ramais_ativos = df[COL_ORIGEM].nunique()
contatos_unicos_geral = df["_destino_final"].nunique()

colA, colB, colC = st.columns(3)
colA.metric("Registros (total)", int(total_ligacoes))
colB.metric("Ramais ativos", int(ramais_ativos))
colC.metric("Contatos únicos (geral)", int(contatos_unicos_geral))

# =========================
# MÉTRICAS GERAIS DE ATENDIMENTO
# =========================
if COL_ESTADO in df.columns:
    atendidas = contar_atendidas(df, COL_ESTADO, COL_DURACAO)
    nao_atendidas = contar_nao_atendidas(df, COL_ESTADO)
    falhou = contar_falhou(df, COL_ESTADO)
    congestion = contar_congestion(df, COL_ESTADO)
    taxa_atendimento = (atendidas / total_ligacoes * 100) if total_ligacoes > 0 else 0

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("📞 Total de ligações", int(total_ligacoes))
    c2.metric("✅ Atendidas (+2 min)", int(atendidas))
    c3.metric("❌ Não atendidas", int(nao_atendidas))
    c4.metric("⚠️ Falhou", int(falhou))
    c5.metric("📶 Congestionadas", int(congestion))
    c6.metric("📊 Taxa de atendimento", f"{taxa_atendimento:.1f}%")

# =========================
# RESULTADO DO RAMAL
# =========================
df_meu = df[df[COL_ORIGEM].astype(str) == str(meu_ramal)].copy()

meu_unicos = df_meu["_destino_final"].nunique()
meu_total = len(df_meu)
meu_repetidos = max(0, meu_total - meu_unicos)

st.subheader("✅ Resultado do ramal selecionado")

c1, c2, c3 = st.columns(3)
c1.metric(f"{meu_nome} ({meu_ramal}) - contatos únicos", int(meu_unicos))
c2.metric("Total de ligações", int(meu_total))
c3.metric("Repetições ignoradas", int(meu_repetidos))

if COL_ESTADO in df_meu.columns:
    atendidas_meu = contar_atendidas(df_meu, COL_ESTADO, COL_DURACAO)
    nao_atendidas_meu = contar_nao_atendidas(df_meu, COL_ESTADO)
    falhou_meu = contar_falhou(df_meu, COL_ESTADO)
    congestion_meu = contar_congestion(df_meu, COL_ESTADO)
    taxa_atendimento_meu = (atendidas_meu / meu_total * 100) if meu_total > 0 else 0

    c4, c5, c6, c7, c8 = st.columns(5)
    c4.metric("✅ Atendidas (+2 min)", int(atendidas_meu))
    c5.metric("❌ Não atendidas", int(nao_atendidas_meu))
    c6.metric("⚠️ Falhou", int(falhou_meu))
    c7.metric("📶 Congestionadas", int(congestion_meu))
    c8.metric("📊 Taxa de atendimento", f"{taxa_atendimento_meu:.1f}%")

st.success(f"{meu_nome} ({meu_ramal}) ligou para {meu_unicos} contatos únicos no total.")

# =========================
# RANKING POR RAMAL
# =========================
st.subheader("🏆 Ranking por ramal (contatos únicos)")

ranking = (
    df.groupby(COL_ORIGEM)
    .agg(
        ligacoes=("_destino_final", "size"),
        unicos=("_destino_final", pd.Series.nunique),
    )
    .reset_index()
)

ranking["repeticoes"] = ranking["ligacoes"] - ranking["unicos"]
ranking["Nome"] = ranking[COL_ORIGEM].astype(str).map(RAMAL_PARA_NOME).fillna("Sem nome")

if COL_ESTADO in df.columns:
    atendidas_por_ramal = (
        df.assign(_atendida_regra=((df[COL_ESTADO].astype(str).str.upper().str.strip() == "ATENDIDA") & (df[COL_DURACAO] >= 60)).astype(int))
        .groupby(COL_ORIGEM)["_atendida_regra"]
        .sum()
        .reset_index(name="atendidas_2min")
    )

    ranking = ranking.merge(atendidas_por_ramal, on=COL_ORIGEM, how="left")
    ranking["atendidas_2min"] = ranking["atendidas_2min"].fillna(0).astype(int)
    ranking["taxa_atendimento"] = ranking.apply(
        lambda row: (row["atendidas_2min"] / row["ligacoes"] * 100) if row["ligacoes"] > 0 else 0,
        axis=1
    )
    ranking = ranking[["Nome", COL_ORIGEM, "ligacoes", "unicos", "repeticoes", "atendidas_2min", "taxa_atendimento"]]
    ranking["taxa_atendimento"] = ranking["taxa_atendimento"].map(lambda x: f"{x:.1f}%")
else:
    ranking = ranking[["Nome", COL_ORIGEM, "ligacoes", "unicos", "repeticoes"]]

ranking = ranking.sort_values("unicos", ascending=False)
st.dataframe(ranking, use_container_width=True, height=320)

# =========================
# CONTATOS MAIS REPETIDOS DO RAMAL
# =========================
st.subheader("🔁 Contatos mais repetidos do ramal selecionado")

top_rep = (
    df_meu.groupby("_destino_final")
    .size()
    .reset_index(name="tentativas")
    .sort_values("tentativas", ascending=False)
)

st.dataframe(top_rep.head(300), use_container_width=True, height=320)

# =========================
# DADOS COMPLETOS
# =========================
with st.expander("🔎 Ver dados completos (para conferência)"):
    st.dataframe(df, use_container_width=True, height=420)

# =========================
# EXPORT
# =========================
def to_excel_bytes(df_to_save: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_to_save.to_excel(writer, index=False, sheet_name="DADOS")
    return output.getvalue()

st.subheader("📦 Exportar")
exp1, exp2 = st.columns(2)

with exp1:
    st.download_button(
        "⬇️ Baixar DADOS (Excel)",
        data=to_excel_bytes(df),
        file_name="ligacoes_todos_os_registros.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

with exp2:
    st.download_button(
        "⬇️ Baixar RANKING (CSV)",
        data=ranking.to_csv(index=False).encode("utf-8-sig"),
        file_name="ranking_ramais.csv",
        mime="text/csv",
        use_container_width=True
    )
import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Painel Profissional de Ligações", page_icon="📞", layout="wide")

st.title("📞 Painel Profissional de Ligações por Ramal")
st.caption("Contatos únicos • Ranking • Atendidas com mínimo de 1 minuto")

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
# DETECTAR COLUNA DURAÇÃO
# =========================
COL_DURACAO = None

for col in df.columns:
    nome = str(col).lower()
    if "dura" in nome or "tempo" in nome or "duration" in nome:
        COL_DURACAO = col
        break

# =========================
# NORMALIZAÇÃO
# =========================
df[COL_ORIGEM] = df[COL_ORIGEM].astype(str).str.replace(r"\D+", "", regex=True)

df[COL_DESTINO] = df[COL_DESTINO].astype(str).str.strip()

df = df[df[COL_DESTINO] != ""]

df["_destino_limpo"] = df[COL_DESTINO].astype(str).str.replace(r"\D+", "", regex=True)

df["_destino_final"] = df["_destino_limpo"].where(df["_destino_limpo"] != "", df[COL_DESTINO])

# =========================
# DURAÇÃO EM SEGUNDOS
# =========================
if COL_DURACAO:

    df[COL_DURACAO] = (
        df[COL_DURACAO]
        .astype(str)
        .str.replace(r"[^\d]", "", regex=True)
    )

    df[COL_DURACAO] = pd.to_numeric(df[COL_DURACAO], errors="coerce").fillna(0)

else:

    df["_duracao_fake"] = 0
    COL_DURACAO = "_duracao_fake"

# =========================
# SIDEBAR
# =========================
st.sidebar.header("Informações")

st.sidebar.write("Total de registros:", len(df))
st.sidebar.write("Coluna duração detectada:", COL_DURACAO)

# =========================
# ESCOLHER RAMAL
# =========================
st.sidebar.header("Ramal")

ramais_dados = df[COL_ORIGEM].unique().tolist()

ramais = sorted(set(list(RAMAL_PARA_NOME.keys()) + ramais_dados))

opcoes = []

for r in ramais:
    nome = RAMAL_PARA_NOME.get(r, "Sem nome")
    opcoes.append(f"{r} - {nome}")

selecionado = st.sidebar.selectbox("Escolha o ramal", opcoes)

meu_ramal = selecionado.split(" - ")[0]

meu_nome = RAMAL_PARA_NOME.get(meu_ramal, "Sem nome")

# =========================
# FUNÇÕES
# =========================
def atendidas(df_base):

    if COL_ESTADO not in df_base.columns:
        return 0

    estado = df_base[COL_ESTADO].astype(str).str.upper().str.strip()

    return int(((estado == "ATENDIDA") & (df_base[COL_DURACAO] >= 60)).sum())

def nao_atendidas(df_base):

    estado = df_base[COL_ESTADO].astype(str).str.upper().str.strip()

    return int(((estado == "NÃO ATENDIDA") | (estado == "NAO ATENDIDA")).sum())

def falhou(df_base):

    estado = df_base[COL_ESTADO].astype(str).str.upper().str.strip()

    return int((estado == "FALHOU").sum())

def congestion(df_base):

    estado = df_base[COL_ESTADO].astype(str).str.upper().str.strip()

    return int((estado == "CONGESTION").sum())

# =========================
# VISÃO GERAL
# =========================
st.subheader("📊 Visão Geral")

total = len(df)

ramais_ativos = df[COL_ORIGEM].nunique()

contatos_unicos = df["_destino_final"].nunique()

c1, c2, c3 = st.columns(3)

c1.metric("Total de ligações", total)
c2.metric("Ramais ativos", ramais_ativos)
c3.metric("Contatos únicos", contatos_unicos)

# =========================
# MÉTRICAS GERAIS
# =========================
if COL_ESTADO in df.columns:

    att = atendidas(df)

    nao = nao_atendidas(df)

    fal = falhou(df)

    con = congestion(df)

    taxa = (att / total * 100) if total > 0 else 0

    c1, c2, c3, c4, c5, c6 = st.columns(6)

    c1.metric("📞 Total", total)
    c2.metric("✅ Atendidas (+1min)", att)
    c3.metric("❌ Não atendidas", nao)
    c4.metric("⚠️ Falhou", fal)
    c5.metric("📶 Congestion", con)
    c6.metric("📊 Taxa atendimento", f"{taxa:.1f}%")

# =========================
# RAMAL
# =========================
df_meu = df[df[COL_ORIGEM] == meu_ramal]

st.subheader(f"👤 {meu_nome} ({meu_ramal})")

meu_total = len(df_meu)

meu_unicos = df_meu["_destino_final"].nunique()

meu_rep = meu_total - meu_unicos

c1, c2, c3 = st.columns(3)

c1.metric("Contatos únicos", meu_unicos)

c2.metric("Total ligações", meu_total)

c3.metric("Repetições", meu_rep)

if COL_ESTADO in df.columns:

    att_meu = atendidas(df_meu)

    nao_meu = nao_atendidas(df_meu)

    fal_meu = falhou(df_meu)

    con_meu = congestion(df_meu)

    taxa_meu = (att_meu / meu_total * 100) if meu_total > 0 else 0

    c1, c2, c3, c4, c5 = st.columns(5)

    c1.metric("Atendidas (+1min)", att_meu)
    c2.metric("Não atendidas", nao_meu)
    c3.metric("Falhou", fal_meu)
    c4.metric("Congestion", con_meu)
    c5.metric("Taxa", f"{taxa_meu:.1f}%")

# =========================
# RANKING
# =========================
st.subheader("🏆 Ranking")

ranking = (

    df.groupby(COL_ORIGEM)

    .agg(

        ligacoes=("_destino_final", "size"),

        unicos=("_destino_final", pd.Series.nunique)

    )

    .reset_index()

)

ranking["Nome"] = ranking[COL_ORIGEM].map(RAMAL_PARA_NOME).fillna("Sem nome")

ranking = ranking.sort_values("unicos", ascending=False)

st.dataframe(ranking, use_container_width=True)

# =========================
# MAIS REPETIDOS
# =========================
st.subheader("🔁 Números mais repetidos")

top_rep = (

    df_meu.groupby("_destino_final")

    .size()

    .reset_index(name="tentativas")

    .sort_values("tentativas", ascending=False)

)

st.dataframe(top_rep.head(100), use_container_width=True)

# =========================
# EXPORT
# =========================
def to_excel(df):

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        df.to_excel(writer, index=False)

    return output.getvalue()

st.download_button(

    "Baixar dados",

    to_excel(df),

    "dados_ligacoes.xlsx",

    use_container_width=True

)
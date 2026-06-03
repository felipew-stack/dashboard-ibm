import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(
    page_title="Dashboard IBM",
    layout="wide"
)

# ==========================
# CABEÇALHO COM LOGO
# ==========================

logo_col, titulo_col = st.columns([1, 4])

with logo_col:
    st.image("logo_3am.png", width=180)

with titulo_col:
    st.title("Dashboard de Localidades IBM")
    st.caption("Monitoramento de localidades abertas, pendentes e concluídas")

# ==========================
# UPLOAD DO EXCEL
# ==========================

st.subheader("Importação do Relatório")

arquivo_upload = st.file_uploader(
    "Selecione o arquivo Excel atualizado",
    type=["xlsx"]
)

if arquivo_upload is None:
    st.info("Carregue o arquivo Excel para visualizar o dashboard.")
    st.stop()

# ==========================
# LEITURA DO EXCEL
# ==========================

df_total = pd.read_excel(arquivo_upload, sheet_name="Localidades Total")
df_pendentes = pd.read_excel(
    arquivo_upload, sheet_name="Localidades Pendentes")

# ==========================
# AJUSTES
# ==========================

df_total["Cidade"] = df_total["Cidade"].astype(str).str.strip()
df_pendentes["Cidade"] = df_pendentes["Cidade"].astype(str).str.strip()

if "Data da Solicitação" in df_total.columns:
    df_total["Data da Solicitação"] = pd.to_datetime(
        df_total["Data da Solicitação"],
        errors="coerce"
    ).dt.strftime("%d/%m/%Y")

if "Data da Solicitação" in df_pendentes.columns:
    df_pendentes["Data da Solicitação"] = pd.to_datetime(
        df_pendentes["Data da Solicitação"],
        errors="coerce"
    ).dt.strftime("%d/%m/%Y")

df_concluidas = df_total[
    ~df_total["Cidade"].isin(df_pendentes["Cidade"])
]

# ==========================
# CÁLCULOS GERAIS
# ==========================

total = len(df_total)
pendentes = len(df_pendentes)
concluidas = len(df_concluidas)
percentual = round((concluidas / total) * 100, 1) if total > 0 else 0

alta = len(df_pendentes[df_pendentes["Prioridade"] == "Alta"])
media = len(df_pendentes[df_pendentes["Prioridade"] == "Média"])
baixa = len(df_pendentes[df_pendentes["Prioridade"] == "Baixa"])

st.divider()

# ==========================
# INDICADORES
# ==========================

col1, col2, col3, col4 = st.columns(4)

col1.metric("Total de Localidades", total)
col2.metric("Localidades Concluídas", concluidas)
col3.metric("Localidades Pendentes", pendentes)
col4.metric("Percentual de Conclusão", f"{percentual}%")

st.progress(percentual / 100)

col5, col6, col7 = st.columns(3)

col5.metric("🔴 Prioridade Alta", alta)
col6.metric("🟡 Prioridade Média", media)
col7.metric("🟢 Prioridade Baixa", baixa)

st.divider()

# ==========================
# GRÁFICOS
# ==========================

graf1, graf2 = st.columns(2)

with graf1:
    dados_status = pd.DataFrame({
        "Status": ["Concluídas", "Pendentes"],
        "Quantidade": [concluidas, pendentes]
    })

    fig_status = px.pie(
        dados_status,
        names="Status",
        values="Quantidade",
        hole=0.55,
        title="Status Geral das Localidades"
    )

    st.plotly_chart(fig_status, use_container_width=True)

with graf2:
    uf_count = df_pendentes["UF"].value_counts().reset_index()
    uf_count.columns = ["UF", "Quantidade"]

    fig_uf = px.bar(
        uf_count,
        x="UF",
        y="Quantidade",
        text_auto=True,
        title="Pendências por Estado"
    )

    st.plotly_chart(fig_uf, use_container_width=True)

st.divider()

# ==========================
# FILTROS
# ==========================

st.subheader("Filtros de Localidades Pendentes")

filtro1, filtro2, filtro3 = st.columns(3)

with filtro1:
    busca = st.text_input("Pesquisar cidade")

with filtro2:
    prioridade = st.selectbox(
        "Filtrar por prioridade",
        ["Todas", "Alta", "Média", "Baixa"]
    )

with filtro3:
    uf = st.selectbox(
        "Filtrar por UF",
        ["Todas"] + sorted(df_pendentes["UF"].dropna().unique().tolist())
    )

# ==========================
# TABELA EDITÁVEL
# ==========================

tabela = df_pendentes.copy()

if busca:
    tabela = tabela[
        tabela["Cidade"].str.contains(busca, case=False, na=False)
    ]

if prioridade != "Todas":
    tabela = tabela[tabela["Prioridade"] == prioridade]

if uf != "Todas":
    tabela = tabela[tabela["UF"] == uf]

st.subheader("Localidades Pendentes")
st.caption("Altere a prioridade diretamente na tabela, se necessário.")

tabela_editada = st.data_editor(
    tabela,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Prioridade": st.column_config.SelectboxColumn(
            "Prioridade",
            options=["Alta", "Média", "Baixa"],
            required=True
        )
    },
    disabled=[
        coluna for coluna in tabela.columns if coluna != "Prioridade"
    ]
)

# ==========================
# LOCALIDADES CONCLUÍDAS
# ==========================

with st.expander("Ver localidades concluídas"):
    st.dataframe(
        df_concluidas,
        use_container_width=True,
        hide_index=True
    )

# ==========================
# DOWNLOADS
# ==========================

st.divider()
st.subheader("Exportação de Relatórios")

col_down1, col_down2 = st.columns(2)

with col_down1:
    st.download_button(
        label="Baixar pendentes editadas em CSV",
        data=tabela_editada.to_csv(index=False).encode("utf-8-sig"),
        file_name="localidades_pendentes_editadas.csv",
        mime="text/csv"
    )

with col_down2:
    st.download_button(
        label="Baixar concluídas em CSV",
        data=df_concluidas.to_csv(index=False).encode("utf-8-sig"),
        file_name="localidades_concluidas.csv",
        mime="text/csv"
    )

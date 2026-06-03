import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(
    page_title="Dashboard IBM",
    layout="wide"
)

# ==========================
# CONFIGURAÇÃO DO ARQUIVO
# ==========================

ARQUIVO_EXCEL = "Report IBM 01-06 (1) (3).xlsx"

# ==========================
# CABEÇALHO COM LOGO
# ==========================

logo_col, titulo_col = st.columns([1, 4])

with logo_col:
    st.image("logo_3am.png", width=180)

with titulo_col:
    st.title("Projeto IBM")
    st.caption(
        "Monitoramento de localidades abertas, pendentes, concluídas e relatório detalhado"
    )

# ==========================
# FUNÇÃO PARA CORRIGIR TEXTOS
# ==========================


def corrigir_relatorio(texto):
    if pd.isna(texto):
        return ""

    texto = str(texto).strip()

    correcoes = {
        "Encontramos recursos na localidades": "Encontramos recursos na localidade",
        "Encontramos recursos nas localidades": "Encontramos recursos na localidade",
        "Encontramos recursos no localidades": "Encontramos recursos na localidade",
        "Encontramos recursos na localidade, mas declinaram": "Encontramos recursos na localidade, mas eles declinaram",
        "porem": "porém",
        "Porem": "Porém",
        "tecnicos": "técnicos",
        "Tecnicos": "Técnicos",
        "nao": "não",
        "Nao": "Não",
        "deu retorno": "retornou",
        "nega a proposta": "negou a proposta",
        "negou a proprosta": "negou a proposta",
        "proprosta": "proposta",
        "Buscando": "Em busca de recurso técnico para a localidade."
    }

    for errado, certo in correcoes.items():
        texto = texto.replace(errado, certo)

    return texto


# ==========================
# LEITURA DO EXCEL
# ==========================

try:
    df_total = pd.read_excel(
        ARQUIVO_EXCEL,
        sheet_name="Localidades Total"
    )

    df_pendentes = pd.read_excel(
        ARQUIVO_EXCEL,
        sheet_name="Localidades Pendentes"
    )

except FileNotFoundError:
    st.error(f"Arquivo '{ARQUIVO_EXCEL}' não encontrado na pasta do projeto.")
    st.stop()

except Exception as erro:
    st.error("Erro ao carregar a planilha.")
    st.write(erro)
    st.stop()

# ==========================
# AJUSTES
# ==========================

df_total["Cidade"] = df_total["Cidade"].astype(str).str.strip()
df_pendentes["Cidade"] = df_pendentes["Cidade"].astype(str).str.strip()

for df in [df_total, df_pendentes]:
    if "Data da Solicitação" in df.columns:
        df["Data da Solicitação"] = pd.to_datetime(
            df["Data da Solicitação"],
            errors="coerce"
        ).dt.strftime("%d/%m/%Y")

    if "Previsão" in df.columns:
        df["Previsão"] = pd.to_datetime(
            df["Previsão"],
            errors="coerce"
        ).dt.strftime("%d/%m/%Y")

if "Relatório Detalhado" in df_pendentes.columns:
    df_pendentes["Relatório Detalhado"] = df_pendentes[
        "Relatório Detalhado"
    ].apply(corrigir_relatorio)

df_concluidas = df_total[
    ~df_total["Cidade"].isin(df_pendentes["Cidade"])
]

df_relatorio = df_pendentes.copy()

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

# ==========================
# ABAS
# ==========================

aba_dashboard, aba_pendentes, aba_relatorio, aba_concluidas = st.tabs([
    "Dashboard",
    "Localidades Pendentes",
    "Relatório Detalhado",
    "Localidades Concluídas"
])

# ==========================
# ABA DASHBOARD
# ==========================

with aba_dashboard:
    st.divider()

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

# ==========================
# ABA LOCALIDADES PENDENTES
# ==========================

with aba_pendentes:
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

    st.dataframe(
        tabela,
        use_container_width=True,
        hide_index=True
    )

# ==========================
# ABA RELATÓRIO DETALHADO
# ==========================

with aba_relatorio:
    st.subheader("Relatório Detalhado das Localidades")
    st.caption("Situação atual de cada localidade pendente.")

    if "Relatório Detalhado" not in df_relatorio.columns:
        st.warning(
            "A coluna 'Relatório Detalhado' não foi encontrada na planilha.")
    else:
        col_rel1, col_rel2, col_rel3 = st.columns(3)

        with col_rel1:
            busca_relatorio = st.text_input(
                "Pesquisar localidade no relatório")

        with col_rel2:
            uf_relatorio = st.selectbox(
                "Filtrar UF no relatório",
                ["Todas"] +
                sorted(df_relatorio["UF"].dropna().unique().tolist())
            )

        with col_rel3:
            prioridade_relatorio = st.selectbox(
                "Filtrar prioridade no relatório",
                ["Todas", "Alta", "Média", "Baixa"]
            )

        tabela_relatorio = df_relatorio.copy()

        if busca_relatorio:
            tabela_relatorio = tabela_relatorio[
                tabela_relatorio["Cidade"].astype(str).str.contains(
                    busca_relatorio,
                    case=False,
                    na=False
                )
            ]

        if uf_relatorio != "Todas":
            tabela_relatorio = tabela_relatorio[
                tabela_relatorio["UF"] == uf_relatorio
            ]

        if prioridade_relatorio != "Todas":
            tabela_relatorio = tabela_relatorio[
                tabela_relatorio["Prioridade"] == prioridade_relatorio
            ]

        st.write(f"Total filtrado: {len(tabela_relatorio)} localidade(s)")

        for _, linha in tabela_relatorio.iterrows():
            cidade = linha.get("Cidade", "")
            uf = linha.get("UF", "")
            prioridade = linha.get("Prioridade", "")
            data = linha.get("Data da Solicitação", "")
            previsao = linha.get("Previsão", "")
            relatorio = linha.get("Relatório Detalhado", "")

            with st.container(border=True):
                st.markdown(f"### {cidade} - {uf}")
                st.markdown(f"**Prioridade:** {prioridade}")
                st.markdown(f"**Data da Solicitação:** {data}")
                st.markdown(f"**Previsão:** {previsao}")
                st.markdown("**Situação:**")
                st.write(relatorio)

# ==========================
# ABA LOCALIDADES CONCLUÍDAS
# ==========================

with aba_concluidas:
    st.subheader("Localidades Concluídas")

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

col_down1, col_down2, col_down3 = st.columns(3)

with col_down1:
    st.download_button(
        label="Baixar pendentes em CSV",
        data=df_pendentes.to_csv(index=False).encode("utf-8-sig"),
        file_name="localidades_pendentes.csv",
        mime="text/csv"
    )

with col_down2:
    st.download_button(
        label="Baixar concluídas em CSV",
        data=df_concluidas.to_csv(index=False).encode("utf-8-sig"),
        file_name="localidades_concluidas.csv",
        mime="text/csv"
    )

with col_down3:
    st.download_button(
        label="Baixar relatório detalhado em CSV",
        data=df_relatorio.to_csv(index=False).encode("utf-8-sig"),
        file_name="relatorio_detalhado.csv",
        mime="text/csv"
    )

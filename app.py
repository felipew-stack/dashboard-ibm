import streamlit as st
import pandas as pd
import plotly.express as px
import base64
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Dashboard IBM", layout="wide")

ARQUIVO_EXCEL = "Report IBM 01-06 (1) (3).xlsx"
ARQUIVO_LOGO = "logo_3am.png"


def imagem_base64(caminho):
    try:
        with open(caminho, "rb") as arquivo:
            return base64.b64encode(arquivo.read()).decode()
    except FileNotFoundError:
        return None


def salvar_excel(df_total, df_pendentes):
    with pd.ExcelWriter(
        ARQUIVO_EXCEL,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace"
    ) as writer:
        df_total.to_excel(writer, sheet_name="Localidades Total", index=False)
        df_pendentes.to_excel(
            writer,
            sheet_name="Localidades Pendentes",
            index=False
        )


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


def gerar_excel_relatorio(df):
    output = BytesIO()

    df_export = df.copy()

    colunas_preferidas = [
        "UF",
        "Cidade",
        "Data da Solicitação",
        "Previsão",
        "Prioridade",
        "Relatório Detalhado"
    ]

    colunas_existentes = [
        coluna for coluna in colunas_preferidas
        if coluna in df_export.columns
    ]

    df_export = df_export[colunas_existentes]

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_export.to_excel(
            writer,
            index=False,
            sheet_name="Relatório Detalhado"
        )

        worksheet = writer.sheets["Relatório Detalhado"]

        azul_escuro = "003B71"
        branco = "FFFFFF"
        cinza_claro = "F2F6FA"
        vermelho = "F8D7DA"
        amarelo = "FFF3CD"
        verde = "D4EDDA"

        borda_fina = Side(style="thin", color="D9E2EC")

        for cell in worksheet[1]:
            cell.fill = PatternFill(
                start_color=azul_escuro,
                end_color=azul_escuro,
                fill_type="solid"
            )
            cell.font = Font(color=branco, bold=True)
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )
            cell.border = Border(
                left=borda_fina,
                right=borda_fina,
                top=borda_fina,
                bottom=borda_fina
            )

        worksheet.row_dimensions[1].height = 28

        for row in worksheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(
                    vertical="top",
                    wrap_text=True
                )
                cell.border = Border(
                    left=borda_fina,
                    right=borda_fina,
                    top=borda_fina,
                    bottom=borda_fina
                )

                if cell.row % 2 == 0:
                    cell.fill = PatternFill(
                        start_color=cinza_claro,
                        end_color=cinza_claro,
                        fill_type="solid"
                    )

        if "Prioridade" in df_export.columns:
            coluna_prioridade = df_export.columns.get_loc("Prioridade") + 1

            for row in range(2, worksheet.max_row + 1):
                celula = worksheet.cell(row=row, column=coluna_prioridade)
                valor = str(celula.value).strip()

                if valor == "Alta":
                    celula.fill = PatternFill(
                        start_color=vermelho,
                        end_color=vermelho,
                        fill_type="solid"
                    )
                    celula.font = Font(bold=True)

                elif valor == "Média":
                    celula.fill = PatternFill(
                        start_color=amarelo,
                        end_color=amarelo,
                        fill_type="solid"
                    )
                    celula.font = Font(bold=True)

                elif valor == "Baixa":
                    celula.fill = PatternFill(
                        start_color=verde,
                        end_color=verde,
                        fill_type="solid"
                    )
                    celula.font = Font(bold=True)

        larguras = {
            "UF": 10,
            "Cidade": 28,
            "Data da Solicitação": 22,
            "Previsão": 18,
            "Prioridade": 16,
            "Relatório Detalhado": 90
        }

        for idx, coluna in enumerate(df_export.columns, start=1):
            letra = get_column_letter(idx)
            worksheet.column_dimensions[letra].width = larguras.get(coluna, 25)

        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

    return output.getvalue()


try:
    df_total = pd.read_excel(ARQUIVO_EXCEL, sheet_name="Localidades Total")
    df_pendentes = pd.read_excel(
        ARQUIVO_EXCEL,
        sheet_name="Localidades Pendentes"
    )
except FileNotFoundError:
    st.error(f"Arquivo '{ARQUIVO_EXCEL}' não encontrado.")
    st.stop()
except Exception as erro:
    st.error("Erro ao carregar a planilha.")
    st.write(erro)
    st.stop()


colunas_padrao = [
    "UF",
    "Cidade",
    "Data da Solicitação",
    "Previsão",
    "Prioridade",
    "Relatório Detalhado"
]

for coluna in colunas_padrao:
    if coluna not in df_total.columns:
        df_total[coluna] = ""

    if coluna not in df_pendentes.columns:
        df_pendentes[coluna] = ""


df_total["Cidade"] = df_total["Cidade"].astype(str).str.strip()
df_total["UF"] = df_total["UF"].astype(str).str.strip().str.upper()

df_pendentes["Cidade"] = df_pendentes["Cidade"].astype(str).str.strip()
df_pendentes["UF"] = df_pendentes["UF"].astype(str).str.strip().str.upper()
df_pendentes["Prioridade"] = df_pendentes["Prioridade"].astype(str).str.strip()

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
    ~(
        df_total["Cidade"].astype(str).str.upper().isin(
            df_pendentes["Cidade"].astype(str).str.upper()
        )
    )
]

total = len(df_total)
pendentes = len(df_pendentes)
concluidas = len(df_concluidas)
percentual = round((concluidas / total) * 100, 1) if total > 0 else 0

alta = len(df_pendentes[df_pendentes["Prioridade"] == "Alta"])
media = len(df_pendentes[df_pendentes["Prioridade"] == "Média"])
baixa = len(df_pendentes[df_pendentes["Prioridade"] == "Baixa"])


st.markdown(
    """
    <style>
        .logo-box {
            background-color: #FFFFFF;
            padding: 14px 20px;
            border-radius: 16px;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            box-shadow: 0px 2px 10px rgba(0,0,0,0.12);
            border: 1px solid rgba(0,0,0,0.08);
        }

        .logo-box img {
            width: 230px;
        }

        .header-title {
            padding-top: 12px;
        }
    </style>
    """,
    unsafe_allow_html=True
)


logo_col, titulo_col = st.columns([1.2, 4])

with logo_col:
    logo_base64 = imagem_base64(ARQUIVO_LOGO)

    if logo_base64:
        st.markdown(
            f"""
            <div class="logo-box">
                <img src="data:image/png;base64,{logo_base64}">
            </div>
            """,
            unsafe_allow_html=True
        )
    else:
        st.warning(f"Logo '{ARQUIVO_LOGO}' não encontrada.")

with titulo_col:
    st.markdown('<div class="header-title">', unsafe_allow_html=True)
    st.title("Projeto IBM")
    st.caption(
        "Monitoramento de localidades abertas, pendentes, concluídas e relatório detalhado"
    )
    st.markdown("</div>", unsafe_allow_html=True)


aba_dashboard, aba_relatorio, aba_total, aba_concluidas = st.tabs([
    "Dashboard",
    "Relatório Detalhado",
    "Localidades Total",
    "Localidades Concluídas"
])


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

    st.subheader("Adicionar Nova Localidade")
    st.caption(
        "Cadastre uma nova localidade. Ela será adicionada em Localidades Total e também em Localidades Pendentes."
    )

    with st.expander("Adicionar nova localidade", expanded=False):
        add_col1, add_col2, add_col3 = st.columns(3)

        with add_col1:
            nova_uf = st.text_input("UF", max_chars=2).upper()
            nova_cidade = st.text_input("Cidade")

        with add_col2:
            nova_data = st.date_input("Data da Solicitação")
            nova_previsao = st.date_input("Previsão")

        with add_col3:
            nova_prioridade = st.selectbox(
                "Prioridade",
                ["Alta", "Média", "Baixa"]
            )

            novo_relatorio = st.text_area(
                "Relatório Detalhado",
                height=120,
                placeholder="Digite a situação da localidade..."
            )

        if st.button("Adicionar Localidade"):
            if not nova_uf or not nova_cidade:
                st.warning("Preencha pelo menos UF e Cidade.")
            else:
                cidade_formatada = nova_cidade.strip().title()
                uf_formatada = nova_uf.strip().upper()

                existe_total = df_total[
                    (df_total["Cidade"].astype(str).str.upper() == cidade_formatada.upper()) &
                    (df_total["UF"].astype(str).str.upper()
                     == uf_formatada.upper())
                ]

                existe_pendente = df_pendentes[
                    (df_pendentes["Cidade"].astype(str).str.upper() == cidade_formatada.upper()) &
                    (df_pendentes["UF"].astype(
                        str).str.upper() == uf_formatada.upper())
                ]

                if not existe_total.empty:
                    st.warning(
                        "Essa localidade já existe em Localidades Total.")
                elif not existe_pendente.empty:
                    st.warning(
                        "Essa localidade já existe em Localidades Pendentes.")
                else:
                    nova_linha = {
                        "UF": uf_formatada,
                        "Cidade": cidade_formatada,
                        "Data da Solicitação": nova_data.strftime("%d/%m/%Y"),
                        "Previsão": nova_previsao.strftime("%d/%m/%Y"),
                        "Prioridade": nova_prioridade,
                        "Relatório Detalhado": novo_relatorio.strip()
                    }

                    for coluna in df_total.columns:
                        if coluna not in nova_linha:
                            nova_linha[coluna] = ""

                    for coluna in df_pendentes.columns:
                        if coluna not in nova_linha:
                            nova_linha[coluna] = ""

                    df_total = pd.concat(
                        [df_total, pd.DataFrame([nova_linha])[
                            df_total.columns]],
                        ignore_index=True
                    )

                    df_pendentes = pd.concat(
                        [df_pendentes, pd.DataFrame([nova_linha])[
                            df_pendentes.columns]],
                        ignore_index=True
                    )

                    salvar_excel(df_total, df_pendentes)

                    st.success("Localidade adicionada com sucesso.")
                    st.rerun()

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
            title="Status Geral das Localidades",
            color="Status",
            color_discrete_map={
                "Concluídas": "#00A6D6",
                "Pendentes": "#003B71"
            }
        )

        fig_status.update_traces(
            textposition="inside",
            textinfo="percent+label"
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
            title="Pendências por Estado",
            color="Quantidade",
            color_continuous_scale="Blues"
        )

        fig_uf.update_layout(
            xaxis_title="Estado",
            yaxis_title="Quantidade de Pendências",
            showlegend=False
        )

        st.plotly_chart(fig_uf, use_container_width=True)

    st.divider()

    st.subheader("Resumo das Localidades Pendentes")
    st.caption(
        "Consulte, filtre, altere a prioridade e marque localidades como concluídas diretamente pelo Dashboard."
    )

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
    tabela["__linha_original"] = tabela.index
    tabela["Concluir"] = False

    colunas_editor = [
        col for col in [
            "UF",
            "Cidade",
            "Data da Solicitação",
            "Previsão",
            "Prioridade",
            "Concluir",
            "__linha_original"
        ]
        if col in tabela.columns
    ]

    tabela = tabela[colunas_editor]

    if busca:
        tabela = tabela[
            tabela["Cidade"].str.contains(busca, case=False, na=False)
        ]

    if prioridade != "Todas":
        tabela = tabela[tabela["Prioridade"] == prioridade]

    if uf != "Todas":
        tabela = tabela[tabela["UF"] == uf]

    tabela_editada = st.data_editor(
        tabela,
        use_container_width=True,
        hide_index=True,
        key="editor_pendentes_dashboard",
        disabled=[
            col for col in tabela.columns
            if col not in ["Prioridade", "Concluir"]
        ],
        column_config={
            "Prioridade": st.column_config.SelectboxColumn(
                "Prioridade",
                options=["Alta", "Média", "Baixa"],
                required=True
            ),
            "Concluir": st.column_config.CheckboxColumn(
                "Concluir",
                help="Marque para concluir a localidade",
                default=False
            ),
            "__linha_original": None
        }
    )

    if st.button("Salvar alterações de pendentes"):
        indices_concluir = []

        for _, linha in tabela_editada.iterrows():
            indice_original = int(linha["__linha_original"])
            nova_prioridade = str(linha["Prioridade"]).strip()

            if nova_prioridade in ["Alta", "Média", "Baixa"]:
                df_pendentes.loc[
                    indice_original,
                    "Prioridade"
                ] = nova_prioridade

            if bool(linha["Concluir"]):
                indices_concluir.append(indice_original)

        if indices_concluir:
            df_pendentes = df_pendentes.drop(
                index=indices_concluir
            ).reset_index(drop=True)

        salvar_excel(df_total, df_pendentes)

        st.success("Alterações salvas com sucesso.")
        st.rerun()


with aba_relatorio:
    st.subheader("Relatório Detalhado das Localidades")
    st.caption(
        "Consulte, filtre e edite o texto do relatório detalhado diretamente pelos cards."
    )

    if "Relatório Detalhado" not in df_pendentes.columns:
        st.warning(
            "A coluna 'Relatório Detalhado' não foi encontrada na planilha."
        )
    else:
        col_rel1, col_rel2, col_rel3 = st.columns(3)

        with col_rel1:
            busca_relatorio = st.text_input(
                "Pesquisar localidade no relatório"
            )

        with col_rel2:
            uf_relatorio = st.selectbox(
                "Filtrar UF no relatório",
                ["Todas"] + sorted(
                    df_pendentes["UF"].dropna().unique().tolist()
                )
            )

        with col_rel3:
            prioridade_relatorio = st.selectbox(
                "Filtrar prioridade no relatório",
                ["Todas", "Alta", "Média", "Baixa"]
            )

        tabela_relatorio = df_pendentes.copy()
        tabela_relatorio["__linha_original"] = tabela_relatorio.index

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

        relatorios_editados = {}

        for _, linha in tabela_relatorio.iterrows():
            indice_original = int(linha["__linha_original"])
            cidade = linha.get("Cidade", "")
            uf = linha.get("UF", "")
            prioridade = linha.get("Prioridade", "")
            data = linha.get("Data da Solicitação", "")
            previsao = linha.get("Previsão", "")
            relatorio = linha.get("Relatório Detalhado", "")

            if pd.isna(previsao) or previsao == "None" or previsao == "NaT":
                previsao = "-"

            if pd.isna(relatorio):
                relatorio = ""

            with st.container(border=True):
                st.markdown(f"### {cidade} - {uf}")
                st.markdown(f"**Prioridade:** {prioridade}")
                st.markdown(f"**Data da Solicitação:** {data}")
                st.markdown(f"**Previsão:** {previsao}")
                st.markdown("**Situação:**")

                novo_relatorio = st.text_area(
                    label="Editar relatório",
                    value=str(relatorio),
                    key=f"relatorio_editavel_{indice_original}",
                    height=120,
                    label_visibility="collapsed"
                )

                relatorios_editados[indice_original] = novo_relatorio

        if st.button("Salvar alterações do relatório"):
            for indice_original, novo_relatorio in relatorios_editados.items():
                df_pendentes.loc[
                    indice_original,
                    "Relatório Detalhado"
                ] = str(novo_relatorio).strip()

            salvar_excel(df_total, df_pendentes)

            st.success("Relatório detalhado atualizado com sucesso.")
            st.rerun()


with aba_total:
    st.subheader("Localidades Total")
    st.caption(
        "Lista com todas as localidades abertas. Marque uma localidade como Excluir apenas se ela foi cadastrada por engano."
    )

    filtro_total1, filtro_total2 = st.columns(2)

    with filtro_total1:
        busca_total = st.text_input("Pesquisar cidade no total")

    with filtro_total2:
        uf_total = st.selectbox(
            "Filtrar UF no total",
            ["Todas"] + sorted(df_total["UF"].dropna().unique().tolist())
        )

    tabela_total = df_total.copy()
    tabela_total["__linha_original"] = tabela_total.index
    tabela_total["Excluir"] = False

    colunas_total_editor = [
        col for col in [
            "UF",
            "Cidade",
            "Data da Solicitação",
            "Previsão",
            "Prioridade",
            "Relatório Detalhado",
            "Excluir",
            "__linha_original"
        ]
        if col in tabela_total.columns
    ]

    tabela_total = tabela_total[colunas_total_editor]

    if busca_total:
        tabela_total = tabela_total[
            tabela_total["Cidade"].astype(str).str.contains(
                busca_total,
                case=False,
                na=False
            )
        ]

    if uf_total != "Todas":
        tabela_total = tabela_total[tabela_total["UF"] == uf_total]

    tabela_total_editada = st.data_editor(
        tabela_total,
        use_container_width=True,
        hide_index=True,
        key="editor_localidades_total",
        disabled=[
            col for col in tabela_total.columns
            if col not in ["Excluir"]
        ],
        column_config={
            "Excluir": st.column_config.CheckboxColumn(
                "Excluir",
                help="Marque somente se a localidade foi cadastrada por engano",
                default=False
            ),
            "__linha_original": None
        }
    )

    confirmar_exclusao_total = st.checkbox(
        "Confirmo que desejo excluir as localidades marcadas do Total e dos Pendentes"
    )

    if st.button("Excluir localidades marcadas"):
        if not confirmar_exclusao_total:
            st.warning("Marque a confirmação antes de excluir.")
        else:
            linhas_para_excluir = tabela_total_editada[
                tabela_total_editada["Excluir"] == True
            ]

            if linhas_para_excluir.empty:
                st.warning("Nenhuma localidade foi marcada para exclusão.")
            else:
                indices_total_excluir = linhas_para_excluir[
                    "__linha_original"
                ].astype(int).tolist()

                localidades_excluir = df_total.loc[
                    indices_total_excluir,
                    ["UF", "Cidade"]
                ].copy()

                df_total = df_total.drop(
                    index=indices_total_excluir
                ).reset_index(drop=True)

                for _, localidade in localidades_excluir.iterrows():
                    uf_excluir = str(localidade["UF"]).upper().strip()
                    cidade_excluir = str(localidade["Cidade"]).upper().strip()

                    df_pendentes = df_pendentes[
                        ~(
                            (df_pendentes["UF"].astype(str).str.upper().str.strip() == uf_excluir) &
                            (df_pendentes["Cidade"].astype(
                                str).str.upper().str.strip() == cidade_excluir)
                        )
                    ].reset_index(drop=True)

                salvar_excel(df_total, df_pendentes)

                st.success("Localidade(s) excluída(s) com sucesso.")
                st.rerun()


with aba_concluidas:
    st.subheader("Localidades Concluídas")
    st.caption(
        "As localidades aparecem aqui automaticamente quando são removidas da lista de pendentes."
    )

    st.dataframe(
        df_concluidas,
        use_container_width=True,
        hide_index=True
    )


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
        label="Baixar relatório detalhado em Excel",
        data=gerar_excel_relatorio(df_pendentes),
        file_name="relatorio_detalhado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

import requests
import pandas as pd
import locale
from dateutil import parser
from dash import Dash, html, dcc, dash_table
from dash.dependencies import Input, Output
import dash_bootstrap_components as dbc
import plotly.express as px
import io
import base64
import socket
import threading
import tkinter as tk
from tkinter import ttk
import webbrowser
import time

################### PRIMEIRA PARTE APENAS PARA ABRIR UMA JANELA QUANDO FOR EXECUTÁVEL, FIZ ISSO PENSANDO EM FUTURAMENTE TORNAR UM PROGRAMA EXECUTÁVEL ##########
def get_local_ip():
    hostname = socket.gethostname()
    local_ip = socket.gethostbyname(hostname)
    return local_ip

def iniciar_dashboard():
    app.run(host="0.0.0.0", port=8050, debug=False)

def abrir_janela():
    root = tk.Tk()
    root.title("Dashboard - Inicializando")
    root.geometry("400x150")

    label = tk.Label(root, text="Carregando dashboard...", font=("Arial", 12))
    label.pack(pady=10)

    progress = ttk.Progressbar(root, mode="indeterminate", length=300)
    progress.pack(pady=10)
    progress.start(10)

    ip = get_local_ip()
    url = f"http://{ip}:8050"

    def abrir_navegador():
        webbrowser.open(url)

    link_label = tk.Label(root, text=f"Acesse em: {url}", fg="blue", cursor="hand2")
    link_label.pack(pady=10)
    link_label.bind("<Button-1>", lambda e: abrir_navegador())

    # Iniciar o servidor em outra thread
    threading.Thread(target=iniciar_dashboard, daemon=True).start()

    # Fechar a janela de loading após alguns segundos
    def fechar_loading():
        progress.stop()
        root.title("Dashboard em execução")
        label.config(text="Dashboard iniciado com sucesso!")

    root.after(5000, fechar_loading)  # 5s de espera
    root.mainloop()
    
########INÍCIO DO CODIGO##########
# Configurar idioma para português
locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")

# =========================
# API 1 - CONTRATOS (com paginação)
# =========================
url_contratos = "https://dadosabertos.compras.gov.br/modulo-contratacoes/1_consultarContratacoes_PNCP_14133"
params_contratos = {
    "pagina": 1,
    "tamanhoPagina": 500,
    "unidadeOrgaoCodigoUnidade": "785000",
    "dataPublicacaoPncpInicial": "2025-01-01",
    "dataPublicacaoPncpFinal": "2025-12-31",
    "codigoModalidade": 6,
}
# A lista que vai dar origem ao dataframe:

contratos_list = []

response1 = requests.get(url_contratos, params=params_contratos)

if response1.status_code == 200:
    data = response1.json()
    resultados = data.get("resultado", [])

    for contrato in resultados:
        data_iso = contrato.get("dataPublicacaoPncp", None)
        if data_iso:
            try:
                data_obj = parser.isoparse(data_iso)
                data_formatada = data_obj
            except:
                data_formatada = "N/A"
        else:
            data_formatada = "N/A"

        contratos_list.append(
            {
                "Número da Compra": contrato.get("numeroCompra", "N/A"),
                "Objeto": contrato.get("objetoCompra", "N/A"),
                "Processo NUP": contrato.get("processo", "N/A"),
                "Unidade Gestora": contrato.get("unidadeOrgaoCodigoUnidade", "N/A"),
                "Nome da Unidade Gestora": contrato.get("unidadeOrgaoNomeUnidade", "N/A"),
                "Data Publicação PNCP": data_formatada,
                "Valor Total Estimado": contrato.get("valorTotalEstimado", "N/A"),
                "Valor Total Homologado": contrato.get("valorTotalHomologado", "N/A"),
            }
        )

tabela_contratos = pd.DataFrame(contratos_list)

# 🔹 Converter colunas numéricas
tabela_contratos["Valor Total Estimado"] = pd.to_numeric(
    tabela_contratos["Valor Total Estimado"], errors="coerce"
)
tabela_contratos["Valor Total Homologado"] = pd.to_numeric(
    tabela_contratos["Valor Total Homologado"], errors="coerce"
)

# 🔹 Criar colunas de diferença e desconto
tabela_contratos["Diferença Nominal"] = (
    tabela_contratos["Valor Total Estimado"] - tabela_contratos["Valor Total Homologado"]
)

tabela_contratos["% Desconto"] = (
    tabela_contratos["Diferença Nominal"] / tabela_contratos["Valor Total Estimado"] * 100
)

# Arredondar e formatar com % (ex: 44,92 %)
tabela_contratos["% Desconto"] = tabela_contratos["% Desconto"].round(2).astype(str) + " %"

tabela_contratos_1 = tabela_contratos

# 🔹 Criar um novo dataframe com agregação por mês
tabela_contratos_mes = (tabela_contratos_1.groupby(tabela_contratos_1["Data Publicação PNCP"].dt.to_period("M"))["Valor Total Homologado"].sum().reset_index())

# 🔹 Renomear as colunas
tabela_contratos_mes.columns = ["AnoMes", "Valor Total Homologado"]

# 🔹 Converter AnoMes para string no formato Mês/Ano
tabela_contratos_mes["AnoMes"] = tabela_contratos_mes["AnoMes"].dt.strftime("%b/%Y")


tabela_contratos["Data Publicação PNCP"] = tabela_contratos["Data Publicação PNCP"].dt.date

# Ordenar pela data (mais recentes primeiro)
tabela_contratos = tabela_contratos.sort_values(by="Data Publicação PNCP", ascending=False).reset_index(drop=True)

# =========================
# API 2 - ITENS CONTRATADOS
# =========================
url_itens = "https://dadosabertos.compras.gov.br/modulo-contratacoes/2_consultarItensContratacoes_PNCP_14133"
params_itens = {
    "pagina": 1,
    "tamanhoPagina": 500,
    "unidadeOrgaoCodigoUnidade": "785000",
    "dataInclusaoPncpInicial": "2025-01-01",
    "dataInclusaoPncpFinal": "2025-12-31",
    "codigoModalidade": 6,
}

itens_list = []

response2 = requests.get(url_itens, params=params_itens)

if response2.status_code == 200:
    data = response2.json()
    resultados = data.get("resultado", [])

    for contrato in resultados:
        data_iso = contrato.get("dataInclusaoPncp", None)

        if data_iso:
            try:
                data_obj = parser.isoparse(data_iso)
                data_formatada = data_obj.strftime("%d/%m/%Y - %A")
            except:
                data_formatada = "N/A"
        else:
            data_formatada = "N/A"

        itens_list.append(
            {
                "Id da Compra": contrato.get("numeroControlePNCPCompra", "N/A"),
                "Data Publicação PNCP": data_formatada,
                "Número do Item": contrato.get("numeroItemCompra", "N/A"),
                "Status do item": contrato.get("situacaoCompraItemNome", "N/A"),
                "CATMAT/CATSER": str(contrato.get("codItemCatalogo", "N/A")),
                "Descrição Resumida": contrato.get("descricaoResumida", "N/A"),
                "Descrição Detalhada": contrato.get("descricaodetalhada", "N/A"),
                "Quantidade": contrato.get("quantidade", "N/A"),
                "Valor Unitário Estimado": contrato.get("valorUnitarioEstimado", "N/A"),
                "Valor Total Estimado": contrato.get("valorTotal", "N/A"),
                "Valor Unitário Final": contrato.get("valorUnitarioResultado", "N/A"),
                "Valor Total Final": contrato.get("valorTotalResultado", "N/A"),
                "Nome do Vencedor": contrato.get("nomeFornecedor", "N/A"),
                "CNPJ do Vencedor": contrato.get("codFornecedor", "N/A"),
            }
        )


tabela_itens = pd.DataFrame(itens_list)

# Gráfico por CATMAT/CATSER
df_catmat = tabela_itens.groupby("CATMAT/CATSER")["Valor Total Final"].sum().reset_index()

# Indicadores principais
valor_estimado_total = tabela_contratos["Valor Total Estimado"].sum()
valor_homologado_total = tabela_contratos["Valor Total Homologado"].sum()
economia_nominal = valor_estimado_total - valor_homologado_total
economia_percentual = (economia_nominal / valor_estimado_total * 100) if valor_estimado_total else 0


# =========================
# DASHBOARD
# =========================
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP], suppress_callback_exceptions=True)

app.layout = dbc.Container(
    [
        html.H1(
            "📊 Painel de Contratações Públicas (PNCP)", className="text-center my-4"
        ),
        dbc.Row([
    dbc.Col(
        dbc.Card(
            dbc.CardBody([
                html.H4("💰 Economia Nominal", className="card-title"),
                html.H2(
                    f"R$ {economia_nominal:,.2f}"
                    .replace(",", "X").replace(".", ",").replace("X", "."),
                    className="card-text text-success"
                ),
                html.P(
                    "Diferença entre o valor estimado inicialmente, antes da disputa no Compras.gov, e o valor homologado após o pregão.",
                    className="text-muted small mt-2"
                )
            ]),
            className="shadow-sm border-success border-2"
        ),
        md=6
    ),
    dbc.Col(
        dbc.Card(
            dbc.CardBody([
                html.H4("📉 Economia Percentual", className="card-title"),
                html.H2(
                    f"{economia_percentual:.2f} %",
                    className="card-text text-success"
                ),
                html.P(
                    "Percentual de economia obtido em relação ao valor estimado total.",
                    className="text-muted small mt-2"
                )
            ]),
            className="shadow-sm border-success border-2"
        ),
        md=6
    ),
], className="mb-4"),


        dcc.Tabs(
            [
                # ==============================
                # ABA 1 - DISPENSAS ELETRÔNICAS
                # ==============================
                dcc.Tab(
                    label="Dispensas Eletrônicas",
                    children=[
                        html.Br(),
                        dbc.Button(
                            "⬇️ Baixar Tabela em Excel",
                            id="download-btn-contratos",
                            color="primary",
                            className="mb-3",
                        ),
                        dcc.Download(id="download-contratos"),

                        dbc.Card(
                            [
                                dbc.CardHeader("📑 Tabela de Dispensas Eletrônicas"),
                                dbc.CardBody(
                                    dash_table.DataTable(
                                        id="tabela-contratos",
                                        columns=[{"name": i, "id": i} for i in tabela_contratos.columns],
                                        data=tabela_contratos.to_dict("records"),
                                        page_size=10,
                                        style_table={"overflowX": "auto"},
                                        style_cell={"textAlign": "left"},
                                    )
                                ),
                            ],
                            className="mb-4 shadow-sm border rounded p-2",
                        ),

                        dbc.Card(
                            [
                                dbc.CardHeader("💰 Valor Total Homologado por Mês"),
                                dbc.CardBody(
                                    dcc.Graph(id="grafico-contratos-mes",  # <<< id adicionado
                                        figure=px.bar(
                                            tabela_contratos_mes,
                                            x="AnoMes",
                                            y="Valor Total Homologado",
                                        )
                                    )
                                ),
                            ],
                            className="mb-4 shadow-sm border rounded p-2",
                        ),
                        
                    ],
                ),
                # ==============================
                # ABA 2 - ITENS DE CONTRATAÇÕES
                # ==============================
                dcc.Tab(
                    label="Itens das Dispensas Eletrônicas",
                    children=[
                        html.Br(),
                        dbc.Button(
                            "⬇️ Baixar Tabela em Excel",
                            id="download-btn-itens",
                            color="success",
                            className="mb-3",
                        ),
                        dcc.Download(id="download-itens"),

                        dbc.Card(
                            [
                                dbc.CardHeader("📦 Tabela de Dispensas Eletrônicas por Itens"),
                                dbc.CardBody(
                                    dash_table.DataTable(
                                        id="tabela-itens",
                                        columns=[{"name": i, "id": i} for i in tabela_itens.columns],
                                        data=tabela_itens.to_dict("records"),
                                        page_size=10,
                                        style_table={"overflowX": "auto"},
                                        style_cell={"textAlign": "left"},
                                    )
                                ),
                            ],
                           className="mb-4 shadow-sm border rounded p-2",
                        ),

                        dbc.Card(
                            [
                                dbc.CardHeader("📦 Distribuição de Itens por Status"),
                                dbc.CardBody(
                                    dcc.Graph(id="grafico-itens-status",  # <<< id adicionado
                                        figure=px.histogram(
                                            tabela_itens,
                                            x="Status do item",
                                        )
                                    )
                                ),
                            ],
                            className="mb-4 shadow-sm border-0",
                        ),
                        dbc.Card([
                            dbc.CardHeader("💹 Valor Homologado Total por CATMAT/CATSER"),
                                dbc.CardBody(
                                    dcc.Graph(
                                        id="grafico-catmat",
                                        figure= px.bar(
                                        df_catmat, x="CATMAT/CATSER", y="Valor Total Final",))
        ),
    ],
    className="mb-4 shadow-sm border rounded p-2",
),
                    ],
                ),
            ]
        ),
    ],
    fluid=True,
)

# ==============================
# CALLBACKS PARA DOWNLOAD
# ==============================
@app.callback(
    Output("download-contratos", "data"),
    Input("download-btn-contratos", "n_clicks"),
    prevent_initial_call=True,
)
def download_contratos(n_clicks):
    buffer = io.BytesIO()
    tabela_contratos.to_excel(buffer, index=False)
    buffer.seek(0)
    return dcc.send_bytes(buffer.getvalue(), "Contratos.xlsx")


@app.callback(
    Output("download-itens", "data"),
    Input("download-btn-itens", "n_clicks"),
    prevent_initial_call=True,
)
def download_itens(n_clicks):
    buffer = io.BytesIO()
    tabela_itens.to_excel(buffer, index=False)
    buffer.seek(0)
    return dcc.send_bytes(buffer.getvalue(), "Itens.xlsx")


# ==============================
# RUN
# ==============================
if __name__ == "__main__":
    abrir_janela()



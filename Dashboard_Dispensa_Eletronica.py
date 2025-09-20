import requests
import pandas as pd
import locale
from dateutil import parser
from dash import Dash, html, dcc, dash_table
from dash.dependencies import Input, Output
import dash_bootstrap_components as dbc
from datetime import datetime, timedelta
import plotly.express as px
import io
import socket
import threading
import tkinter as tk
from tkinter import ttk
import webbrowser

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

def definir_status(dias):
    if dias <= 90:
        return "🔴"
    elif dias <= 180:
        return "🟡"
    else:
        return "🟢"

# =========================
# Buscar UASGs pelo CNPJ
# =========================
def buscar_uasgs(cnpj="00394502000144"):
    url = "https://dadosabertos.compras.gov.br/modulo-uasg/1_consultarUasg"
    params = {"cnpjCpfOrgao": cnpj, "statusUasg": True}
    r = requests.get(url, params=params)
    if r.status_code == 200:
        data = r.json()
        return data.get("resultado", [])
    return []

# =========================
# JANELA DE SELEÇÃO DE UASG
# =========================
def selecionar_uasg():
    uasgs = buscar_uasgs()
    if not uasgs:
        print("Nenhuma UASG encontrada para este CNPJ.")
        return None, None

    root = tk.Tk()
    root.title("Dashboard - Seleção da UASG")
    root.geometry("600x250")

    tk.Label(root, text="Selecione a UASG:", font=("Arial", 12)).pack(pady=5)

    opcoes = [f"{u['codigoUasg']} - {u['nomeUasg']}" for u in uasgs]
    combo = ttk.Combobox(root, values=opcoes, width=80)
    combo.pack(pady=10)
    combo.current(0)

    selecionada = {"codigo": None, "nome": None}

    def confirmar():
        valor = combo.get()
        if valor:
            codigo = valor.split(" - ")[0]
            nome = valor.split(" - ")[1]
            selecionada["codigo"] = codigo
            selecionada["nome"] = nome
        root.destroy()

    tk.Button(root, text="Confirmar", command=confirmar).pack(pady=10)

    root.mainloop()
    return selecionada["codigo"], selecionada["nome"]

# =========================
# SELECIONAR AQUI ANTES DO DASHBOARD
# =========================
codigo, nome = selecionar_uasg()

# Configurar idioma para português
locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")

# =========================
# API 1 - CONTRATOS (com paginação)
# =========================
url_contratos = "https://dadosabertos.compras.gov.br/modulo-contratacoes/1_consultarContratacoes_PNCP_14133"
params_contratos = {
    "pagina": 1,
    "tamanhoPagina": 500,
    "unidadeOrgaoCodigoUnidade": codigo,
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
# FORMATANDO A DATA:
        data_iso = contrato.get("dataPublicacaoPncp", None)
        if data_iso:
            try:
                data_obj = parser.isoparse(data_iso)
                data_formatada = data_obj
            except:
                data_formatada = "N/A"
        else:
            data_formatada = "N/A"
####
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

if not tabela_contratos.empty:
    # 🔹 Converter colunas numéricas
    if "Valor Total Estimado" in tabela_contratos.columns:
        tabela_contratos["Valor Total Estimado"] = pd.to_numeric(
            tabela_contratos["Valor Total Estimado"], errors="coerce"
        )
    if "Valor Total Homologado" in tabela_contratos.columns:
        tabela_contratos["Valor Total Homologado"] = pd.to_numeric(
            tabela_contratos["Valor Total Homologado"], errors="coerce"
        )

    # Criar colunas derivadas
    tabela_contratos["Diferença Nominal"] = (
        tabela_contratos["Valor Total Estimado"] - tabela_contratos["Valor Total Homologado"]
    )
    tabela_contratos["% Desconto"] = pd.to_numeric(
        tabela_contratos["Diferença Nominal"] / tabela_contratos["Valor Total Estimado"] * 100
    )
    tabela_contratos["% Desconto"] = tabela_contratos["% Desconto"].round(2).astype(str) + " %"

else:
    # Garante que as colunas existam mesmo se não houver dados
    tabela_contratos = pd.DataFrame(columns=[
        "Número da Compra", "Objeto", "Processo NUP",
        "Unidade Gestora", "Nome da Unidade Gestora",
        "Data Publicação PNCP", "Valor Total Estimado",
        "Valor Total Homologado", "Diferença Nominal", "% Desconto"
    ])


# 🔹 Criar colunas de diferença e desconto
tabela_contratos["Diferença Nominal"] = (
    tabela_contratos["Valor Total Estimado"] - tabela_contratos["Valor Total Homologado"]
)

tabela_contratos["% Desconto"] = (tabela_contratos["Diferença Nominal"] / tabela_contratos["Valor Total Estimado"] * 100
)

if not tabela_contratos.empty:
    # 🔹 Garantir que as colunas numéricas estejam corretas
    tabela_contratos["Valor Total Estimado"] = pd.to_numeric(
        tabela_contratos["Valor Total Estimado"], errors="coerce"
    )
    tabela_contratos["Valor Total Homologado"] = pd.to_numeric(
        tabela_contratos["Valor Total Homologado"], errors="coerce"
    )

    # 🔹 Calcular Diferença Nominal
    tabela_contratos["Diferença Nominal"] = (
        tabela_contratos["Valor Total Estimado"] - tabela_contratos["Valor Total Homologado"]
    )

    # 🔹 Calcular % Desconto de forma segura
    tabela_contratos["% Desconto"] = tabela_contratos.apply(
        lambda row: (row["Diferença Nominal"] / row["Valor Total Estimado"] * 100)
        if row["Valor Total Estimado"] not in [0, None, float('nan')] else 0,
        axis=1
    )

    # 🔹 Arredondar e criar coluna formatada
    tabela_contratos["% Desconto"] = tabela_contratos["% Desconto"].round(2)
    tabela_contratos["% Desconto_fmt"] = tabela_contratos["% Desconto"].astype(str) + " %"

    tabela_contratos_1 = tabela_contratos.copy()

    # 🔹 Agregação por mês
    tabela_contratos_mes = (
        tabela_contratos_1.groupby(tabela_contratos_1["Data Publicação PNCP"].dt.to_period("M"))[
            "Valor Total Homologado"
        ].sum().reset_index()
    )
    tabela_contratos_mes.columns = ["AnoMes", "Valor Total Homologado"]
    tabela_contratos_mes["AnoMes"] = tabela_contratos_mes["AnoMes"].dt.strftime("%b/%Y")

    tabela_contratos["Data Publicação PNCP"] = tabela_contratos["Data Publicação PNCP"].dt.date
    tabela_contratos = tabela_contratos.sort_values(
        by="Data Publicação PNCP", ascending=False
    ).reset_index(drop=True)
else:
    tabela_contratos_1 = tabela_contratos.copy()
    tabela_contratos_mes = pd.DataFrame(columns=["AnoMes", "Valor Total Homologado"])
    tabela_contratos["% Desconto"] = pd.Series(dtype=float)
    tabela_contratos["% Desconto_fmt"] = pd.Series(dtype=str)


# =========================
# API 2 - ITENS CONTRATADOS
# =========================
url_itens = "https://dadosabertos.compras.gov.br/modulo-contratacoes/2_consultarItensContratacoes_PNCP_14133"
params_itens = {
    "pagina": 1,
    "tamanhoPagina": 500,
    "unidadeOrgaoCodigoUnidade": codigo,
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

if not tabela_itens.empty:
    df_catmat = tabela_itens.groupby(["CATMAT/CATSER"])["Valor Total Final"].sum().reset_index()
else:
    # Cria dataframe vazio com as colunas esperadas
    df_catmat = pd.DataFrame(columns=["CATMAT/CATSER", "Valor Total Final"])

# Indicadores principais
valor_estimado_total = tabela_contratos["Valor Total Estimado"].sum()
valor_homologado_total = tabela_contratos["Valor Total Homologado"].sum()
economia_nominal = valor_estimado_total - valor_homologado_total
economia_percentual = (economia_nominal / valor_estimado_total * 100) if valor_estimado_total else 0

# =========================
# API 3 - ATAS DE REGISTROS DE PREÇO
# =========================

# data de hoje
hoje = datetime.today().date()

# intervalo de 1 ano
um_ano_atras = hoje - timedelta(days=365)

# formatar para o padrão da API (YYYY-MM-DD)
data_min = um_ano_atras.strftime("%Y-%m-%d")
data_max = hoje.strftime("%Y-%m-%d")

url_atas = "https://dadosabertos.compras.gov.br/modulo-arp/1_consultarARP"

params2 = {
    "pagina": 1,
    "tamanhoPagina": 500,
    "dataVigenciaInicialMin": data_min,
    "dataVigenciaInicialMax": data_max,
    "codigoUnidadeGerenciadora": codigo
}
response2 = requests.get(url_atas, params=params2)
dados2 = response2.json()
resultados = dados2.get("resultado", [])

atas_list = []

def formatar_data(data_str):
    if data_str:
        try:
            data_obj = parser.isoparse(data_str)
            return data_obj.strftime("%d/%m/%Y")
        except Exception:
            return "N/A"
    return "N/A"

for ata in resultados:
    atas_list.append({"Número da Ata": ata.get("numeroAtaRegistroPreco", "N/A"),
                      "Unidade Gerenciadora": ata.get("codigoUnidadeGerenciadora", "N/A"),
                      "Número de Compra": ata.get("numeroCompra", "N/A"),
                      "Ano da Compra": ata.get("anoCompra", "N/A"),
                      "Data da Assinatura": formatar_data(ata.get("dataAssinatura", "N/A")),
                      "Vigência Inicial": formatar_data(ata.get("dataVigenciaInicial", "N/A")),
                      "Vigência Final": formatar_data(ata.get("dataVigenciaFinal", "N/A")),
                      "Valor Total": ata.get("valorTotal", "N/A"),
                      "Objeto": ata.get("objeto", "N/A"),
                      "Número de Controle da Ata": ata.get("numeroControlePncpAta", "N/A"),
                      "Número de Controle PNCP": ata.get("numeroControlePncpCompra", "N/A"),
                      "Id da Compra": ata.get("idCompra", "N/A"),
                      
                      })

df_atas = pd.DataFrame(atas_list)
df_atas_exibido = df_atas

if not df_atas.empty and "Vigência Final" in df_atas.columns:
    # garantir que a coluna esteja em datetime
    df_atas["Vigência Final Date"] = pd.to_datetime(
        df_atas["Vigência Final"], format="%d/%m/%Y", errors="coerce"
    )
    # dias restantes
    hoje = datetime.today()
    df_atas["Dias Restantes"] = (df_atas["Vigência Final Date"] - hoje).dt.days
    # ordenar decrescente
    df_atas_sorted = df_atas.sort_values(by="Vigência Final Date", ascending=True)
else:
    # cria um dataframe vazio com colunas esperadas (evita crash no layout)
    df_atas = pd.DataFrame(columns=[
        "Número da Ata","Unidade Gerenciadora","Número de Compra","Ano da Compra",
        "Data da Assinatura","Vigência Inicial","Vigência Final","Valor Total",
        "Objeto","Número de Controle da Ata","Número de Controle PNCP","Id da Compra",
        "Vigência Final Date","Dias Restantes"
    ])
    df_atas_sorted = df_atas.copy()

# garantir que a coluna esteja em datetime
df_atas["Vigência Final Date"] = pd.to_datetime(df_atas["Vigência Final"], format="%d/%m/%Y", errors="coerce")
# dias restantes
hoje = datetime.today()
df_atas["Dias Restantes"] = (df_atas["Vigência Final Date"] - hoje).dt.days
# ordenar decrescente
df_atas_sorted = df_atas.sort_values(by="Vigência Final Date", ascending=True)

# 🔹 Criar figure de forma segura
if not tabela_itens.empty and "Status do item" in tabela_itens.columns:
    figure_status = px.histogram(
        tabela_itens,
        x="Status do item",
        y="Valor Total Final",
        color="Status do item",
        title="Distribuição por Status do Item",
    )
else:
    # placeholder vazio para não quebrar o Dash
    figure_status = px.histogram(
        pd.DataFrame({"Status do item": [], "Valor Total Final": []}),
        x="Status do item",
        y="Valor Total Final",
        title="Nenhum dado disponível"
    )

# =========================
# DASHBOARD
# =========================
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP], suppress_callback_exceptions=True)

app.layout = dbc.Container(
    [
        html.H1(
            "📊 Painel de Contratações Públicas (PNCP)", className="text-center my-4"
        ),
        # Card da UASG
        dbc.Card(
            dbc.CardBody([
                html.H4("🏛️ Minha UASG", className="card-title"),
                html.H2(f"{codigo} – {nome}", className="card-text text-primary"),
                html.P(
                    "Dashboard criado a partir da requisição da API de dados abertos do Sistema do Governo Federal - Comprasgov.br",
                    className="text-muted small mt-2"
                )
            ]),
            className="shadow-sm border-primary border-2 mb-4"
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
                    """Diferença entre o valor inicialmente estimado antes da realização do pregão no Compras.gov 
                    e o valor homologado ao final do processo. Ressalta-se que este montante refere-se exclusivamente às 
                    compras diretas da Unidade Gerenciadora (UG) e não inclui valores provenientes de licitações""",
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
                    "Percentual de economia obtido em relação a diferença entre o valor Homologado e Valor Estimado.",
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
                                        style_table={
                                    "overflowX": "auto",
                                    "overflowY": "auto",
                                    },
                                style_header={
                                    "whiteSpace": "normal",
                                    "width": "auto",
                                    "fontWeight": "bold",
                                    "textAlign": "center",
                                    },
                                style_data={                                    
                                    "textAlign": "left"
                                    },
                                
                                fixed_rows={'headers': True},  # cabeçalho fixo ao rolar
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
                                       style_table={
                                    "overflowX": "auto",
                                    "overflowY": "auto",
                                    },
                                style_header={
                                    "whiteSpace": "normal",
                                    "width": "auto",
                                    "fontWeight": "bold",
                                    "textAlign": "center",
                                    },
                                style_data={                                    
                                    "textAlign": "left"
                                    },
                                
                                fixed_rows={'headers': True},  # cabeçalho fixo ao rolar
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
                                        figure=figure_status
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
            
# ------------------------------
# ABA 3 - ATAS DE REGISTRO DE PREÇO
# ------------------------------     
     
            dcc.Tab(
                label="Atas de Registro de Preço",
                children=[
                    html.Br(),
                    dbc.Button("⬇️ Baixar Tabela em Excel", id="download-btn-atas", color="warning", className="mb-3"),
                    dcc.Download(id="download-atas"),

                    dbc.Card([
                        dbc.CardHeader("🗂️ Tabela de Atas de Registro de Preço"),
                        dbc.CardBody(
                            dash_table.DataTable(
                                id="tabela-atas",
                                columns=[{"name": i, "id": i} for i in df_atas_exibido.columns],
                                data=df_atas_exibido.to_dict("records"),
                                style_table={
                                    "overflowX": "auto",
                                    "overflowY": "auto",
                                    },
                                style_header={
                                    "whiteSpace": "normal",
                                    "width": "auto",
                                    "fontWeight": "bold",
                                    "textAlign": "center",
                                    },
                                style_data={                                    
                                    "textAlign": "left"
                                    },
                                fixed_rows={'headers': True},  # cabeçalho fixo ao rolar
                            )
                        ),
                    ], className="mb-4 shadow-sm border rounded p-2"),
                    dbc.Card([
                        dbc.CardHeader("⏳ Prazo de Vigência das Atas"),
                        dbc.CardBody([
                            html.Ul([
                            html.Li([
                    html.Span(definir_status(row['Dias Restantes']) + " "),
                    html.Span(f"Ata {row['Número da Ata']} | "),
                    html.Span(f"Compra {row['Número de Compra']} - {row['Ano da Compra']} | "),
                    html.Span(f"{row.get('Objeto', 'Sem descrição')} | "),
                    html.Span(f"Vigência até {row['Vigência Final']} ("),
                    html.Strong(f"{row['Dias Restantes']} dias restantes"),  # <<< negrito aqui
                    html.Span(")")
                ])
                for _, row in df_atas_sorted.iterrows()
            ],
            style={"listStyleType": "none", "paddingLeft": "0", "margin": 0}
        )
    ], style={"maxHeight": "500px", "overflowY": "auto"}),  # <<< Scroll vertical),
], className="mb-4 shadow-sm border rounded p-2")],
            )]
        )
    ],
    fluid=True)

# ==============================
# CALLBACKS PARA DOWNLOAD
# ==============================

# == BOTÃO PARA DOWNLOAD DA PLANILHA DE CONTRATOS ==
@app.callback(
    Output("download-contratos", "data"),
    Input("download-btn-contratos", "n_clicks"),
    prevent_initial_call=True,
)
def download_contratos(n_clicks):
    buffer = io.BytesIO()
    tabela_contratos.to_excel(buffer, index=False)
    buffer.seek(0)
    return dcc.send_bytes(buffer.getvalue(), f"Contratos_{codigo}_{nome}.xlsx")

# == BOTÃO PARA DOWNLOAD DA PLANILHA DE ITENS ==
@app.callback(
    Output("download-itens", "data"),
    Input("download-btn-itens", "n_clicks"),
    prevent_initial_call=True,
)
def download_itens(n_clicks):
    buffer = io.BytesIO()
    tabela_itens.to_excel(buffer, index=False)
    buffer.seek(0)
    return dcc.send_bytes(buffer.getvalue(), f"Itens_{codigo}_{nome}.xlsx")

# == BOTÃO PARA DOWNLOAD DA PLANILHA DE ATAS ==
@app.callback(
    Output("download-atas", "data"),
    Input("download-btn-atas", "n_clicks"),
    prevent_initial_call=True,
)
def download_atas(n_clicks):
    buffer = io.BytesIO()
    df_atas_exibido.to_excel(buffer, index=False)
    buffer.seek(0)
    return dcc.send_bytes(buffer.getvalue(), f"Atas_{codigo}_{nome}.xlsx")

# ==============================
# RUN
# ==============================
if __name__ == "__main__":
    abrir_janela()

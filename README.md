# Dashboard de Contratações Públicas (PNCP)

Este repositório contém um **Dashboard em Python** para visualização e análise de dados de contratações públicas do Sistema Nacional de Compras do Governo Federal (Compras.gov.br). A aplicação permite acompanhar contratos, itens de contratação e atas de registro de preço de uma **UASG** específica, obtidos através das APIs de dados abertos do governo.

---

## Funcionalidades

- Seleção interativa da **UASG** a partir do CNPJ.
- Consulta e exibição de **contratos** com cálculo de:
  - Valor total estimado e homologado.
  - Diferença nominal e percentual de desconto.
  - Agrupamento por mês.
- Consulta e exibição de **itens contratados**, incluindo:
  - Distribuição por status.
  - Valor total por categoria (CATMAT/CATSER).
- Consulta e exibição de **atas de registro de preço**, com:
  - Datas de vigência.
  - Dias restantes.
  - Indicação visual de status (🔴, 🟡, 🟢).
- **Download em Excel** para contratos, itens e atas.
- Interface interativa usando **Dash**, com layout moderno e responsivo.
- Janela de carregamento inicial em **Tkinter** mostrando IP local e status do dashboard.

---

## Tecnologias utilizadas

- Python 3.10+
- [Dash](https://dash.plotly.com/) e [Dash Bootstrap Components](https://dash-bootstrap-components.opensource.faculty.ai/)
- [Plotly Express](https://plotly.com/python/plotly-express/)
- [Tkinter](https://docs.python.org/3/library/tkinter.html) para interface de carregamento
- [Pandas](https://pandas.pydata.org/) para manipulação de dados
- [Requests](https://docs.python-requests.org/) para chamadas de API
- [Dateutil](https://dateutil.readthedocs.io/) para parsing de datas

---

## Como usar

1. Clone este repositório:
   ```bash
   git clone https://github.com/seu-usuario/dashboard-pncp.git
   cd dashboard-pncp

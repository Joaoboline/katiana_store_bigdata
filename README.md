
# Katiana — Excel + Python + pandas + statsmodels + Plotly (MVP)

## Como usar (local)
1. Crie um ambiente e instale as dependências:
   ```bash
   pip install pandas statsmodels plotly xlsxwriter openpyxl kaleido
   ```

2. Coloque seus dados em **/mnt/data/katiana_excel_plotly_mvp/data/vendas_input.xlsx** (aba `Vendas`). Use as colunas:
   `data, sku, produto, categoria, loja, cliente_id, qtd, preco, desconto, custo`.

3. Rode o script:
   ```bash
   python run_pipeline.py --input data/vendas_input.xlsx --output out/katiana_excel_plotly_dashboard.xlsx
   ```

4. Abra o Excel gerado em `out/katiana_excel_plotly_dashboard.xlsx`.
   - Se o pacote **kaleido** estiver instalado, os **gráficos Plotly** serão embutidos como **imagens PNG** no Dashboard.
   - Sempre geramos os **gráficos Plotly em HTML** em `plots/` para interatividade completa no navegador.

## O que o pipeline faz
- Processa a planilha de vendas e calcula métricas (receita, margem, etc.).
- Constrói **modelo estrela** (dimensões + fato).
- Gera **KPIs** e **previsões (SARIMAX)**.
- Cria **gráficos Plotly** (HTML + PNG opcional) e os inclui no **Dashboard** do Excel ou usa **gráficos nativos do Excel** como fallback.

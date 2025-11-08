import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from statsmodels.tsa.statespace.sarimax import SARIMAX
from datetime import datetime
import os

input_path = "data/vendas_input.xlsx"
output_excel = "out/katiana_excel_plotly_dashboard.xlsx"
plots_path = "plots/"

os.makedirs(plots_path, exist_ok=True)
os.makedirs("out", exist_ok=True)

raw = pd.read_excel(input_path)
raw.columns = raw.columns.str.lower().str.strip()


if "qtd" in raw.columns and "preco" in raw.columns:
    raw["valor_total"] = raw["qtd"] * raw["preco"]
else:
    raise KeyError("A planilha precisa ter as colunas 'qtd' e 'preco'.")


if "data" in raw.columns:
    raw["data"] = pd.to_datetime(raw["data"], errors="coerce")
else:
    raise KeyError("A planilha precisa ter a coluna 'data'.")

raw = raw.dropna(subset=["data"])


df_diario = (
    raw.groupby("data", as_index=False)
    .agg({"valor_total": "sum"})
    .sort_values("data")
)


fig_receita = px.line(
    df_diario,
    x="data",
    y="valor_total",
    title="📈 Receita diária - Katiana Store",
    labels={"data": "Data", "valor_total": "Receita Total (R$)"},
)
fig_receita.update_traces(mode="lines+markers")
fig_receita.write_html(os.path.join(plots_path, "plot_serie_forecast.html"))

if "categoria" in raw.columns:
    df_categoria = (
        raw.groupby("categoria", as_index=False)
        .agg({"valor_total": "sum"})
        .sort_values("valor_total", ascending=False)
    )

    fig_cat = px.bar(
        df_categoria,
        x="categoria",
        y="valor_total",
        title="🏷️ Receita por categoria de produto",
        labels={"valor_total": "Receita Total (R$)", "categoria": "Categoria"},
    )
    fig_cat.write_html(os.path.join(plots_path, "plot_receita_categoria.html"))


df_diario = df_diario.set_index("data")
df_diario = df_diario.asfreq("D", fill_value=0)

modelo = SARIMAX(df_diario["valor_total"], order=(1, 1, 1), seasonal_order=(1, 1, 1, 7))
resultado = modelo.fit(disp=False)

forecast = resultado.get_forecast(steps=30)

forecast_df = forecast.conf_int(alpha=0.10)
forecast_df["forecast"] = forecast.predicted_mean
forecast_df.index.name = "data"

forecast_df["lower valor_total"] = forecast_df["lower valor_total"].clip(lower=0)
forecast_df["upper valor_total"] = forecast_df.apply(
    lambda row: min(row["upper valor_total"], row["forecast"] * 1.5),
    axis=1
)
forecast_df["upper valor_total"] = forecast_df[["upper valor_total", "forecast"]].max(axis=1)

fig_forecast = go.Figure()
fig_forecast.add_trace(
    go.Scatter(
        x=df_diario.index,
        y=df_diario["valor_total"],
        mode="lines",
        name="Histórico",
        line=dict(color="royalblue"),
    )
)
fig_forecast.add_trace(
    go.Scatter(
        x=forecast_df.index,
        y=forecast_df["forecast"],
        mode="lines+markers",
        name="Previsão",
        line=dict(color="orange"),
    )
)

fig_forecast.add_trace(
    go.Scatter(
        x=list(forecast_df.index) + list(forecast_df.index[::-1]),
        y=list(forecast_df["upper valor_total"]) + list(forecast_df["lower valor_total"][::-1]),
        fill="toself",
        fillcolor="rgba(255,165,0,0.2)",
        line=dict(color="rgba(255,255,255,0)"),
        hoverinfo="skip",
        showlegend=True,
        name="Intervalo (90%)",
    )
)

fig_forecast.update_layout(
    title="📊 Previsão de Receita - Próximos 30 dias (intervalo corrigido)",
    xaxis_title="Data",
    yaxis_title="Receita Total (R$)",
)
fig_forecast.write_html(os.path.join(plots_path, "plot_previsao_30_dias.html"))

with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
    raw.to_excel(writer, sheet_name="Vendas_Originais", index=False)
    df_diario.reset_index().to_excel(writer, sheet_name="Receita_Diaria", index=False)
    forecast_df.reset_index().to_excel(writer, sheet_name="Previsao_30_Dias", index=False)

print("✅ Pipeline executado com sucesso!")
print(f"📊 Arquivo Excel gerado em: {output_excel}")
print(f"🌐 Gráficos salvos em: {plots_path}")


try:
    import kaleido
except ImportError:
    print("⚠️ Biblioteca 'kaleido' não encontrada. Instalando...")
    os.system("pip install kaleido")


fig_receita.write_image("plots/plot_receita_diaria.png")
if "categoria" in raw.columns:
    fig_cat.write_image("plots/plot_receita_categoria.png")
fig_forecast.write_image("plots/plot_previsao_30_dias.png")


from openpyxl import load_workbook
from openpyxl.drawing.image import Image

wb = load_workbook(output_excel)
ws = wb.create_sheet("Dashboard")

ws["A1"] = "📊 Dashboard Katiana Store"
ws["A1"].font = ws["A1"].font.copy(bold=True, size=16)

img1 = Image("plots/plot_receita_diaria.png")
img2 = Image("plots/plot_previsao_30_dias.png")

ws.add_image(img1, "A3")
ws.add_image(img2, "J3")

if "categoria" in raw.columns:
    img3 = Image("plots/plot_receita_categoria.png")
    ws.add_image(img3, "A25")

import time
time.sleep(1)

wb.save(output_excel)
print("🧩 Gráficos inseridos na aba 'Dashboard' do Excel com sucesso!")

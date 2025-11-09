import os
import pandas as pd
import plotly.graph_objects as go
from statsmodels.tsa.statespace.sarimax import SARIMAX
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import warnings

warnings.filterwarnings("ignore")

input_excel_path = "data/vendas_input.xlsx"
output_excel = "out/katiana_excel_plotly_dashboard.xlsx"
plots_path = "plots"

os.makedirs("out", exist_ok=True)
os.makedirs("plots", exist_ok=True)

raw = pd.read_excel(input_excel_path, sheet_name="Vendas")
raw.columns = [col.strip().lower().replace(" ", "_") for col in raw.columns]

if "preco" in raw.columns:
    raw["valor_total"] = raw["preco"]
else:
    raise ValueError("A coluna 'preco' não foi encontrada na planilha.")

MARGEM_LUCRO = 0.30
raw["Lucro (R$)"] = raw["valor_total"] * MARGEM_LUCRO
print("✅ Colunas 'valor_total' e 'Lucro (R$)' criadas com sucesso.")


raw["data"] = pd.to_datetime(raw["data"], errors="coerce")
raw = raw.dropna(subset=["data"])
raw = raw.sort_values("data")


df_diario = raw.groupby("data")[["valor_total", "Lucro (R$)"]].sum().reset_index()

fig_receita = go.Figure()
fig_receita.add_trace(
    go.Scatter(
        x=df_diario["data"],
        y=df_diario["valor_total"],
        mode="lines+markers",
        name="Receita (R$)",
        line=dict(color="royalblue"),
    )
)
fig_receita.update_layout(
    title="📈 Receita Diária — Katiana Store",
    xaxis_title="Data",
    yaxis_title="Receita Total (R$)",
)
fig_receita.write_html(os.path.join(plots_path, "plot_receita_diaria.html"))

fig_lucro = go.Figure()
fig_lucro.add_trace(
    go.Bar(
        x=df_diario["data"],
        y=df_diario["Lucro (R$)"],
        name="Lucro Diário (R$)",
        marker_color="green",
    )
)
fig_lucro.update_layout(
    title="💵 Lucro Diário — Margem de 30%",
    xaxis_title="Data",
    yaxis_title="Lucro (R$)",
)
fig_lucro.write_html(os.path.join(plots_path, "plot_lucro_diario.html"))

print("📊 Gráficos de receita e lucro gerados com sucesso!")


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

forecast_df = forecast_df.rename(
    columns={
        "lower valor_total": "Limite Inferior (R$)",
        "forecast": "Previsão (R$)",
        "upper valor_total": "Limite Superior (R$)",
    }
)


fig_forecast = go.Figure()
fig_forecast.add_trace(
    go.Scatter(
        x=df_diario.index,
        y=df_diario["valor_total"],
        mode="lines",
        name="Histórico de Receita",
        line=dict(color="royalblue"),
    )
)
fig_forecast.add_trace(
    go.Scatter(
        x=forecast_df.index,
        y=forecast_df["Previsão (R$)"],
        mode="lines+markers",
        name="Previsão de Receita",
        line=dict(color="orange"),
    )
)
fig_forecast.add_trace(
    go.Scatter(
        x=list(forecast_df.index) + list(forecast_df.index[::-1]),
        y=list(forecast_df["Limite Superior (R$)"]) + list(forecast_df["Limite Inferior (R$)"][::-1]),
        fill="toself",
        fillcolor="rgba(255,165,0,0.2)",
        line=dict(color="rgba(255,255,255,0)"),
        hoverinfo="skip",
        showlegend=True,
        name="Intervalo de Confiança (90%)",
    )
)
fig_forecast.update_layout(
    title="📊 Previsão de Receita — Próximos 30 Dias",
    xaxis_title="Data",
    yaxis_title="Receita Total (R$)",
)
fig_forecast.write_html(os.path.join(plots_path, "plot_previsao_30_dias.html"))


with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
    raw.to_excel(writer, sheet_name="Vendas_Originais", index=False)
    df_diario.reset_index().to_excel(writer, sheet_name="Receita_Diaria", index=False)
    forecast_df.reset_index().to_excel(writer, sheet_name="Previsao_30_Dias", index=False)

print("✅ Dados exportados para Excel com sucesso!")


fig_receita.write_image("plots/plot_receita_diaria.png")
fig_lucro.write_image("plots/plot_lucro_diario.png")
fig_forecast.write_image("plots/plot_previsao_30_dias.png")

wb = load_workbook(output_excel)
ws = wb.create_sheet("Dashboard")

ws["A1"] = "📊 Dashboard Katiana Store — Análise de Vendas e Lucro"
ws["A1"].font = ws["A1"].font.copy(bold=True, size=16)

img1 = Image("plots/plot_receita_diaria.png")
img2 = Image("plots/plot_lucro_diario.png")
img3 = Image("plots/plot_previsao_30_dias.png")

ws.add_image(img1, "A3")
ws.add_image(img2, "J3")
ws.add_image(img3, "A25")

wb.save(output_excel)
print("✅ Dashboard finalizado com sucesso!")

print("\n🏁 Pipeline completo executado com êxito!")
print(f"📘 Arquivo Excel: {output_excel}")

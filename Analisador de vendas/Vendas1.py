import pandas as pd
from pathlib import Path

base_dir = Path(__file__).parent

arquivo_entrada = base_dir / "input" / "vendas.xlsx"
arquivo_saida = base_dir / "output" / "relatorio_vendas.xlsx"

df = pd.read_excel(arquivo_entrada)

df["Total"] = df["Quantidade"] * df["Valor_Unitario"]

total_vendido = df["Total"].sum()
total_unidades = df["Quantidade"].sum()
media_por_unidade = total_vendido / total_unidades

resumo_produtos = df.groupby("Produto").agg({
    "Quantidade": "sum",
    "Total": "sum"
})

ranking_faturamento = resumo_produtos.sort_values(by="Total", ascending=False)
ranking_quantidade = resumo_produtos.sort_values(by="Quantidade", ascending=False)

resumo_geral = pd.DataFrame({
    "Métrica": [
        "Faturamento total",
        "Total de unidades vendidas",
        "Média por unidade"
    ],
    "Valor": [
        total_vendido,
        total_unidades,
        media_por_unidade
    ]
})

with pd.ExcelWriter(arquivo_saida) as writer:
    resumo_geral.to_excel(writer, sheet_name="Resumo", index=False)
    ranking_faturamento.to_excel(writer, sheet_name="Ranking_Faturamento")
    ranking_quantidade.to_excel(writer, sheet_name="Ranking_Quantidade")

print("Relatório gerado com sucesso")

# Salvar relatório
with pd.ExcelWriter("relatorio_vendas.xlsx") as writer:
    resumo_geral.to_excel(writer, sheet_name="Resumo", index=False)
    ranking_faturamento.to_excel(writer, sheet_name="Ranking_Faturamento")
    ranking_quantidade.to_excel(writer, sheet_name="Ranking_Quantidade")

print("Relatório gerado com sucesso: relatorio_vendas.xlsx")
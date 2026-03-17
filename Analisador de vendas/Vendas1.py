import pandas as pd

df = pd.read_excel("vendas.xlsx")

# Criar total por linha
df["Total"] = df["Quantidade"] * df["Valor_Unitario"]

# Métricas gerais
total_vendido = df["Total"].sum()
total_unidades = df["Quantidade"].sum()
media_por_linha = df["Total"].mean()
media_por_unidade = total_vendido / total_unidades

# Resumo por produto em valor
resumo_valor = df.groupby("Produto")["Total"].sum().sort_values(ascending=False)

# Resumo por produto em quantidade
resumo_quantidade = df.groupby("Produto")["Quantidade"].sum().sort_values(ascending=False)

# Produto mais vendido em valor
produto_mais_valor = resumo_valor.idxmax()
valor_produto_mais_valor = resumo_valor.max()

# Produto mais vendido em quantidade
produto_mais_quantidade = resumo_quantidade.idxmax()
quantidade_produto_mais_quantidade = resumo_quantidade.max()

print("=== RESUMO DE VENDAS ===")
print(f"Faturamento total: R$ {total_vendido:.2f}")
print(f"Total de unidades vendidas: {total_unidades}")
print(f"Média por linha de venda: R$ {media_por_linha:.2f}")
print(f"Média por unidade vendida: R$ {media_por_unidade:.2f}")

print("\n=== MAIS VENDIDO EM QUANTIDADE ===")
print(f"Produto: {produto_mais_quantidade}")
print(f"Unidades vendidas: {quantidade_produto_mais_quantidade}")

print("\n=== MAIS VENDIDO EM FATURAMENTO ===")
print(f"Produto: {produto_mais_valor}")
print(f"Faturamento: R$ {valor_produto_mais_valor:.2f}")

print("\n=== FATURAMENTO POR PRODUTO ===")
print(resumo_valor)

print("\n=== QUANTIDADE POR PRODUTO ===")
print(resumo_quantidade)
import pandas as pd

# Leitura dos dados do arquivo Excel
dados = pd.read_excel('PlanilhaVendas.xlsx')

# Calculando a receita total para cada venda
dados['ReceitaTotal'] = dados['Quantidade'] * dados['Preço unitário']

# Determinando o produto mais vendido (em termos de quantidade)
produtoquantidades = dados.groupby('ID do produto')['Quantidade'].sum()
produtomaisvendido = produtoquantidades.idxmax()  # ID do produto mais vendido
quantidadevendida = produtoquantidades.max()      # Quantidade vendida

print(f"O produto mais vendido é o ID: {produtomaisvendido}, com {quantidadevendida} unidades.")

# Calculando a receita total por região
receitaporregiao = dados.groupby('Região')['ReceitaTotal'].sum()
print("\nReceita total por região:")
print(receitaporregiao)

# Identificando o dia com maior número de vendas
vendaspordia = dados['Data da venda'].value_counts()
diacommaisvendas = vendaspordia.idxmax()
vendasnodia = vendaspordia.max()

print(f"\nO dia com maior número de vendas foi: {diacommaisvendas}, com {vendasnodia} vendas.")

# Criando o relatório final
relatorio = dados.groupby('ID do produto').agg(
    quantidadetotalvendida=('Quantidade', 'sum'),
    receitatotal=('ReceitaTotal', 'sum')
).reset_index()

# Ordenando o DataFrame pela receita total em ordem decrescente
relatorio = relatorio.sort_values(by='receitatotal', ascending=False)

# Salvando o relatório em um novo arquivo Excel
relatorio.to_excel('relatorio_vendas.xlsx', index=False)

print("\nRelatório salvo como 'relatorio_vendas.xlsx'.")



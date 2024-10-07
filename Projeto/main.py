import pandas as pd

tabela = pd.read_excel("check.xlsx")

nova_tabela = tabela[tabela["Conformidade"] == "Não"].copy()

# Contar o número de linhas que possuem a palavra "Não"
num_nao = nova_tabela.shape[0]

# Calcular as porcentagens
porcentagem_nao = (num_nao / tabela.shape[0]) * 100

# Adicionar linhas à nova_tabela com os resultados
nova_tabela = nova_tabela._append({"Categoria": "Não", "Contagem": num_nao, "Porcentagem": porcentagem_nao}, ignore_index=True)

# Salvar o DataFrame em um arquivo Excel
nova_tabela.to_excel('Não Conformidade.xlsx', index=False)
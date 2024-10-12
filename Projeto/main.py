import pandas as pd
import PySimpleGUI as sg

# Função para criar uma planilha excel com as não conformidades
def criarNaoConformidade(nomeFile, colunaConformidade):
    tabela = pd.read_excel(nomeFile)

    nova_tabela = tabela[tabela[colunaConformidade] == "Não"].copy()

    # Contar o número de linhas que possuem a palavra "Não"
    num_nao = nova_tabela.shape[0]

    # Calcular as porcentagens
    porcentagem_nao = (num_nao / tabela.shape[0]) * 100

    # Adicionar linhas à nova_tabela com os resultados
    nova_tabela = nova_tabela._append({"Categoria": "Não", "Contagem": num_nao + 1, "Porcentagem": porcentagem_nao}, ignore_index=True)

    # Salvar o DataFrame em um arquivo Excel
    nova_tabela.to_excel('Não Conformidade.xlsx', index=False)

# Interface Gráfica
# Layout
sg.theme("Reddit")
layout = [
    [sg.Text("File"), sg.Input(key="arquivo", size=(20, 1)), sg.FileBrowse()],
    [sg.Text("Nome da Coluna Conformidade"), sg.Input(key="coluna", size=(20, 1))],
    [sg.Button("Criar Planilha")]
]

# Janela
janela = sg.Window("Automação de Não Conformidade", layout)

# Ler os eventos
while True:
    eventos, valores = janela.read()
    if eventos == sg.WINDOW_CLOSED:
        break
    if eventos == "Criar Planilha":
        if valores["coluna"]:  # Verifica se uma coluna foi digitada
            criarNaoConformidade(valores["arquivo"], valores["coluna"])
        else:
            sg.popup_error("Por favor, digite o nome da coluna conformidade/selecione o arquivo.")

janela.close()
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os

# Função para preencher zeros à esquerda na primeira coluna
def preencher_zeros(texto):
    # Se o texto tiver mais de 5 caracteres, remova os 2 últimos caracteres
    if len(texto) > 5:
        texto = texto[:-2]
    # Preencha com zeros à esquerda até 11 dígitos
    return texto.zfill(11)

# Função para combinar as planilhas
def combinar_planilhas():
    # Abrir janela de seleção de arquivo para selecionar múltiplos arquivos Excel
    root = tk.Tk()
    root.withdraw()  # Ocultar a janela principal

    # Permitir a seleção de múltiplos arquivos Excel
    nomes_planilhas = filedialog.askopenfilenames(filetypes=[("Arquivos Excel", "*.xlsx")])

    # Criar uma lista para armazenar os DataFrames das planilhas
    planilhas = []

    # Loop sobre cada planilha e carregar seus dados em um DataFrame
    for planilha in nomes_planilhas:
        dados_planilha = pd.read_excel(planilha)  # Carregar a planilha
        if not dados_planilha.empty:
            print(f"Conteúdo da planilha '{planilha}':")
            print(dados_planilha)  # Imprimir o conteúdo da planilha
            # Converter todas as colunas para o tipo de dados str
            dados_planilha = dados_planilha.astype(str)
            # Preencher zeros à esquerda na primeira coluna até 11 dígitos
            dados_planilha.iloc[:, 0] = dados_planilha.iloc[:, 0].apply(preencher_zeros)
            planilhas.append(dados_planilha)  # Adicionar o DataFrame à lista
        else:
            print(f"A planilha '{planilha}' está vazia.")

    # Verificar se há dados para combinar
    if planilhas:
        # Combinar os DataFrames em um único DataFrame
        dados_combinados = pd.concat(planilhas, ignore_index=True)

        # Determinar o nome do arquivo de saída
        diretorio_area_trabalho = os.path.join(os.path.expanduser('~'), 'Desktop')
        nome_arquivo_saida = 'planilha_combinada.xlsx'
        caminho_arquivo_saida = os.path.join(diretorio_area_trabalho, nome_arquivo_saida)

        # Função para verificar se o arquivo de destino já existe na área de trabalho e criar um novo nome se necessário
        contador = 1
        while os.path.exists(caminho_arquivo_saida):
            nome_arquivo_saida = f'planilha_combinada_{contador}.xlsx'
            caminho_arquivo_saida = os.path.join(diretorio_area_trabalho, nome_arquivo_saida)
            contador += 1

        # Salvar os dados combinados em uma nova planilha na área de trabalho
        dados_combinados.to_excel(caminho_arquivo_saida, index=False)

        print("Planilhas combinadas com sucesso!")
        print(f"Planilha combinada salva na área de trabalho como: '{caminho_arquivo_saida}'")
        print("\nConteúdo da planilha combinada:")
        print(dados_combinados)
    else:
        print("Nenhuma planilha válida foi selecionada.")

# Chamar a função para combinar as planilhas
combinar_planilhas()

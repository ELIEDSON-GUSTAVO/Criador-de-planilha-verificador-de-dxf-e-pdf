import os
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Função para selecionar o arquivo Excel
def selecionar_arquivo_excel():
    Tk().withdraw()  # Oculta a janela principal do Tkinter
    caminho_arquivo = askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    return caminho_arquivo

# Nome da coluna no Excel que contém os códigos
coluna_codigos = "CODIGO"  # Substitua pelo nome exato da coluna no Excel

# Caminho da pasta onde os arquivos estão localizados
caminho_da_pasta = r"T:\14 - PDF\Engenharia"  # Use r"..." para lidar com barras invertidas

# Ler os códigos do arquivo Excel
def ler_codigos_do_excel(caminho_arquivo_excel, coluna_codigos):
    df = pd.read_excel(caminho_arquivo_excel)
    print("Colunas disponíveis no arquivo Excel:", df.columns.tolist())  # Lista as colunas
    df.columns = df.columns.str.strip()  # Remove espaços extras
    if coluna_codigos not in df.columns:
        raise KeyError(f"A coluna '{coluna_codigos}' não foi encontrada no arquivo Excel. Verifique os nomes: {df.columns.tolist()}")
    return df[coluna_codigos].dropna().astype(str).tolist()  # Converte para string e remove valores vazios

# Função para verificar a presença dos códigos nos arquivos da pasta
def verificar_codigos(codigos, caminho_da_pasta):
    presentes = []
    ausentes = []

    # Lista apenas arquivos na pasta (sem diretórios) com extensão .pdf
    arquivos_na_pasta = [os.path.splitext(arquivo)[0] for arquivo in os.listdir(caminho_da_pasta) if arquivo.lower().endswith(".pdf")]
    
    for codigo in codigos:
        # Verifica se o código está contido em algum nome de arquivo
        if any(codigo in arquivo for arquivo in arquivos_na_pasta):
            presentes.append(codigo)
        else:
            ausentes.append(codigo)
    
    return presentes, ausentes

# Fluxo principal
caminho_arquivo_excel = selecionar_arquivo_excel()
if not caminho_arquivo_excel:
    print("Nenhum arquivo Excel foi selecionado. Encerrando o programa.")
else:
    try:
        # Carregar os códigos do Excel
        codigos = ler_codigos_do_excel(caminho_arquivo_excel, coluna_codigos)

        # Verificar os códigos
        presentes, ausentes = verificar_codigos(codigos, caminho_da_pasta)

        # Exibir os resultados
        print("\nCódigos presentes na pasta:")
        for codigo in presentes:
            print(f"- {codigo}")

        print("\nCódigos ausentes na pasta:")
        for codigo in ausentes:
            print(f"- {codigo}")

    except Exception as e:
        print(f"Erro ao processar: {e}")

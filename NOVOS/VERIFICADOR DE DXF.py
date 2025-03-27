import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Caminho da pasta onde os arquivos DXF estão localizados
caminho_da_pasta = r"T:\17 - DXF\Engenharia"

# Função para processar a planilha e verificar arquivos DXF
def processar_arquivo_excel(input_file):
    # Carregar a primeira planilha do arquivo
    df = pd.read_excel(input_file, sheet_name=0)

    # Normaliza os nomes das colunas para minúsculas
    df.columns = df.columns.str.lower()

    # Procurar colunas correspondentes ao código, descrição e material
    colunas_cod = ['número da peça', 'codigo', 'part number', 'partnumber', 'código', 'numero da peça']
    colunas_desc = ['descrição', 'description', 'descricao', 'descriçao']
    colunas_mat = ['material']

    coluna_codigo = next((col for col in colunas_cod if col in df.columns), None)
    coluna_descricao = next((col for col in colunas_desc if col in df.columns), None)
    coluna_material = next((col for col in colunas_mat if col in df.columns), None)

    if not coluna_codigo or not coluna_descricao or not coluna_material:
        print("Erro: Não foram encontradas todas as colunas necessárias (código, descrição, material).")
        return [], []

    # Filtrar os códigos que começam com '02.01.01' ou '02.01.02'
    df_filtrado = df[df[coluna_codigo].astype(str).str.startswith(('02.01.01', '02.01.02'))]

    # Remover materiais indesejados
    materiais_excluir = ['01.04.03', '01.04.04', '01.05.06', '01.04.06', '01.02.01']
    df_filtrado = df_filtrado[~df_filtrado[coluna_material].astype(str).str.startswith(tuple(materiais_excluir), na=False)]

    # Extrair os últimos 5 dígitos do código
    df_filtrado['CODIGO'] = df_filtrado[coluna_codigo].astype(str).str[-5:]

    # Converter os códigos para lista
    codigos_filtrados = df_filtrado['CODIGO'].dropna().astype(str).tolist()

    return codigos_filtrados

# Função para verificar a existência dos códigos nos arquivos DXF
def verificar_codigos(codigos, caminho_da_pasta):
    # Obter os nomes dos arquivos DXF na pasta
    arquivos_dxf = [os.path.splitext(arquivo)[0].strip().lower() for arquivo in os.listdir(caminho_da_pasta) if arquivo.lower().endswith(".dxf")]

    # Listar códigos presentes e ausentes
    presentes = {}
    ausentes = []

    for codigo in codigos:
        codigo = codigo.strip().lower()  # Normaliza o código
        # Filtra todos os arquivos que contenham o código no nome
        arquivos_encontrados = [arquivo for arquivo in arquivos_dxf if codigo in arquivo]
        
        if arquivos_encontrados:
            presentes[codigo] = arquivos_encontrados  # Armazena uma lista de arquivos encontrados para o código
        else:
            ausentes.append(codigo)

    return presentes, ausentes

# Função para abrir a interface de seleção
def selecionar_arquivo():
    root = tk.Tk()
    root.withdraw()

    # Selecionar o arquivo de entrada
    input_file = filedialog.askopenfilename(
        title="Selecione a planilha de entrada",
        filetypes=[("Arquivos Excel", "*.xlsx")]
    )
    if not input_file:
        print("Nenhum arquivo selecionado.")
        return

    # Processar os dados
    codigos_filtrados = processar_arquivo_excel(input_file)

    if not codigos_filtrados:
        print("Nenhum código válido encontrado.")
        return

    # Verificar os códigos na pasta DXF
    presentes, ausentes = verificar_codigos(codigos_filtrados, caminho_da_pasta)

    # Exibir resultados
    print("\nCódigos presentes na pasta DXF:")
    for codigo, arquivos in presentes.items():
        print(f"\nCódigo: {codigo}")
        for arquivo in arquivos:
            print(f" - {arquivo}")

    print("\nCódigos ausentes na pasta DXF:")
    for codigo in ausentes:
        print(f"- {codigo}")

# Executar o programa
if __name__ == "__main__":
    selecionar_arquivo()

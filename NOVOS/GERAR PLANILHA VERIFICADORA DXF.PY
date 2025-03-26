import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Função para filtrar, ajustar e gerar a nova planilha com os últimos 5 dígitos
def filtrar_gerar_planilha(input_file, output_file):
    # Carregar a planilha da primeira planilha
    df = pd.read_excel(input_file, sheet_name=0)  # ou use sheet_name="NomeDaPlanilha" se souber o nome da planilha

    # Procurar por colunas com diferentes variações de "Número da peça" ou "CODIGO"
    colunas_possiveis = ['Número da peça', 'CODIGO', 'Código', 'codigo', 'Numero da peça']
    coluna_codigo = None
    for coluna in colunas_possiveis:
        if coluna in df.columns:
            coluna_codigo = coluna
            break
    
    if coluna_codigo is None:
        print("Nenhuma coluna com o nome esperado ('Número da peça', 'part number', 'CODIGO', 'Código') encontrada.")
        return

    # Procurar por colunas com diferentes variações de "Descrição" ou "Description"
    colunas_possiveis_desc = ['Descrição', 'Description', 'Descricao', 'description']
    coluna_descricao = None
    for coluna in colunas_possiveis_desc:
        if coluna in df.columns:
            coluna_descricao = coluna
            break
    
    if coluna_descricao is None:
        print("Nenhuma coluna com o nome esperado ('Descrição', 'Description') encontrada.")
        return

    # Procurar por colunas com diferentes variações de "Material"
    colunas_possiveis_material = ['Material', 'Material', 'Material', 'material']
    coluna_material = None
    for coluna in colunas_possiveis_material:
        if coluna in df.columns:
            coluna_material = coluna
            break
    
    if coluna_material is None:
        print("Nenhuma coluna com o nome esperado ('Material') encontrada.")
        return

    # Filtrar os itens cujo código começa com '02.01.01' ou '02.01.02'
    filtro = df[coluna_codigo].astype(str).str.startswith(('02.01.01', '02.01.02'))
    df_filtrado = df[filtro].copy()  # Garantir que seja uma cópia

    # Excluir os itens cujos materiais começam com '01.04.03', '01.04.04', '01.05.06', ou '01.04.06'
    materiais_excluir = ['01.04.03', '01.04.04', '01.05.06', '01.04.06']
    for material in materiais_excluir:
        df_filtrado = df_filtrado[~df_filtrado[coluna_material].str.startswith(material, na=False)]

    # Criar uma nova coluna contendo os últimos 5 dígitos do código encontrado
    df_filtrado['CODIGO'] = df_filtrado[coluna_codigo].astype(str).str[-5:]

    # Selecionar apenas a coluna 'CODIGO'
    df_exportado = df_filtrado[['CODIGO']]

    # Exportar para uma nova planilha
    df_exportado.to_excel(output_file, index=False)
    print(f"Planilha exportada com sucesso para: {output_file}")

# Função para abrir a interface de seleção
def selecionar_arquivos():
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal

    # Selecionar o arquivo de entrada
    input_file = filedialog.askopenfilename(
        title="Selecione a planilha de entrada",
        filetypes=[("Arquivos Excel", "*.xlsx")]
    )
    if not input_file:
        print("Nenhum arquivo de entrada selecionado.")
        return

    # Selecionar o local para salvar o arquivo de saída
    output_file = filedialog.asksaveasfilename(
        title="Selecione onde salvar a planilha filtrada",
        defaultextension=".xlsx",
        filetypes=[("Arquivos Excel", "*.xlsx")]
    )
    if not output_file:
        print("Nenhum local de salvamento selecionado.")
        return

    # Chamar a função para processar os dados
    filtrar_gerar_planilha(input_file, output_file)

# Executar o programa
if __name__ == "__main__":
    selecionar_arquivos()

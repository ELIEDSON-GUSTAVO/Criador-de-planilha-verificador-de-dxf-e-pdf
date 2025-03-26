import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Função para filtrar, ajustar links e exportar os dados
def filtrar_ajustar_e_exportar(input_file, output_file):
    # Carregar a planilha da primeira planilha
    df = pd.read_excel(input_file, sheet_name=0)  # ou use sheet_name="NomeDaPlanilha" se souber o nome da planilha

    # Garantir que as colunas essenciais existam
    colunas_necessarias = ['Part Number', 'QTY', 'Description', 'Mass', 'Material', 'File Path']
    for coluna in colunas_necessarias:
        if coluna not in df.columns:
            print(f"Coluna '{coluna}' não encontrada no arquivo.")
            return

    # Filtrar os itens cujo Part Number começa com '02.01.01' ou '02.01.02'
    filtro = df['Part Number'].astype(str).str.startswith(('02.01.01', '02.01.02'))
    df_filtrado = df[filtro]

    # Renomear as colunas para o formato desejado
    df_exportado = df_filtrado.rename(columns={
        'Part Number': 'CODIGO',
        'QTY': 'QTD',
        'Description': 'DESCRIÇAO',
        'Mass': 'MASSA',
        'Material': 'MATERIAL',
        'File Path': 'LINK'
    })

    # Função para ajustar o LINK no formato desejado
    def ajustar_link(file_path):
        # Verifica se o caminho começa com "T:\\14 - PDF\\Engenharia"
        if file_path.startswith("T:\\14 - PDF\\Engenharia"):
            # Substitui "14 - PDF" por "17 - DXF" e altera a extensão para .ipt.dxf
            new_link = file_path.replace("T:\\14 - PDF\\Engenharia", "T:\\17 - DXF\\Engenharia")
            new_link = new_link.replace(".idw.pdf", ".ipt.dxf")
            return new_link
        return file_path  # Caso o link não se encaixe no padrão, retorna como está

    # Aplicar a função de ajuste de link
    df_exportado['LINK'] = df_exportado['LINK'].apply(ajustar_link)

    # Selecionar apenas as colunas necessárias
    colunas_exportadas = ['CODIGO', 'QTD', 'DESCRIÇAO', 'MASSA', 'MATERIAL', 'LINK']
    df_exportado = df_exportado[colunas_exportadas]

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
    filtrar_ajustar_e_exportar(input_file, output_file)

# Executar o programa
if __name__ == "__main__":
    selecionar_arquivos()

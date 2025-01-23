import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog
import os

# Função para validar o formato do código
def validate_codigo(value):
    return bool(re.match(r'^\d{2}\.\d{2}\.\d{2}\.\d{10}$', str(value)))

# Função para validar quantidade (QND)
def validate_qnd(value):
    try:
        value = float(value)
        return value if value > 0 else None  # Se a quantidade for válida e maior que 0, retorna, caso contrário retorna None
    except ValueError:
        return None  # Se não for um número válido, retorna None

# Função para processar os dados
def process_data(df):
    # Definindo os títulos e a ordem das colunas para a planilha de exportação
    headers = ['ITENS', 'CODIGO', 'QND', 'DESCRIÇÃO', 'MASS', 'MATERIAL', 'LINK']
    error_log = []  # Lista para armazenar erros encontrados

    # Criar um DataFrame vazio com as colunas na ordem definida
    export_df = pd.DataFrame(columns=headers)

    # Processar cada linha da tabela de importação
    for index, row in df.iterrows():
        # Desconsiderar linhas vazias (verifica se algum campo importante está preenchido)
        if pd.isnull(row['CODIGO']) and pd.isnull(row['DESCRIÇÃO']):
            continue  # Pula para a próxima linha se ambos os campos estiverem vazios

        # Validar o formato do código
        if pd.notnull(row.get('CODIGO')) and not validate_codigo(row['CODIGO']):
            error_log.append(f"Erro no código na linha {index + 2}: '{row['CODIGO']}' não está no formato correto.")

        # Validar a presença de descrição
        if pd.isnull(row.get('DESCRIÇÃO')):
            error_log.append(f"Descrição ausente na linha {index + 2}.")

        # Validar e garantir que a quantidade seja maior que 0
        qnd = validate_qnd(row.get('QND'))
        if qnd is None:
            error_log.append(f"Quantidade inválida ou ausente na linha {index + 2}.")

        if qnd is None:
            continue  # Se a quantidade for inválida, pula esta linha

        # Criar uma nova linha para exportação com dados válidos
        new_row = pd.DataFrame([{
            'ITENS': index + 1,  # Gerar sequencialmente, começando em 1
            'CODIGO': row.get('CODIGO', ''),
            'QND': qnd,  # Usar a quantidade validada
            'DESCRIÇÃO': row.get('DESCRIÇÃO', ''),
            'MASS': 0.1,  # Define "MASS" como 0.1
            'MATERIAL': '',  # "MATERIAL" fica vazio
            'LINK': ''  # "LINK" fica vazio
        }])

        # Usar pd.concat para adicionar a nova linha ao DataFrame de exportação
        export_df = pd.concat([export_df, new_row], ignore_index=True)

    # Exibir os erros encontrados
    if error_log:
        print("Erros encontrados:")
        for error in error_log:
            print(error)

    return export_df, error_log

# Função para salvar o dataframe em Excel
def save_to_excel(df, folder_path, file_name='editado.xlsx'):
    df.to_excel(os.path.join(folder_path, file_name), index=False)

# Função para salvar o dataframe em CSV com ponto e vírgula
def save_to_csv(df, folder_path, file_name='editado.csv'):
    df.to_csv(os.path.join(folder_path, file_name), index=False, sep=';', decimal=',')  # Especificando o separador como ponto e vírgula e separador decimal como vírgula

# Função para carregar o arquivo selecionado
def carregar_arquivo():
    file_path = filedialog.askopenfilename(title="Selecione a planilha", filetypes=[("Arquivos Excel", "*.xlsx")])
    if file_path:
        folder_path = filedialog.askdirectory(title="Selecione o diretório de destino")  # Solicita o diretório de destino
        if folder_path:
            main(file_path, folder_path)  # Chama a função principal com o caminho do arquivo e o diretório selecionado

# Função para salvar ambos os arquivos (Excel e CSV) com o mesmo nome base
def salvar_arquivos(df, folder_path, base_name='EXPORTAÇÃO'):
    if df.empty:
        print("Erro: Nenhuma linha com quantidade válida foi encontrada. O arquivo CSV não será gerado.")
        return  # Não salva os arquivos se o DataFrame estiver vazio

    # Salva os arquivos com o nome base + extensões .xlsx e .csv no diretório selecionado
    save_to_excel(df, folder_path, f"{base_name}.xlsx")
    save_to_csv(df, folder_path, f"{base_name}.csv")
    print(f"Arquivos salvos como: {os.path.join(folder_path, base_name + '.xlsx')} e {os.path.join(folder_path, base_name + '.csv')}")

# Função principal para importar, processar e exportar os dados
def main(input_file, folder_path):
    try:
        df = pd.read_excel(input_file)
        print("Planilha carregada com sucesso.")
    except Exception as e:
        print(f"Erro ao carregar a planilha: {e}")
        return

    # Processar os dados
    export_df, error_log = process_data(df)

    # Mostrar a tabela processada
    print("\nTabela processada:")
    print(export_df)

    # Salvar ambos os arquivos (Excel e CSV) no diretório escolhido
    salvar_arquivos(export_df, folder_path)

# Exemplo de uso com interface gráfica
if __name__ == "__main__":  # Corrigido para __name__ == "__main__"
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    carregar_arquivo()  # Chama a função para carregar o arquivo e escolher a pasta

import pandas as pd
import pdfplumber
import tkinter as tk
from tkinter import filedialog, messagebox
import logging
from fuzzywuzzy import fuzz
import re

# Configuração de logging
logging.basicConfig(filename='processamento.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Caminho do banco de dados
arquivo_excel = r"C:\Users\Eliedson.silva\Desktop\automa - DIREITOS AUTORAIS PELA LICENSE MIT - ELIEDSON\DADOS.xlsx"

def escolher_arquivo_pedido():
    """Abre uma janela para selecionar o arquivo PDF do pedido de venda."""
    arquivo_pdf = filedialog.askopenfilename(title="Selecione o Pedido de Venda", filetypes=[("Arquivos PDF", "*.pdf")])
    if not arquivo_pdf:
        messagebox.showwarning("Aviso", "Nenhum arquivo selecionado!")
        return None
    return arquivo_pdf

def escolher_local_salvar():
    """Abre uma janela para escolher onde salvar o arquivo Excel gerado."""
    arquivo_saida = filedialog.asksaveasfilename(title="Salvar Planilha", defaultextension=".xlsx", filetypes=[("Arquivo Excel", "*.xlsx")])
    if not arquivo_saida:
        messagebox.showwarning("Aviso", "Nenhum local de salvamento selecionado!")
        return None
    return arquivo_saida

def ler_excel(arquivo_excel):
    """Lê o banco de dados Excel com as peças disponíveis e remove linhas vazias."""
    try:
        df = pd.read_excel(arquivo_excel, sheet_name="BANCO DE DADOS", dtype=str)
        df = df.dropna(how='all')  # Remove linhas completamente vazias
        df = df.rename(columns={
            "LOCAL": "Nome do Item",
            "CODIGO": "Código",
            "QT": "Quantidade",
            "UN. MEDIDA": "Unidade",
            "DESCRIÇÃO": "Descrição"
        })
        return df
    except Exception as e:
        logging.error(f"Erro ao ler o arquivo Excel: {e}")
        messagebox.showerror("Erro", f"Erro ao ler o arquivo Excel: {e}")
        return None

def ler_pdf(arquivo_pdf):
    """Extrai texto do PDF e retorna uma lista de linhas."""
    linhas_relevantes = []
    try:
        with pdfplumber.open(arquivo_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    for linha in texto.split("\n"):
                        linha = re.sub(r'\s+', ' ', linha.strip())  # Remove espaços extras
                        linhas_relevantes.append(linha)
        return linhas_relevantes
    except Exception as e:
        logging.error(f"Erro ao ler o arquivo PDF: {e}")
        messagebox.showerror("Erro", f"Erro ao ler o arquivo PDF: {e}")
        return None

def gerar_novo_excel(df_pecas, pedido_venda, arquivo_saida):
    """Gera um novo Excel com os itens encontrados e não encontrados."""
    dados_saida = []
    itens_nao_encontrados = []

    for linha in pedido_venda:
        item_pedido = linha.strip()
        melhor_correspondencia = None
        maior_score = 0

        for _, componente in df_pecas.iterrows():
            nome_item = str(componente['Nome do Item'])
            if isinstance(nome_item, str):
                score = fuzz.partial_ratio(nome_item.lower(), item_pedido.lower())
                if score > maior_score:
                    maior_score = score
                    melhor_correspondencia = componente

        # Garantir que 'melhor_correspondencia' é tratado corretamente
        if melhor_correspondencia is not None:
            nome_item_correspondente = melhor_correspondencia['Nome do Item']
        else:
            nome_item_correspondente = 'Nenhum'
        
        print(f"Item do PDF: {item_pedido} | Melhor Correspondência: {nome_item_correspondente} | Score: {maior_score}")

        if maior_score > 50:  # Ajustado para permitir mais correspondências
            nome_item_correspondente = melhor_correspondencia['Nome do Item']
            componentes_do_item = df_pecas[df_pecas['Nome do Item'] == nome_item_correspondente]
            for _, componente in componentes_do_item.iterrows():
                dados_saida.append({
                    "Código": componente['Código'],
                    "Quantidade": componente.get('Quantidade', 'N/A'),
                    "Unidade": componente.get('Unidade', 'N/A'),
                    "Descrição": componente.get('Descrição', 'N/A'),
                    "Nome do Item": componente['Nome do Item']
                })
        else:
            itens_nao_encontrados.append(item_pedido)

    df_saida = pd.DataFrame(dados_saida).drop_duplicates().sort_values(by="Nome do Item")
    df_nao_encontrados = pd.DataFrame({"Itens Não Encontrados": itens_nao_encontrados})

    try:
        with pd.ExcelWriter(arquivo_saida, engine='xlsxwriter') as writer:
            df_saida.to_excel(writer, sheet_name="Itens Encontrados", index=False)
            if not df_nao_encontrados.empty:
                df_nao_encontrados.to_excel(writer, sheet_name="Itens Não Encontrados", index=False)
        messagebox.showinfo("Sucesso", f"Planilha gerada com sucesso!\nSalva em: {arquivo_saida}")
        logging.info(f"Planilha gerada com sucesso: {arquivo_saida}")
    except Exception as e:
        logging.error(f"Erro ao gerar o arquivo Excel: {e}")
        messagebox.showerror("Erro", f"Erro ao gerar o arquivo Excel: {e}")

def processar_pedido():
    """Executa o fluxo completo."""
    arquivo_pdf = escolher_arquivo_pedido()
    if not arquivo_pdf:
        return
    arquivo_saida = escolher_local_salvar()
    if not arquivo_saida:
        return
    df_pecas = ler_excel(arquivo_excel)
    if df_pecas is None:
        return
    pedido_venda = ler_pdf(arquivo_pdf)
    if pedido_venda is None:
        return
    gerar_novo_excel(df_pecas, pedido_venda, arquivo_saida)

# Criar interface gráfica
root = tk.Tk()
root.title("Gerador de Planilha de Pedido")
root.geometry("400x200")

btn_processar = tk.Button(root, text="Selecionar Pedido e Gerar Planilha", command=processar_pedido, padx=10, pady=5)
btn_processar.pack(pady=20)

root.mainloop()

import unicodedata
import pdfplumber
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import re
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from fuzzywuzzy import fuzz

# Função para normalizar texto (remover acentos e tornar minúsculo)
def normalizar_texto(texto):
    return ''.join((c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')).lower()

# Função para extrair texto de um PDF
def extrair_texto_pdf(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            texto = ""
            for pagina in pdf.pages:
                pagina_texto = pagina.extract_text()
                if pagina_texto:
                    texto += pagina_texto
                tabelas = pagina.extract_tables()
                for tabela in tabelas:
                    for linha in tabela:
                        if linha:
                            texto += " ".join([str(item) if item is not None else '' for item in linha]) + "\n"
        return texto
    except Exception as e:
        print(f"Erro ao abrir o arquivo PDF: {pdf_path}. Detalhes: {str(e)}")
        return ""

# Função para extrair modelo, entre eixo e carroceria
def extrair_dados_caminhao(texto):
    # Regex mais flexível para capturar o modelo
    modelo = re.search(r"Marca/Modelo:\s*(.*?)(?:\n|Ano:|QUANTIDADE:|$)", texto, re.IGNORECASE)
    
    # Regex para capturar o entre eixo
    entre_eixo = re.search(r"Entre\s*eixo\s*\(mm\):\s*([\d.,]+)", texto, re.IGNORECASE)
    
    # Regex para capturar a carroceria
    carroceria = re.search(r"CARROCERIA\s*(.*?)\s*(?:PARA|COM|,|$)", texto, re.IGNORECASE)

    modelo = modelo.group(1).strip() if modelo else "Desconhecido"

    if entre_eixo:
        entre_eixo_valor = entre_eixo.group(1).replace(".", "").replace(",", ".")
        try:
            entre_eixo = float(entre_eixo_valor)
        except ValueError:
            entre_eixo = None
    else:
        entre_eixo = None

    carroceria = carroceria.group(1).strip() if carroceria else "Desconhecida"

    return modelo, entre_eixo, carroceria

# Função para extrair itens do caminhão abaixo de "ITENS DE PROJETOS"
def extrair_itens_estrutura(texto):
    # Regex mais flexível para capturar "ITENS DE PROJETOS" ou "ITENS PROJETOS"
    match = re.search(r"ITENS\s*DE?\s*PROJETOS\s*(.*?)(?:\n\n|\Z)", texto, re.DOTALL | re.IGNORECASE)
    
    # Se não encontrar "ITENS DE PROJETOS", tenta capturar o texto após "CARROCERIA"
    if not match:
        match = re.search(r"CARROCERIA\s*(.*?)(?:\n\n|\Z)", texto, re.DOTALL | re.IGNORECASE)
    
    # Se ainda não encontrar, tenta capturar o texto após "PRODUTO"
    if not match:
        match = re.search(r"PRODUTO\s*(.*?)(?:\n\n|\Z)", texto, re.DOTALL | re.IGNORECASE)
    
    # Se ainda não encontrar, tenta capturar o texto após "DADOS DO CAMINHÃO"
    if not match:
        match = re.search(r"DADOS\s*DO\s*CAMINHÃO\s*(.*?)(?:\n\n|\Z)", texto, re.DOTALL | re.IGNORECASE)
    
    if match:
        estrutura = match.group(1)
        # Remove linhas que contêm "EN" seguido de números ou "PEDIDO REFERENTE"
        estrutura = "\n".join([linha for linha in estrutura.split("\n") if not re.search(r"\bEN\d+\b|PEDIDO REFERENTE", linha)])
        return estrutura.strip()
    
    return ""

# Função para calcular similaridade usando TF-IDF
def calcular_similaridade(texto1, texto2):
    # Verifica se os textos são válidos (não vazios e contêm palavras)
    if not texto1.strip() or not texto2.strip():
        return 0.0  # Retorna 0 se um dos textos for vazio

    vectorizer = TfidfVectorizer()
    try:
        tfidf_matrix = vectorizer.fit_transform([texto1, texto2])
        return cosine_similarity(tfidf_matrix[0:1], tfidf_matrix[1:2])[0][0] * 100
    except ValueError:
        # Caso ocorra um erro (por exemplo, textos sem palavras significativas)
        return 0.0

# Função para comparar PDFs antigos
def comparar_com_pedidos_antigos(novo_pdf, diretorio_pdfs):
    texto_novo_pdf = extrair_texto_pdf(novo_pdf)
    modelo_novo, entre_eixo_novo, carroceria_novo = extrair_dados_caminhao(texto_novo_pdf)
    estrutura_novo = extrair_itens_estrutura(texto_novo_pdf)

    similaridades = []
    modelos_iguais = []
    entre_eixo_iguais = []
    carroceria_iguais = []

    for nome_pdf in os.listdir(diretorio_pdfs):
        caminho_pdf = os.path.join(diretorio_pdfs, nome_pdf)

        if not nome_pdf.endswith('.pdf') or not os.path.isfile(caminho_pdf):
            continue

        texto_pdf_antigo = extrair_texto_pdf(caminho_pdf)
        modelo_antigo, entre_eixo_antigo, carroceria_antigo = extrair_dados_caminhao(texto_pdf_antigo)
        estrutura_antigo = extrair_itens_estrutura(texto_pdf_antigo)

        modelo_novo_normalizado = normalizar_texto(modelo_novo)
        modelo_antigo_normalizado = normalizar_texto(modelo_antigo)

        similaridade_estrutura = calcular_similaridade(estrutura_novo, estrutura_antigo)

        if entre_eixo_novo and entre_eixo_antigo and entre_eixo_novo == entre_eixo_antigo:
            similaridade_estrutura += 10

        similaridades.append((nome_pdf, similaridade_estrutura, modelo_antigo))
        if fuzz.ratio(modelo_novo_normalizado, modelo_antigo_normalizado) > 80:  # Ajustado para permitir comparação mais flexível
            modelos_iguais.append((nome_pdf, similaridade_estrutura))

        if entre_eixo_novo and entre_eixo_antigo and entre_eixo_novo == entre_eixo_antigo:
            entre_eixo_iguais.append((nome_pdf, similaridade_estrutura))

        if carroceria_novo and carroceria_antigo and carroceria_novo == carroceria_antigo:
            carroceria_iguais.append((nome_pdf, similaridade_estrutura))

        print(f"📄 Comparando com o PDF: {nome_pdf}")
        print(f"🔍 Modelo Novo: {modelo_novo} | Modelo Antigo: {modelo_antigo}")
        print(f"🔍 Similaridade Estrutural: {similaridade_estrutura:.2f}%")
        print("-" * 40)

    similaridades.sort(key=lambda x: x[1], reverse=True)
    modelos_iguais.sort(key=lambda x: x[1], reverse=True)
    entre_eixo_iguais.sort(key=lambda x: x[1], reverse=True)
    carroceria_iguais.sort(key=lambda x: x[1], reverse=True)

    resultado = []

    if similaridades:
        resultado.append("\n📋 Os 6 PDFs mais similares são:")
        for i, (nome_pdf, similaridade, _) in enumerate(similaridades[:6], 1):
            resultado.append(f"{i}. {nome_pdf} - Similaridade da estrutura: {similaridade:.2f}%")
    else:
        resultado.append("❌ Nenhum PDF similar encontrado com o novo PDF.")

    if modelos_iguais:
        resultado.append("\n🎨 Os 6 PDFs mais parecidos com o mesmo modelo de caminhão são:")
        for i, (nome_pdf, similaridade) in enumerate(modelos_iguais[:6], 1):
            resultado.append(f"{i}. {nome_pdf} - Similaridade da estrutura: {similaridade:.2f}%")
    else:
        resultado.append("❌ Nenhum PDF encontrado com o mesmo modelo de caminhão.")

    if entre_eixo_iguais:
        resultado.append("\n🔧 Os 6 PDFs mais parecidos com o mesmo entre eixo são:")
        for i, (nome_pdf, similaridade) in enumerate(entre_eixo_iguais[:6], 1):
            resultado.append(f"{i}. {nome_pdf} - Similaridade da estrutura: {similaridade:.2f}%")
    else:
        resultado.append("❌ Nenhum PDF encontrado com o mesmo entre eixo.")

    if carroceria_iguais:
        resultado.append("\n🚚 Os 6 PDFs mais parecidos com a mesma carroceria são:")
        for i, (nome_pdf, similaridade) in enumerate(carroceria_iguais[:6], 1):
            resultado.append(f"{i}. {nome_pdf} - Similaridade da estrutura: {similaridade:.2f}%")
    else:
        resultado.append("❌ Nenhum PDF encontrado com a mesma carroceria.")

    print("\nResultado Final:")
    print("\n".join(resultado))

# Selecionar o novo PDF
def selecionar_novo_pdf():
    Tk().withdraw()
    return askopenfilename(title="Selecione o novo PDF", filetypes=[("PDF files", "*.pdf")])

# Execução
if __name__ == "__main__":
    diretorio_pdfs = r"T:\1 - DOC PROJETOS\PD DE VENDA"

    novo_pdf = selecionar_novo_pdf()

    if novo_pdf:
        comparar_com_pedidos_antigos(novo_pdf, diretorio_pdfs)
    else:
        print("❌ Nenhum arquivo foi selecionado.")
import unicodedata
import pdfplumber
import os
import threading
from tkinter import Tk, Label, Button, Entry, Text, Scrollbar
from tkinter import ttk
from tkinter.filedialog import askopenfilename
import time

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

# Função para buscar uma frase específica nos PDFs
def buscar_frase_em_pdf(pasta_pdfs, frase_busca, progresso_callback):
    resultados = []
    total_pdfs = len([f for f in os.listdir(pasta_pdfs) if f.endswith('.pdf')])

    for idx, nome_pdf in enumerate(os.listdir(pasta_pdfs)):
        caminho_pdf = os.path.join(pasta_pdfs, nome_pdf)
        if not nome_pdf.endswith('.pdf') or not os.path.isfile(caminho_pdf):
            continue

        texto_pdf = extrair_texto_pdf(caminho_pdf)
        texto_normalizado = normalizar_texto(texto_pdf)
        frase_normalizada = normalizar_texto(frase_busca)

        if frase_normalizada in texto_normalizado:
            resultados.append(f"{nome_pdf} - Encontrada a frase '{frase_busca}'")

        # Atualiza a barra de progresso
        progresso_callback(idx + 1, total_pdfs)

    return resultados

# Função que será chamada em uma thread para buscar a frase nos PDFs
def buscar_com_thread(frase_busca, pasta_pdfs, resultado_texto, barra_progresso):
    def progresso_callback(atual, total):
        barra_progresso['value'] = (atual / total) * 100
        window.update_idletasks()  # Atualiza a interface

    resultados = buscar_frase_em_pdf(pasta_pdfs, frase_busca, progresso_callback)

    # Atualiza os resultados na interface após a busca
    if resultados:
        resultado_texto.delete(1.0, "end")  # Limpa os resultados anteriores
        resultado_texto.insert("1.0", "\n".join(resultados) + "\n")
    else:
        resultado_texto.delete(1.0, "end")
        resultado_texto.insert("1.0", f"❌ A frase '{frase_busca}' não foi encontrada em nenhum PDF.\n")

    # Finaliza a barra de progresso
    barra_progresso['value'] = 100  # Completa a barra de progresso

# Função para iniciar a busca em uma nova thread
def iniciar_busca(entrada_frase, buscar_button, resultado_texto, barra_progresso):
    frase = entrada_frase.get()
    if not frase:
        resultado_texto.insert("1.0", "❌ Por favor, insira uma frase para busca.\n")
        return

    # Desabilitar o botão de buscar durante a execução
    buscar_button.config(state="disabled")
    barra_progresso['value'] = 0  # Reseta a barra de progresso

    # Iniciar a busca em uma thread separada
    def thread_func():
        buscar_com_thread(frase, diretorio_pdfs, resultado_texto, barra_progresso)
        # Habilitar o botão novamente após a execução
        buscar_button.config(state="normal")

    threading.Thread(target=thread_func, daemon=True).start()

# Função para abrir uma interface gráfica
def interface_busca_frase():
    global window  # Tornando a janela global para poder ser referenciada na thread

    # Configuração da janela principal
    window = Tk()
    window.title("Buscador de Frase em PDFs")
    window.geometry("600x500")
    
    # Permitindo que a janela se redimensione
    window.grid_rowconfigure(0, weight=0)
    window.grid_rowconfigure(1, weight=1)
    window.grid_rowconfigure(2, weight=0)
    window.grid_rowconfigure(3, weight=1)
    window.grid_columnconfigure(0, weight=1)

    # Configuração do layout
    label = Label(window, text="Insira a frase a ser buscada:")
    label.grid(row=0, column=0, padx=20, pady=5, sticky="w")
    
    entrada_frase = Entry(window, width=40)
    entrada_frase.grid(row=1, column=0, padx=20, pady=5, sticky="ew")

    buscar_button = Button(window, text="Buscar", command=lambda: iniciar_busca(entrada_frase, buscar_button, resultado_texto, barra_progresso))
    buscar_button.grid(row=2, column=0, padx=20, pady=5, sticky="ew")

    # Barra de progresso
    barra_progresso = ttk.Progressbar(window, orient="horizontal", length=400, mode="determinate")
    barra_progresso.grid(row=3, column=0, padx=20, pady=5, sticky="ew")

    # Caixa de texto para mostrar os resultados
    resultado_texto = Text(window, height=20, width=60)
    resultado_texto.grid(row=4, column=0, padx=20, pady=5, sticky="nsew")

    scrollbar = Scrollbar(window, command=resultado_texto.yview)
    scrollbar.grid(row=4, column=1, sticky="ns")
    resultado_texto.config(yscrollcommand=scrollbar.set)

    # Iniciar a interface gráfica
    window.mainloop()

# Pasta onde os PDFs estão localizados
diretorio_pdfs = r"T:\1 - DOC PROJETOS\PD DE VENDA"

# Executando a interface
if __name__ == "__main__":
    interface_busca_frase()

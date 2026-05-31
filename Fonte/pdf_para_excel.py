import os
import re
import sys
import tkinter as tk
from tkinter import messagebox

import pandas as pd
import pdfplumber

# =========================
# 1. CONFIGURAÇÕES INICIAIS
# =========================
# BASE_DIR = pasta onde o script/executável é executado ou a pasta atual ao rodar em modo script
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.getcwd()

# Lista de postos militares usada para identificar nomes no PDF
POSTOS = [
    "Soldado", "Sd",
    "Cabo", "Cb",
    "3º Sgt", "2º Sgt", "1º Sgt", "Sgt",
    "Sub Ten", "Subtenente",
    "2º Ten", "1º Ten", "Ten",
    "Cap", "Capitão",
    "Maj", "Major",
    "Ten Cel", "TCel",
    "Cel", "Coronel",
    "Gen"
]

# Padrão regex para localizar o trecho entre "4)" e "5)"
PDF_PATTERN = r"4\).*?(?=5\))"


# =========================
# 2. FUNÇÕES DE PROCESSAMENTO
# =========================

def localizar_pdf(pdf_file):
    """Retorna o caminho completo do PDF, considerando pasta atual, pasta do script e caminho absoluto."""
    if os.path.isabs(pdf_file) and os.path.isfile(pdf_file):
        return pdf_file

    possiveis = [
        os.path.join(BASE_DIR, pdf_file),
        os.path.join(SCRIPT_DIR, pdf_file),
    ]

    for caminho in possiveis:
        if os.path.isfile(caminho):
            return caminho

    return None


def encontrar_pdfs():
    """Retorna a lista de arquivos PDF na pasta atual e na pasta do script."""
    encontrados = []
    for pasta in [BASE_DIR, SCRIPT_DIR]:
        if not os.path.isdir(pasta):
            continue
        for nome in sorted(os.listdir(pasta)):
            if nome.lower().endswith(".pdf"):
                caminho = os.path.join(pasta, nome)
                if os.path.isfile(caminho) and caminho not in encontrados:
                    encontrados.append(caminho)
    return encontrados


def extrair_texto_pagina(pdf_file, page_number):
    """Abre o PDF e extrai o texto da página informada."""
    caminho = localizar_pdf(pdf_file)
    if caminho is None:
        raise FileNotFoundError(f"Arquivo PDF não encontrado: {pdf_file}")

    with pdfplumber.open(caminho) as pdf:
        if page_number < 1 or page_number > len(pdf.pages):
            raise ValueError(f"Página inválida. O PDF tem {len(pdf.pages)} páginas.")
        texto = pdf.pages[page_number - 1].extract_text() or ""
    return texto


def extrair_data_e_cidade(texto):
    """Extrai a data de regresso, hora e cidade do bloco do PDF."""
    bloco = re.search(PDF_PATTERN, texto, re.DOTALL)
    trecho = bloco.group(0) if bloco else texto

    # Data no formato típico do PDF: 221200MAI026 ou 221200MAI26
    data_match = re.search(
        r"Regressou em\s*([0-9]{2})([0-9]{4})?([A-ZÁÉÍÓÚÂÊÔÃÕÇ]{3,4})([0-9]{2,4})",
        trecho,
        re.IGNORECASE,
    )
    data_regresso = None
    hora_regresso = None
    if data_match:
        dia = data_match.group(1)
        hora_regresso = data_match.group(2)
        mes_texto = data_match.group(3).upper()
        ano = data_match.group(4)

        meses = {
            "JAN": 1,
            "FEV": 2,
            "MAR": 3,
            "ABR": 4,
            "MAI": 5,
            "JUN": 6,
            "JUL": 7,
            "AGO": 8,
            "SET": 9,
            "OUT": 10,
            "NOV": 11,
            "DEZ": 12,
        }

        mes = meses.get(mes_texto[:3])
        ano_int = int(ano)
        if ano_int < 100:
            ano_int += 2000

        if mes:
            data_regresso = f"{dia}/{mes:02d}/{ano_int}"

    cidade = None
    cidade_match = re.search(
        r"da Guarni(?:[çc]ão|cao) de\s+([A-ZÁÉÍÓÚÂÊÔÃÕÇ\- ]+?)(?=,| onde| que|\.|$)",
        trecho,
        re.IGNORECASE,
    )
    if cidade_match:
        cidade = cidade_match.group(1).strip()

    return data_regresso, hora_regresso, cidade


def extrair_nomes(texto):
    """Localiza o trecho entre 4) e 5) e extrai Posto + Nome com regex."""
    bloco = re.search(PDF_PATTERN, texto, re.DOTALL)
    if not bloco:
        return None

    regex_postos = "|".join(re.escape(p) for p in POSTOS)
    padrao_nomes = rf"({regex_postos})\s+([A-ZÁÉÍÓÚÂÊÔÃÕÇ ]+)"
    resultados = re.findall(padrao_nomes, bloco.group(0))

    if not resultados:
        return None

    data_regresso, hora_regresso, cidade = extrair_data_e_cidade(texto)
    dados = []
    for posto, nome in resultados:
        nome_limpo = re.sub(r"\s+", " ", nome).strip()
        linha = {
            "Posto": posto,
            "Nome": nome_limpo,
        }
        if data_regresso:
            linha["Data Regresso"] = data_regresso
        if hora_regresso:
            linha["Hora"] = hora_regresso
        if cidade:
            linha["Cidade"] = cidade
        dados.append(linha)

    return dados


def salvar_excel(dados, nome_arquivo):
    """Cria um arquivo Excel a partir dos dados extraídos."""
    df = pd.DataFrame(dados)
    caminho_saida = os.path.join(BASE_DIR, nome_arquivo)
    if os.path.exists(caminho_saida):
        try:
            os.remove(caminho_saida)
        except PermissionError:
            raise PermissionError(
                f"Não foi possível sobrescrever o arquivo existente: {caminho_saida}. Feche-o e tente novamente."
            )
    df.to_excel(caminho_saida, index=False, engine="openpyxl")
    return caminho_saida


# =========================
# 3. MODO INTERATIVO COM GUI
# =========================

def gerar_excel():
    """Lê o PDF e gera o arquivo Excel quando o botão for clicado."""
    pdf_selecionado = pdf_var.get()
    pagina_texto = pagina_var.get().strip()

    if not pdf_selecionado:
        messagebox.showwarning("Aviso", "Nenhum arquivo PDF encontrado na pasta.")
        return

    if not pagina_texto.isdigit():
        messagebox.showerror("Erro", "Digite um número de página válido.")
        return

    pagina_num = int(pagina_texto)
    caminho_pdf = pdf_map.get(pdf_selecionado, pdf_selecionado)

    try:
        texto = extrair_texto_pagina(caminho_pdf, pagina_num)
        dados = extrair_nomes(texto)

        if dados is None:
            linhas = [linha.strip() for linha in texto.splitlines() if linha.strip()]
            if not linhas:
                raise ValueError("Não foi possível extrair texto dessa página.")
            dados = [{"Texto": linha} for linha in linhas]

        nome_saida = f"resultado_{os.path.splitext(pdf_selecionado)[0]}_p{pagina_num}.xlsx"
        caminho_saida = salvar_excel(dados, nome_saida)
        messagebox.showinfo("Sucesso", f"Arquivo gerado:\n{caminho_saida}")
    except Exception as error:
        messagebox.showerror("Erro", str(error))


def criar_gui():
    """Cria a interface gráfica para escolher o PDF e a página."""
    global pdf_var, pdf_map
    pdfs = encontrar_pdfs()
    pdf_map = {os.path.basename(caminho): caminho for caminho in pdfs}

    root = tk.Tk()
    root.title("PDF para Excel")
    root.resizable(False, False)

    frame = tk.Frame(root, padx=16, pady=16)
    frame.pack()

    tk.Label(frame, text="Escolha o PDF:").grid(row=0, column=0, sticky="w")
    if pdfs:
        primeiro_pdf = os.path.basename(pdfs[0])
        pdf_var = tk.StringVar(value=primeiro_pdf)
        pdf_menu = tk.OptionMenu(frame, pdf_var, *pdf_map.keys())
        pdf_menu.config(width=40)
        pdf_menu.grid(row=0, column=1, pady=4)
    else:
        pdf_var = tk.StringVar(value="")
        tk.Label(frame, text="(Nenhum PDF disponível)", fg="gray").grid(row=0, column=1, pady=4, sticky="w")

    tk.Label(frame, text="Página:").grid(row=1, column=0, sticky="w")
    global pagina_var
    pagina_var = tk.StringVar(value="1")
    tk.Entry(frame, textvariable=pagina_var, width=10).grid(row=1, column=1, sticky="w", pady=4)

    tk.Button(frame, text="Gerar Excel", command=gerar_excel, width=18).grid(row=2, column=0, columnspan=2, pady=12)

    if not pdfs:
        tk.Label(frame, text="Nenhum arquivo PDF encontrado nesta pasta.", fg="red").grid(row=3, column=0, columnspan=2)

    root.mainloop()


# =========================
# 4. MODO LINHA DE COMANDO
# =========================

def executar_cli(pdf_file, page_number):
    """Executa a extração diretamente a partir da linha de comando."""
    texto = extrair_texto_pagina(pdf_file, page_number)
    dados = extrair_nomes(texto)

    if dados is None:
        linhas = [linha.strip() for linha in texto.splitlines() if linha.strip()]
        if not linhas:
            raise ValueError("Não foi possível extrair texto dessa página.")
        dados = [{"Texto": linha} for linha in linhas]

    nome_saida = f"resultado_{os.path.splitext(pdf_file)[0]}_p{page_number}.xlsx"
    caminho_saida = salvar_excel(dados, nome_saida)
    print(f"Arquivo gerado: {caminho_saida}")


# =========================
# 5. ENTRADA PRINCIPAL
# =========================

if __name__ == "__main__":
    if len(sys.argv) >= 3:
        # Exemplo: python pdf_para_excel.py arquivo.pdf 4
        arquivo_pdf = sys.argv[1]
        pagina = int(sys.argv[2])
        executar_cli(arquivo_pdf, pagina)
    else:
        criar_gui()

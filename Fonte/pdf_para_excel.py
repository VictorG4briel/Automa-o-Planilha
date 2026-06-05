import os
import re
import sys
import tkinter as tk
from tkinter import messagebox

import pandas as pd

# =========================
# 1. CONFIGURAÇÕES INICIAIS
# =========================
# BASE_DIR = pasta onde o script/executável é executado ou a pasta atual ao rodar em modo script
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.getcwd()

# Lista de postos militares usada para identificar nomes no arquivo de texto
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

# =========================
# 2. FUNÇÕES DE PROCESSAMENTO
# =========================

def localizar_txt(txt_file):
    """Retorna o caminho completo do arquivo TXT, considerando pasta atual, pasta do script e caminho absoluto."""
    if os.path.isabs(txt_file) and os.path.isfile(txt_file):
        return txt_file

    possiveis = [
        os.path.join(BASE_DIR, txt_file),
        os.path.join(SCRIPT_DIR, txt_file),
    ]

    for caminho in possiveis:
        if os.path.isfile(caminho):
            return caminho

    return None


def encontrar_txts():
    """Retorna a lista de arquivos TXT na pasta atual e na pasta do script."""
    encontrados = []
    for pasta in [BASE_DIR, SCRIPT_DIR]:
        if not os.path.isdir(pasta):
            continue
        for nome in sorted(os.listdir(pasta)):
            if nome.lower().endswith(".txt"):
                caminho = os.path.join(pasta, nome)
                if os.path.isfile(caminho) and caminho not in encontrados:
                    encontrados.append(caminho)
    return encontrados


def extrair_texto_arquivo(txt_file):
    """Lê o arquivo de texto e retorna seu conteúdo."""
    caminho = localizar_txt(txt_file)
    if caminho is None:
        raise FileNotFoundError(f"Arquivo TXT não encontrado: {txt_file}")

    for encoding in ("utf-8", "latin-1", "cp1252"):
        try:
            with open(caminho, "r", encoding=encoding) as f:
                return f.read()
        except UnicodeDecodeError:
            continue

    raise UnicodeDecodeError(
        "Não foi possível ler o arquivo TXT com codificações comuns (utf-8, latin-1, cp1252)."
    )


def extrair_data_e_cidade(texto):
    """Extrai a data de regresso, hora e cidade do texto completo."""
    trecho = texto

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
    """Extrai Posto + Nome do texto completo."""
    trecho = texto

    regex_postos = "|".join(re.escape(p) for p in POSTOS)
    padrao_nomes = rf"({regex_postos})\s+([A-Za-zÁÉÍÓÚÂÊÔÃÕÇáéíóúâêôãõç ]+)"
    resultados = re.findall(padrao_nomes, trecho, re.IGNORECASE)

    if not resultados:
        return None

    data_regresso, hora_regresso, cidade = extrair_data_e_cidade(texto)
    dados = []
    for posto, nome in resultados:
        nome_limpo = re.sub(r"\s+", " ", nome).strip()
        linha = {
            "Posto": posto.strip(),
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
    """Lê o TXT e gera o arquivo Excel quando o botão for clicado."""
    txt_selecionado = txt_var.get()

    if not txt_selecionado:
        messagebox.showwarning("Aviso", "Nenhum arquivo TXT encontrado na pasta.")
        return

    caminho_txt = txt_map.get(txt_selecionado, txt_selecionado)

    try:
        texto = extrair_texto_arquivo(caminho_txt)
        dados = extrair_nomes(texto)

        if dados is None:
            linhas = [linha.strip() for linha in texto.splitlines() if linha.strip()]
            if not linhas:
                raise ValueError("Não foi possível extrair texto do arquivo TXT.")
            dados = [{"Texto": linha} for linha in linhas]

        nome_saida = f"resultado_{os.path.splitext(txt_selecionado)[0]}.xlsx"
        caminho_saida = salvar_excel(dados, nome_saida)
        messagebox.showinfo("Sucesso", f"Arquivo gerado:\n{caminho_saida}")
    except Exception as error:
        messagebox.showerror("Erro", str(error))


def criar_gui():
    """Cria a interface gráfica para escolher o TXT."""
    global txt_var, txt_map
    txts = encontrar_txts()
    txt_map = {os.path.basename(caminho): caminho for caminho in txts}

    root = tk.Tk()
    root.title("TXT para Excel")
    root.resizable(False, False)

    frame = tk.Frame(root, padx=16, pady=16)
    frame.pack()

    tk.Label(frame, text="Escolha o TXT:").grid(row=0, column=0, sticky="w")
    if txts:
        primeiro_txt = os.path.basename(txts[0])
        txt_var = tk.StringVar(value=primeiro_txt)
        txt_menu = tk.OptionMenu(frame, txt_var, *txt_map.keys())
        txt_menu.config(width=40)
        txt_menu.grid(row=0, column=1, pady=4)
    else:
        txt_var = tk.StringVar(value="")
        tk.Label(frame, text="(Nenhum TXT disponível)", fg="gray").grid(row=0, column=1, pady=4, sticky="w")

    tk.Button(frame, text="Gerar Excel", command=gerar_excel, width=18).grid(row=1, column=0, columnspan=2, pady=12)

    if not txts:
        tk.Label(frame, text="Nenhum arquivo TXT encontrado nesta pasta.", fg="red").grid(row=2, column=0, columnspan=2)

    root.mainloop()


if __name__ == "__main__":
    criar_gui()

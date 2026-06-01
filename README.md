possivel alteração para o codigo

import os
import re
import sys
import tkinter as tk
from tkinter import messagebox

import pandas as pd

# =========================
# 1. CONFIGURAÇÕES INICIAIS
# =========================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.getcwd()

# Lista de postos militares compatível com o seu Bloco de Notas
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
    """Retorna o caminho completo do TXT."""
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
    """Retorna a lista de arquivos TXT na pasta."""
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


def ler_texto_arquivo(txt_file):
    """Abre o arquivo TXT tratando codificações comuns do Windows."""
    caminho = localizar_txt(txt_file)
    if caminho is None:
        raise FileNotFoundError(f"Arquivo TXT não encontrado: {txt_file}")

    try:
        with open(caminho, 'r', encoding='utf-8') as f:
            texto = f.read()
    except UnicodeDecodeError:
        with open(caminho, 'r', encoding='ansi') as f:
            texto = f.read()
            
    return texto or ""


def processar_linhas_txt(texto):
    """Varre o texto linha por linha isolando Posto e Nome com base no seu exemplo."""
    # Ordena os postos por tamanho decrescente para evitar que 'Ten' morda 'Ten Cel'
    postos_ordenados = sorted(POSTOS, key=len, reverse=True)
    regex_postos = "|".join(re.escape(p) for p in postos_ordenados)
    
    # Captura o posto no início da linha e todo o resto como nome
    padrao = rf"^({regex_postos})\b\s*(.*)"

    dados = []
    linhas = texto.splitlines()

    for linha in linhas:
        linha_limpa = linha.strip()
        if not linha_limpa:
            continue  # Ignora linhas em branco

        # Remove prefixos comuns de listas como "Lista:" ou "Nomes:" se sobrarem na linha
        linha_limpa = re.sub(r"^(Lista|Nomes):\s*", "", linha_limpa, flags=re.IGNORECASE)

        # Procura correspondência com a lista de postos militares
        match = re.match(padrao, linha_limpa, re.IGNORECASE)

        if match:
            posto_encontrado = match.group(1)
            nome_encontrado = match.group(2).strip()
            
            dados.append({
                "Posto": posto_encontrado,
                "Nome": nome_encontrado
            })
        else:
            # Caso a linha não comece com um posto conhecido (ex: cabeçalhos ou nomes sem posto)
            dados.append({
                "Posto": "Não Identificado",
                "Nome": linha_limpa
            })

    return dados if dados else None


def salvar_excel(dados, nome_arquivo):
    """Cria o arquivo Excel final."""
    df = pd.DataFrame(dados)
    caminho_saida = os.path.join(BASE_DIR, nome_arquivo)
    if os.path.exists(caminho_saida):
        try:
            os.remove(caminho_saida)
        except PermissionError:
            raise PermissionError(
                f"Feche o arquivo existente antes de continuar: {caminho_saida}"
            )
    df.to_excel(caminho_saida, index=False, engine="openpyxl")
    return caminho_saida


# =========================
# 3. INTERFACE GRÁFICA (GUI)
# =========================

def gerar_excel():
    txt_selecionado = txt_var.get()

    if not txt_selecionado:
        messagebox.showwarning("Aviso", "Nenhum arquivo .txt encontrado.")
        return

    caminho_txt = txt_map.get(txt_selecionado, txt_selecionado)

    try:
        texto = ler_texto_arquivo(caminho_txt)
        dados = processar_linhas_txt(texto)

        if dados is None:
            raise ValueError("Nenhum dado pôde ser extraído do arquivo.")

        # Define o nome do arquivo Excel de saída baseado no TXT selecionado
        nome_puro = os.path.splitext(txt_selecionado)[0]
        nome_saida = f"resultado_{nome_puro}.xlsx"
        
        caminho_saida = salvar_excel(dados, nome_saida)
        messagebox.showinfo("Sucesso", f"Excel gerado com sucesso:\n{caminho_saida}")
    except Exception as error:
        messagebox.showerror("Erro", str(error))


def criar_gui():
    global txt_var, txt_map
    txts = encontrar_txts()
    txt_map = {os.path.basename(caminho): caminho for caminho in txts}

    root = tk.Tk()
    root.title("Conversor TXT para Excel")
    root.resizable(False, False)

    frame = tk.Frame(root, padx=16, pady=16)
    frame.pack()

    tk.Label(frame, text="Escolha o arquivo TXT:").grid(row=0, column=0, sticky="w")
    if txts:
        primeiro_txt = os.path.basename(txts[0])
        txt_var = tk.StringVar(value=primeiro_txt)
        txt_menu = tk.OptionMenu(frame, txt_var, *txt_map.keys())
        txt_menu.config(width=40)
        txt_menu.grid(row=0, column=1, pady=4)
    else:
        txt_var = tk.StringVar(value="")
        tk.Label(frame, text="(Nenhum arquivo TXT encontrado)", fg="gray").grid(row=0, column=1, pady=4, sticky="w")

    tk.Button(frame, text="Gerar Excel", command=gerar_excel, width=18).grid(row=1, column=0, columnspan=2, pady=12)

    if not txts:
        tk.Label(frame, text="Insira um arquivo .txt na mesma pasta do script.", fg="red").grid(row=2, column=0, columnspan=2)

    root.mainloop()


if __name__ == "__main__":
    criar_gui()

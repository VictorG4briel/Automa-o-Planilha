import pdfplumber
import re
import pandas as pd

# =========================
# 1. ABRIR O PDF E PEGAR A PÁGINA
# =========================
# pdfplumber.open() abre o arquivo PDF
# pdf.pages[índice] seleciona uma página (começa do 0)
# Se o PDF tem 13 páginas e você quer a página 6, use pages[5]

with pdfplumber.open("2026-05-25_O_098_boletim_interno.pdf") as pdf:
    # Extrair texto da página desejada
    pagina = pdf.pages[3]  # Mude o índice conforme necessário
    texto = pagina.extract_text()

# =========================
# 2. EXTRAIR APENAS O CONTEÚDO ENTRE TÓPICO 4 E 5
# =========================
# O padrão regex busca:
# - "4\)" = o número 4 seguido de parêntesis
# - ".*?" = qualquer coisa 
# - "(?=5\))" = até encontrar "5)" (lookahead positivo - não inclui o 5)
# - re.DOTALL = faz o . capturar quebras de linha também

padrao_entre_topicos = r"4\).*?(?=5\))"

bloco_topico4 = re.search(padrao_entre_topicos, texto, re.DOTALL)

if not bloco_topico4:
    print("ERRO: Tópico 4 não encontrado!")
    exit()

conteudo_topico4 = bloco_topico4.group(0)

# =========================
# 3. EXTRAIR POSTOS E NOMES
# =========================
# Lista de postos militares que existem no PDF
postos = [
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

# Cria o padrão regex juntando todos os postos
# re.escape() protege caracteres especiais como "º"
regex_postos = "|".join(re.escape(p) for p in postos)

# Padrão para capturar POSTO + NOME:
# - ({regex_postos}) = captura qualquer um dos postos
# - \s+ = um ou mais espaços
# - ([A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]+) = captura o nome (letras maiúsculas com acentos + espaços)
padrao_nomes = rf"({regex_postos})\s+([A-ZÁÉÍÓÚÂÊÔÃÕÇ ]+)"

# findall() retorna uma lista de tuplas (posto, nome)
resultados = re.findall(padrao_nomes, conteudo_topico4)

# =========================
# 4. LIMPAR OS DADOS
# =========================
# Remove espaços extras no final do nome
dados = []
for posto, nome in resultados:
    nome_limpo = nome.strip()
    # Remove espaços duplicados no meio do nome
    nome_limpo = re.sub(r"\s+", " ", nome_limpo)
    
    dados.append({
        "Posto": posto,
        "Nome": nome_limpo
    })

# =========================
# 5. SALVAR EM ARQUIVO EXCEL
# =========================
if dados:
    # Criar um DataFrame do pandas
    df = pd.DataFrame(dados)
    
    # Salvar em arquivo Excel
    df.to_excel("resultado.xlsx", index=False, engine="openpyxl")
    
    print(f"✓ Sucesso! {len(dados)} nomes foram salvos em 'resultado.xlsx'")
else:
    print("✗ ERRO: Nenhum nome foi encontrado!")


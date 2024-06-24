#Importanto as bibliotecas do python
import pandas as pd
from docx import Document
import re

# Carregar o arquivo DOCX
doc = Document('./datasets/Partnership.docx')

# Inicializar listas para armazenar os nomes e as cotas dos sócios
nomes = []
cotas = []

# Expressão regular para encontrar o nome de cada sócio
padrao = r"\d+\.\s(.*?),.*?(\d+)\s+co"
# Explicação do padrão:
# \d+   -> Corresponde a um ou mais dígitos (o número do sócio)
# \.    -> Corresponde a um ponto (separador entre o número do sócio e o nome)
# \s    -> Corresponde a um caractere de espaço em branco
# (.*?) -> Grupo de captura não ganancioso para capturar o nome do sócio
# ,     -> Corresponde a uma vírgula (separador entre o nome e o número de cotas)
# .*?   -> Corresponde a zero ou mais caracteres (para ignorar o texto entre o nome e o número de cotas)
# (\d+) -> Grupo de captura para capturar o número de cotas (um ou mais dígitos)
# \s+   -> Corresponde a um ou mais caracteres de espaço em branco
# co    -> Corresponde à sequência "co" (para garantir que estamos capturando o número de cotas)

# Iterar sobre os parágrafos do documento
for paragraph in doc.paragraphs:
    # Procurar padrões no texto do parágrafo
    match = re.search(padrao, paragraph.text)
    if match:
        # Extrair o nome do sócio e suas cotas
        nome = match.group(1)
        cota = int(match.group(2))
        nomes.append(nome)
        cotas.append(cota)
print(nomes)
print(cotas)
# Criar um DataFrame com os dados
dados = {"NOME DO SÓCIO": nomes, "NÚMERO DE COTAS": cotas}
df = pd.DataFrame(dados)

# Exportar o DataFrame para um arquivo Excel
df.to_excel("./outputs/resultado_desafio2.xlsx", index=False)
print("Arquivo Excel 'resultado_desafio2.xlsx' criado com sucesso com os dados dos sócios.")
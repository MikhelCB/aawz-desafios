import pandas as pd  
from docx import Document
import re

documento = Document('datasets/Partnership.docx')

nomes = []
cotas = []

# Expressão regular para encontrar o nome de cada sócio
padrao = r"\d+\.\s(.?),.?(\d+)\s+co"
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

for paragrafo in documento.paragraphs:
    pesquisa = re.search(padrao, paragrafo.text)
    if pesquisa: 
        nome = pesquisa.group(1)
        cota = int(pesquisa.group(2))
        nomes.append(nome)
        cotas.append(cota)
    

print(nomes)
print(cotas)
    
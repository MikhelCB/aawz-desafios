#importando a blibliote openpyxl para exportar arquivos em excel e pandas para leitura e interpretação de datasets
import pandas as pd  
import openpyxl

#foi criado uma variavel vendas, para ler a planilha em excel
caminho_arquivo_vendas = './datasets/Vendas.xlsx'
vendas = pd.read_excel(caminho_arquivo_vendas, sheet_name='Vendas')

#Aqui estamos limpando a formatação, estamos trocando o R$ por 'nada' e trocando a ',' por '.' e retornando um float, porque esta em string
def limpar_formatacao(valor):
  if isinstance(valor, str):
    valor = valor.replace('R$ ', '').replace('.', '').replace(',', '.')
  return float(valor)

#aqui estamos aplicando para a formatação que deverá ser descrita, colocando o ',' no lugar de 'v'e etc...
def aplicar_formatacao(valor): # 1,000.00 => 1v000.00 => 1v000,00 => R$ 1.000,00
   return f'R$ {valor:,.2f}'.replace(',', 'v').replace('.', ',').replace('v', '.')


#Estou subtstituindo os valores da coluna com os mesmos dados, porem com uma função aplicada, no caso = apply(limpar_formatacao)
vendas['Valor da Venda'] = vendas['Valor da Venda'].apply(limpar_formatacao)
vendas['Custo da Venda'] = vendas['Custo da Venda'].apply(limpar_formatacao)

def calcular_comissao(linha_da_tabela):
  comissao_inicial = linha_da_tabela['Valor da Venda'] * 0.10
  #Se o canal de venda for igual a online, 80% vai pro vendedor e 20% vai para o marketing
  if linha_da_tabela['Canal de Venda'] == 'Online':
    comissao_do_vendedor = comissao_inicial * 0.80
    comissao_do_marketing = comissao_inicial * 0.20
    # Se não for 'Online' a comissao do vendedor vai ser cheia (comissa_inicial) e marketing = 0
  else:
    comissao_do_vendedor = comissao_inicial
    comissao_do_marketing = 0.0

  # Se a comissao do vendedor for maior ou igual a 1500, a comissao do vendedor vai ser igual a 90% da comissao porque 10% é do gerente se não o gerente não ganha nada.
  if comissao_do_vendedor >= 1500:
    comissao_do_gerente = comissao_do_vendedor * 0.10
    comissao_do_vendedor = comissao_do_vendedor * 0.90

  else:
    comissao_do_gerente = 0.0
  return comissao_do_vendedor, comissao_do_marketing, comissao_do_gerente

#Estamos aplicando a função de calculo para cada uma das linhas na tabela
# Notei que a condição onde calcula-se o valor da comissão do gerente, nunca será satisfeita, porque para nenhum dos registros o valor da comissão é >= 1500, mesmo que somados.
vendas[['Comissão do Vendedor', 'Comissão do Marketing', 'Comissão do Gerente']] = vendas.apply(calcular_comissao, axis=1, result_type='expand')

vendas["Custo das Vendas"] = vendas["Custo da Venda"]

#Gruoupby agrupar por e AGG = vai agregar colunas ao agrupamento feito com groupby, Reset Index = função que irá resetar a tabela pra que as colunas novas não de bug
comissao_do_vendedor = vendas.groupby("Nome do Vendedor").agg({"Custo das Vendas": "sum", "Comissão do Vendedor": "sum"}).reset_index()

comissao_do_vendedor["Custo das Vendas"] = comissao_do_vendedor["Custo das Vendas"].apply(aplicar_formatacao)
comissao_do_vendedor["Comissão do Vendedor"] = comissao_do_vendedor["Comissão do Vendedor"].apply(aplicar_formatacao)

# ------------------------------- tarefa 2 ------------------------------------

pagamentos = pd.read_excel(caminho_arquivo_vendas, sheet_name='Pagamentos')

#Drop= A tabela pagamentos vai ser igual a tabela pagamentos menos a coluna comissão 
pagamentos["Comissão Paga"] = pagamentos["Comissão"].apply(limpar_formatacao)
pagamentos = pagamentos.drop(columns=["Comissão"])

comissao_do_vendedor["Comissão do Vendedor"] = comissao_do_vendedor["Comissão do Vendedor"].apply(limpar_formatacao)

pagamentos_validos = pd.merge(pagamentos, comissao_do_vendedor, on="Nome do Vendedor", how="left")
pagamentos_validos["Diferença"] = pagamentos_validos["Comissão Paga"] - pagamentos_validos["Comissão do Vendedor"]


pagamentos_invalidos = pagamentos_validos[pagamentos_validos["Diferença"] != 0]
pagamentos_validos = pagamentos_validos[pagamentos_validos["Diferença"] == 0]

comissao_do_vendedor["Comissão do Vendedor"] = comissao_do_vendedor["Comissão do Vendedor"].apply(aplicar_formatacao)
pagamentos_invalidos["Comissão Paga"] = pagamentos_invalidos["Comissão Paga"].apply(aplicar_formatacao)
pagamentos_invalidos["Comissão Merecida"] = pagamentos_invalidos["Comissão do Vendedor"].apply(aplicar_formatacao)
pagamentos_invalidos = pagamentos_invalidos.drop(columns=["Comissão do Vendedor"])
pagamentos_invalidos["Diferença"] = pagamentos_invalidos["Diferença"].apply(aplicar_formatacao)

# EXPORTAÇÃO
excel_writer = pd.ExcelWriter('./outputs/resultado_desafio1.xlsx', engine="openpyxl")

# Salvar os resultados da tarefa 1
comissao_do_vendedor.to_excel(excel_writer, sheet_name="Tarefa 1", index=False)

# Salvar resultados da tarefa 2
pagamentos_invalidos.to_excel(excel_writer, sheet_name="Tarefa 2", index=False)

# Persistir alterações
excel_writer._save()

#Print
print("Arquivo com resultados salvo em outputs/resultado_desafio1.xlsx")
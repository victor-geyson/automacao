#Importa a biblioteca Pandas, necessária para trabalhar com a base de dados.
import pandas as pd

#Aqui entra o caminho da planilha, que pode ser absoluto ou relativo. Ex: Base_De_Dados.xlsx

caminho_planilha = "Insira o caminho da planilha aqui" 

#Aqui entra a aba da planilha que você pretende importar, caso você não especifique ele irá importar a primeira aba

sheet = "Insira o nome da aba que você deseja manipular" 

#Importa a planilha e seleciona a aba
base = pd.read_excel(io = caminho_planilha, sheet) 

#Nesse momento você vai ter um objeto "base" que conterá os dados da aba da planilha que você especificou

#O próximo passo é selecionar o a coluna que você deseja utilizar como referência

'''Aqui você faz referência ao nome da coluna que você quer usar de referência
Exemplo, suponha que você tenha uma base que tenha as colunas "Vendedor", "Produto", "Data" e "Venda"
Caso você queirá gerar uma planilha para cada vendedor então a sua coluna_ref = "Vendedor"
Caso você queirá gerar uma planilha para cada produto então a sua coluba_ref = "Produto" e etc.'''

coluna_ref = "Nome da Coluna"  

#Aqui iremos criar uma lista com os valores únicos de cada 
coluna_ref = list(set(base[ColunaRef].tolist()))

'''Nesse momento temos uma base de dados e uma lista com valores únicos
Agora, precisamos criar um loop para:
1) Filtrar a planilha
2) Gerar o arquivo apenas com os dados filtrados
'''

#Cria o loop
for i in coluna_ref: 
#O i é cada valor único da coluna de referência, exemplo, se há três produtos: Carros, Motos e Caminhões, o i iria assunmir o valor de Carro, depois de Moto e Camimnhão 
  basefiltrada = base[base[coluna_ref] == i] #Filtra a base de dados, restando apenas os valores de um determinado valor único da coluna de referência.
#Exemplo, para i = Carros, ele filtrará todos os valores onde a coluna = Carros
  basefiltrada.to_excel(f"{i}.xlsx") #Gera uma planilha com o nome do valor i e salva.
#Exemplo, para i = "Carros", ele salvará como "Carros".xlsx

#Após isso ele irá gerar uma planilha "Moto.xlsx" e depois "Caminhão.xlsx"

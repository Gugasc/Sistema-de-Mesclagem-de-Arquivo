import pandas as pd
import os  # Manipular arquivos/pastas

print("Iniciando o programa")
# Caminho da pasta com os arquivos
pasta = r"C:\VS CODE\Projeto - 1\data"

# Lista com todos os arquivos dentro da pasta.
# os.path.join: vai juntar todos os caminhos na variável "pasta" para a coleta do arquivo;
# os.listdir: vai listar os arquivos dentro da pasta;
arquivos = [os.path.join(pasta, arquivo) for arquivo in os.listdir(pasta)]

# Criação de tabela vazia
tabelaFinal = pd.DataFrame()

# tabela = pd.read_excel(arquivos[0], index_col="Data") ; o index_col pode ser usado para tornar uma coluna em um indice

vendas1 = pd.read_excel(r"C:\VS CODE\Projeto - 1\data\Vendas - 1.xlsx")
# Loop para passar em todos os arquivos e ir juntando na tabela final
for arquivo in arquivos:
    #Vai ler cada arquivo em excel presente na pasta 
    aux = pd.read_excel(arquivo)
    #columns.tolist(): obtem o nome de cada coluna do Dataframe
    colunas_vendas1 = vendas1.columns.tolist()
    colunas_aux = aux.columns.tolist()
    
    #a função set() é usada para criar conjuntos
    if set(colunas_vendas1) == set(colunas_aux):
        print("É válido para mesclagem")
        df = pd.read_excel(arquivo)
        tabelaFinal = pd.concat([tabelaFinal, df], ignore_index= True)
    else:
        print("INVÁLIDO PARA MESCLAGEM: ", arquivo)

print("MESCLAGEM FEITA")

tabelaFinal.to_excel(r"C:\VS CODE\Projeto - 1\Dados_Finais\Vendas Totais.xlsx")
print("Criação de tabela final, concluida")

gerentes = pd.read_excel(r"C:\VS CODE\Projeto - 1\data\Gerentes.xlsx")

#junta duas tabelas, usando como base os valores semelhantes entre elas (como um procv)
tabelaFinal = pd.merge(tabelaFinal, gerentes)

#vai excluir todas as colunas com valores vazios
tabelaFinal = tabelaFinal.dropna(how = 'all', axis=1)
tabelaFinal.to_excel(r"C:\VS CODE\Projeto - 1\Dados_Finais\Vendas Merge.xlsx")
print("Criação de tabela final com os gerentes, concluida")

#vai realizar atransformação do arquivo para .json
tabelaFinal.to_json("Teste.json")
print("Criação de arquivo .JSON, concluido")

# Criação de variável com um id de loja específico
vendasIguatemiEsp = tabelaFinal.loc[tabelaFinal["ID Loja"] == "Iguatemi Esplanada"]
vendasIguatemiEsp.to_excel("C:\VS CODE\Projeto - 1\Dados_Finais\Vendas Iguatemi.xlsx")
print("Criação de arquivo com valor específico, concluido")

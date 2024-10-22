import pandas as pd
import os  # Manipular arquivos/pastas

print("Iniciando o programa")
# Caminho da pasta com os arquivos
pasta = "data"

# Lista com todos os arquivos dentro da pasta.
# os.path.join: vai juntar todos os caminhos na variável "pasta" para a coleta do arquivo;
# os.listdir: vai listar os arquivos dentro da pasta;
arquivos = [os.path.join(pasta, arquivo) for arquivo in os.listdir(pasta)]

# Criação de tabela vazia
tabelaFinal = pd.DataFrame()

# tabela = pd.read_excel(arquivos[0], index_col="Data") ; o index_col pode ser usado para tornar uma coluna em um indice

vendas1 = pd.read_excel("data\\Vendas - 1.xlsx")

def checar_colunas(col_aux, col_vendas):
    colunas_diferentes = []
    for nome in col_aux:
        if nome != col_vendas.pop():
            colunas_diferentes.append(nome)
        else:
            return False
    
    return colunas_diferentes

# Loop para passar em todos os arquivos e ir juntando na tabela final
for arquivo in arquivos:
    #Vai ler cada arquivo em excel presente na pasta 
    aux = pd.read_excel(arquivo)
    #colunas = aux.columns.tolist() #obtem o nome de cada coluna do Dataframe
    colunas_vendas1 = vendas1.columns.tolist()
    colunas_aux = aux.columns.tolist()
    #print(colunas_aux)
    #print(colunas_vendas1)

    if checar_colunas(col_aux=colunas_aux, col_vendas=colunas_vendas1):
        print("Checagem concluida, colunas totalmente diferentes")
    else:
        print("Falha na checagem, colunas iguais.")
        break
    
    """#a função set() é usada para criar conjuntos
    if set(colunas_vendas1) == set(colunas_aux):
        print("É válido para mesclagem")
        df = pd.read_excel(arquivo)
        tabelaFinal = pd.concat([tabelaFinal, df], ignore_index= True)
    else:
        print("INVÁLIDO PARA MESCLAGEM: ", arquivo)"""


print("MESCLAGEM FEITA")





tabelaFinal.to_excel("Dados_Finais\\Vendas Totais.xlsx")
print("Criação de tabela final, concluida")

gerentes = pd.read_excel("data\\Gerentes.xlsx")
try:
    #junta duas tabelas, usando como base os valores semelhantes entre elas (como um procv)
    tabelaFinal = pd.merge(tabelaFinal, gerentes)
except Exception as e:
    print(e)

#vai excluir todas as colunas com valores vazios
tabelaFinal = tabelaFinal.dropna(how = 'all', axis=1)
try:
    tabelaFinal.to_excel("Dados_Finais\\Vendas Merge.xlsx")
    print("Criação de tabela final com os gerentes, concluida")
except Exception as e:
    print(f"Erro:{e}")

#vai realizar atransformação do arquivo para .json
try:
    tabelaFinal.to_json("Teste.json")
    print("Criação de arquivo .JSON, concluido")
except Exception as e:
    print(f"Erro:{e}")


# Criação de variável com um id de loja específico
vendasIguatemiEsp = tabelaFinal.loc[tabelaFinal["ID Loja"] == "Iguatemi Esplanada"]
vendasIguatemiEsp.to_excel("Dados_Finais\\Vendas Iguatemi.xlsx")
print("Criação de arquivo com valor específico, concluido")

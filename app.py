import pandas as pd
import os  # Manipular arquivos/pastas

print("Olá Python!")
# Caminho da pasta com os arquivos
pasta = r"C:\VS CODE\Projeto - 1\data"

# Lista com todos os arquivos dentro da pasta.
# os.path.join: vai juntar todos os caminhos na variável "pasta" para a coleta do arquivo;
# os.listdir: vai listar os arquivos dentro da pasta;
arquivos = [os.path.join(pasta, arquivo) for arquivo in os.listdir(pasta)]

# Criação de tabela vazia
tabelaFinal = pd.DataFrame()

# tabela = pd.read_excel(arquivos[0], index_col="Data") ; o index_col pode ser usado para tornar uma coluna em um indice
# print(tabela)

# Loop para passar em todos os arquivos e ir juntando na tabela final
for arquivo in arquivos:
    # Vai ser usado para armazenar a leitura de um arquivo Excel
    df = pd.read_excel(arquivo, index_col="Código Venda")

    # O concat tem a função de concatenar o df com a tabela final, assim juntando tudo em apenas uma variável
    tabelaFinal = pd.concat([tabelaFinal, df])

tabelaFinal.to_excel(r"C:\VS CODE\Projeto - 1\Dados_Finais\Vendas Totais.xlsx")

print("Mesclagem feita com sucesso")

# Criação de variável com um id de loja específico
vendasIguatemiEsp = tabelaFinal.loc[tabelaFinal["ID Loja"]
                                    == "Iguatemi Esplanada"]
print(vendasIguatemiEsp)

vendasIguatemiEsp.to_excel(
    "C:\VS CODE\Projeto - 1\Dados_Finais\Vendas Iguatemi.xlsx")

print("Envio feito com sucesso")

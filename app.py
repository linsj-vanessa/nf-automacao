import pandas as pd

# Carregar a planilha
df = pd.read_excel('/home/jodibe-pc/Documentos/nf-automacao/dados_nota.xlsm', engine='openpyxl', header=1)

# Verificar os nomes das colunas carregadas (opcional, para conferirmos)
print("Colunas da planilha:", df.columns)

# Selecionar colunas específicas
colunas_desejadas = ['Emitente', 'CNPJ','Nº NF-e', 'Série', 'Emissão', 'Chave na NF-e', 'Desc. Produto', 'Vlr  Total']

# Verificar se todas as colunas desejadas estão presentes
for coluna in colunas_desejadas:
    if coluna not in df.columns:
        print(f"A coluna '{coluna}' não foi encontrada na planilha.")

# Selecionar as colunas desejadas
dados = df[colunas_desejadas]

# Converter para lista de dicionários
lista_dados = dados.to_dict(orient='records')

# Exibir o resultado
print(lista_dados)


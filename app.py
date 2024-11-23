import pandas as pd


# Carregar a planilha
df = pd.read_excel('/home/jodibe-pc/Documentos/nf-automacao/dados_nota.xlsm', engine='openpyxl', header=1)

# # Verificar os nomes das colunas carregadas (opcional, para conferirmos)
# print("Colunas da planilha:", df.columns)

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

# # Exibir o valor da chave 'Emitente' do primeiro item na lista
# print(lista_dados[0]['Emitente'])





import pyautogui
import time

# Pausa inicial para dar tempo de preparar a tela
time.sleep(2)  # 2 segundos de espera para que você consiga focar na tela

# Passo 1: Clicar na primeira localização
pyautogui.click(x=403, y=186)

# Passo 2: Esperar 3 segundos para a página carregar
time.sleep(3)

# Passo 3: Clicar na segunda localização
pyautogui.click(x=256, y=113)

# Passo 4: Clicar na terceira localização
pyautogui.click(x=231, y=149)


# Passo 5: Digitar o CNPJ da primeira linha
# Aqui você pode acessar o CNPJ da primeira linha da lista de dados (ou de qualquer fonte)
cnpj = str(lista_dados[0]['CNPJ'])  # Exemplo, considerando que 'lista_dados' já está carregado

pyautogui.write(cnpj, interval=0.1) 

# Passo 6: Clica em pesquizar depois seleciona o primeiro CNPJ
pyautogui.click(x=160, y=180)
time.sleep(2)
pyautogui.click(x=133, y=271)
time.sleep(3)


# Passo 7: Nota/Série
pyautogui.click(x=322, y=253)
time.sleep(1)
nota = str(lista_dados[0]['Nº NF-e'])
pyautogui.write(nota, interval=0.1)

time.sleep(3)

pyautogui.click(x=436, y=251)
time.sleep(1)
serie = str(lista_dados[0]['Série'])
pyautogui.write(serie, interval=0.1)

time.sleep(1)
pyautogui.click(x=547, y=249)


# Passo 8: Operação
# Exibir uma caixa de entrada para o número da operação
operacao = input("Digite o número da operação: ")  # Pede ao usuário para digitar a operação

# Verificar se a entrada é válida (número)
if operacao.isdigit():
    operacao = str(operacao)  # Certifica-se de que a operação seja tratada como string
else:
    print("Entrada inválida! Por favor, digite um número.")
    exit()  # Encerra o programa caso a entrada não seja válida

# Clicar na localização para inserir a operação
pyautogui.click(x=323, y=286, clicks=2)

# Digitar a operação
pyautogui.write(operacao, interval=0.1)  # Digita o número da operação com intervalo de 0.1 segundos


import pandas as pd
import pyautogui
import time
import tkinter as tk
from tkinter import simpledialog
import os



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


# Função para digitar a data formatada
def digitar_data_emissao(data_emissao):
    if isinstance(data_emissao, pd.Timestamp):
        data_formatada = data_emissao.strftime('%d/%m/%Y')  # Converte para o formato dia/mês/ano
    else:
        data_formatada = str(data_emissao)  # Caso já esteja em string, converte diretamente
    pyautogui.write(data_formatada, interval=0.1)






# Início


# Iniciar o loop para cada linha
for linha in lista_dados:
    time.sleep(5)  # Espera inicial para foco na tela
    pyautogui.press('f5')  # Atualiza a página
    time.sleep(5)


    # Passo 1: Clicar na primeira localização
    pyautogui.click(x=403, y=186)

    # Passo 2: Esperar 3 segundos para a página carregar
    time.sleep(3)

    # Passo 3: Clicar na segunda localização
    pyautogui.click(x=256, y=113)

    # Passo 4: Clicar na terceira localização
    pyautogui.click(x=231, y=149)


    # Passo 5: Digitar o CNPJ da primeira linha
    cnpj = str(linha['CNPJ']) 

    pyautogui.write(cnpj, interval=0.1) 

    # Passo 6: Clica em pesquizar depois seleciona o primeiro CNPJ
    pyautogui.click(x=160, y=180)
    time.sleep(2)
    pyautogui.click(x=133, y=271)
    time.sleep(3)


    # Passo 7: Nota/Série
    pyautogui.click(x=322, y=253)
    time.sleep(1)
    nota = str(linha['Nº NF-e'])
    pyautogui.write(nota, interval=0.1)

    time.sleep(3)

    pyautogui.click(x=436, y=251)
    time.sleep(1)
    serie = str(lista_dados[0]['Série'])
    pyautogui.write(serie, interval=0.1)

    time.sleep(1)
    pyautogui.click(x=547, y=249)




    # Passo 8: Operação - Caixa de diálogo para o número da operação
    operacao = simpledialog.askstring("Operação", "Digite o número da operação:")

    # Verificar se a entrada é válida (não nula e numérica)
    if operacao and operacao.isdigit():
        # Clicar duas vezes na localização para inserir a operação
        time.sleep(2)
        pyautogui.click(x=323, y=286, clicks=2)
        
        # Digitar a operação
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.write(operacao, interval=0.1)  # Digita o número da operação com intervalo de 0.1 segundos
    else:
        print("Operação cancelada ou inválida. Pulando para a próxima nota.")
    continue  # Pula para a próxima iteração do loop    
        
        
        
    pyautogui.click(x=323, y=286)
    pyautogui.press('tab', presses=1, interval=0.2)
    pyautogui.click()




    # Passo 9: Emissão e Saída
    time.sleep(2)

    pyautogui.press('tab', presses=5, interval=0.2)
    
    # coluna 'Emissão' é convertida para string formatada (ex: 'DD/MM/YYYY')
    def digitar_data_emissao(data_emissao):
        if isinstance(data_emissao, pd.Timestamp):
            data_formatada = data_emissao.strftime('%d/%m/%Y')  # Converte para o formato dia/mês/ano
        else:
            data_formatada = str(data_emissao)  # Caso já esteja em string, converte diretamente
        
        pyautogui.write(data_formatada, interval=0.1)
        

    time.sleep(2)
    # Chama a função para digitar a data
    digitar_data_emissao(lista_dados[0]['Emissão'])

    pyautogui.press('tab', presses=1, interval=0.2)

    # Chama a função novamente para digitar a data na nova posição
    digitar_data_emissao(lista_dados[0]['Emissão'])


    # Entrada e vencimento
    pyautogui.press('tab', presses=1, interval=0.2)

    entrada = simpledialog.askstring("entrada", "Digite o vencimento da nota:")

    if entrada and entrada.isdigit():
        time.sleep(2)
            
        # Digitar a operação
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.write(entrada, interval=0.1) 
    else:
        print('erro na digitação') 
        pyautogui.press('tab')      
        

    #Cabeçalho finalizado




    
    # Passo 10: Chave de acesso 
    chave = str(lista_dados[0]['Chave na NF-e'])

    # Dividindo a chave em três partes
    parte1 = chave[:2]         # Os 2 primeiros números
    parte2 = chave[-10:-1]     # 9 penúltimos números (pulando o último)
    parte3 = chave[-1]         # O último número

    # Digit1ação com o PyAutoGUI

    # Passo 10: Digitar a Chave de Acesso

    # 2 primeiros números
    pyautogui.press('tab', presses=5, interval=0.2)
    pyautogui.write(parte1, interval=0.1)

    # 9 penúltimos números
    pyautogui.press('tab', presses=1, interval=0.2)
    pyautogui.write(parte2, interval=0.1)

    # Último número
    pyautogui.press('tab', presses=1, interval=0.2)
    pyautogui.write(parte3, interval=0.1)




    # Passo 11: Itens
    pyautogui.press('tab', presses=9, interval=0.2)
    pyautogui.click(x=1026, y=484)



    #Descrição do produto
    time.sleep(3)
    pyautogui.press('tab', presses=1, interval=0.1)
    time.sleep(2)
    produto = str(lista_dados[0]['Desc. Produto'])
    pyautogui.write(produto, interval=0.2)


    # Padrão da quantidade e unidade
    pyautogui.press('tab', presses=1, interval=0.2)
    pyautogui.write('1', interval=0.1)
    time.sleep(3)
    pyautogui.press('tab', presses=1, interval=0.2)
    time.sleep(2)
    pyautogui.press('tab', presses=1, interval=0.2)
    pyautogui.write('3', interval=0.1)


    time.sleep(5)

    #tributação
    # pyautogui.press('tab', presses=1, interval=0.2)
    # Pressiona a seta para baixo 3 vezes
    pyautogui.click(x=396, y=521)
    time.sleep(3)
    pyautogui.click(x=396, y=521)
    time.sleep(3)
    pyautogui.press('down', presses=3, interval=0.2)
    pyautogui.press('enter')



    # Perguntar a Conta
    conta = simpledialog.askstring("Conta", "Qual a conta?")

    # Verificar se a entrada é válida (não nula e numérica)
    if conta and conta.isdigit():
        pyautogui.click(x=323, y=630)

        
        # Digitar a conta
        pyautogui.write(conta, interval=0.1)  
    else:
        print("Entrada inválida! O programa será encerrado.")
        exit()  # Encerra o programa caso a entrada não seja válida
        

    # Perguntar o valor total

    vl = simpledialog.askstring("Vl Total", "Qual o valor total?")

    # Verificar se a entrada é válida (não nula e numérica)
    if vl and vl.isdigit():
        pyautogui.press('tab', presses=2, interval=0.1)
        time.sleep(2)
        pyautogui.hotkey('ctrl', 'a')
        # Digitar o valor total
        pyautogui.write(vl, interval=0.1)  
        pyautogui.press('tab', presses=2, interval=0.1)
    else:
        print("Entrada inválida! O programa será encerrado.")
        exit()  # Encerra o programa caso a entrada não seja válida
        

    time.sleep(3)
    pyautogui.click(x=1106, y=827)
    time.sleep(2)
    pyautogui.click(x=1447, y=828)
    time.sleep(2)
    pyautogui.click(x=1672, y=709)

    # Histórico
    time.sleep(3)
    pyautogui.click(x=476, y=263)

    frase = f"PAGTO {lista_dados[0]['Nº NF-e']} {lista_dados[0]['Emitente']}"
    pyautogui.write(frase)



    #Destino
    pyautogui.click(x=1172, y=532)
    pyautogui.click(x=1304, y=392, clicks=2)

    departamento = simpledialog.askstring("Departamento", "Qual o departamento?")

    # Verificar se a entrada é válida (não nula e numérica)
    if departamento and departamento.isdigit():
        time.sleep(2)
        # Digitar o valor total
        pyautogui.write(departamento, interval=0.1)  
        pyautogui.press('tab', presses=2, interval=0.1)
    else:
        print("Entrada inválida! O programa será encerrado.")
        exit()  # Encerra o programa caso a entrada não seja válida

    pyautogui.click(x=1410, y=424)

    #Salvar
    pyautogui.click(x=1694, y=680)
    time.sleep(3)


    # Pagamento
    pyautogui.press('tab')
    pyautogui.write(entrada, interval=0.1) 

    # Portador
    pyautogui.press('tab')
    pyautogui.press('down', presses=8)
    time.sleep(3)

    pyautogui.press('tab')
    pyautogui.press('down', presses=4)

    time.sleep(3)
    pyautogui.press('tab')
    pyautogui.press('down', presses=4)
    pyautogui.press('tab', presses=3)
    pyautogui.press('enter', presses=1)

    # 
    if operacao and operacao.isdigit():
        time.sleep(2)
        pyautogui.click(x=323, y=286, clicks=2)
        
        # Digitar a operação
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.write(operacao, interval=0.1)  # Digita o número da operação com intervalo de 0.1 segundos
    else:
        print("Operação cancelada ou inválida. Pulando para a próxima nota.")
    continue  # Pula para a próxima iteração do loop    
        

    print(f"Processo concluído para a nota: {linha['Nº NF-e']}")
    time.sleep(10)



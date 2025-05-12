import pyautogui    # Biblioteca para realizar as ações dos equipamentos de entrada de dados (mouse e teclado)
import time         # Biblioteca para configurar tempos de espera para a aplicação
import openpyxl     # Biblioteca para realizar ações com excel dentro do programa
import pyperclip    # Biblioteca para identificar caractere especiais nos campos da planilha
import json         # Biblioteca para identificar arquivos json
import os           # Biblioteca para identificar caminhos em pastas e arquivos do sistema operacional

# Formulário para cadastro dos funcionários, simulando a utilização de um sistema
# https://docs.google.com/forms/d/e/1FAIpQLSduehX0b4pdnbk1PBPxqWrynw4ADLntbk9uadJ-FMBPR3qPhg/viewform

# Declaração de abertura do documento e planilha correta do sistema
planilha = openpyxl.load_workbook('Funcionarios.xlsx')
pagina = planilha['Funcionarios']

# Acessando a linha 2 da planilha escolhida, para que comece a navegar entre as células, coletando as informações preenchidas
for linha in pagina.iter_rows(min_row = 2):     # Coleta a informação a partir da segunda linha da tabela (a primeira não é necessária nesse caso pois é o cabeçalho do arquivo)
    time.sleep(3)   # Tempo de espera até a próxima execução

# ID
    pyautogui.click(769,452, duration = 1)  # Realiza ação do mouse, encaminhando para o local indicado na coordenada
    funcionario_id = linha[0].value         # identifica a posição da tabela, assim como seu valor, e coloca na variável
    pyperclip.copy(funcionario_id)          # Copia o valor da variável
    pyautogui.hotkey('ctrl', 'v')           # Cola o valor no local selecionado

# Nome
    funcionario_nome = linha[1].value
    pyperclip.copy(funcionario_nome)
    pyautogui.press('tab', interval = 0.5)  # Troca para o próximo campo
    pyautogui.hotkey('ctrl', 'v')

# Cargo
    funcionario_cargo = linha[2].value
    pyperclip.copy(funcionario_cargo)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.hotkey('ctrl', 'v')

# Departamento
    funcionario_departamento = linha[3].value
    pyperclip.copy(funcionario_departamento)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.hotkey('ctrl', 'v')

# Idade
    funcionario_idade = linha[4].value
    pyperclip.copy(funcionario_idade)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.hotkey('ctrl', 'v')

# Data de Nascimento
    funcionario_data_nascimento = linha[5].value
    pyperclip.copy(funcionario_data_nascimento)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.hotkey('ctrl', 'v')

# E-mail
    funcionario_email = linha[6].value
    pyperclip.copy(funcionario_email)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.hotkey('ctrl', 'v')

    # Posiciona a tela para utilizar a função abaixo de uma forma melhor
    pyautogui.press('tab', interval = 0.5)
    pyautogui.press('tab', interval = 0.5)  

# Status
    funcionario_status = linha[7].value
    # Verifica o valor do status inserido na tabela, para movimentar o mouse para a opção desejada
    if funcionario_status == 'ativo':
        pyautogui.click(687,886, duration = 1)
    else:
        pyautogui.click(688,927, duration = 1)

    # Clica no botão "Enviar"
    pyautogui.press('tab', interval = 0.5)
    pyautogui.press('enter', interval = 0.5)
    time.sleep(2)
    # Clica no botão "Enviar outra resposta"
    pyautogui.press('tab', interval = 0.5)
    pyautogui.press('enter', interval = 0.5)

    # Coleta todos os dados e transforma em formato de lista para inserir no arquivo json
    dados_json = {"id": f'{funcionario_id}', "nome": f'{funcionario_nome}', "cargo": f'{funcionario_cargo}', "departamento": f'{funcionario_departamento}', "idade": f'{funcionario_idade}', "data_nascimento": f'{funcionario_data_nascimento}', "email": f'{funcionario_email}', "status": f'{funcionario_status}'}

    # Verifica se o arquivo existe ou não
    if os.path.exists('funcionarios.json'):
        with open('funcionarios.json', 'r+', encoding='utf-8') as file: # Abre o arquivo
            dados = json.load(file)     # Lê os dados
            dados.append(dados_json)    # Insere os dados no final da lista
            file.seek(0)
            json.dump(dados, file, ensure_ascii=False, indent=4)
    else:
        with open('funcionarios.json', 'w', encoding='utf-8') as file:  # Cria o arquivo
            json.dump([dados_json], file, ensure_ascii=False, indent=4)
    
# Limpando a tabela após realizar todos os cadastros
pagina.delete_rows(2, 100)
planilha.save('Funcionarios.xlsx')  # Salva a planilha

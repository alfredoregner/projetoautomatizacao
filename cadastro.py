import pyautogui    # Biblioteca para realizar as ações dos equipamentos de entrada de dados (mouse e teclado)
import time         # Biblioteca para configurar tempos de espera para a aplicação
import openpyxl     # Biblioteca para realizar ações com excel dentro do programa
import pyperclip    # Biblioteca para identificar caractere especiais nos campos da planilha

# Formulário para cadastro dos funcionários, simulando a utilização de um sistema
# https://docs.google.com/forms/d/e/1FAIpQLSduehX0b4pdnbk1PBPxqWrynw4ADLntbk9uadJ-FMBPR3qPhg/viewform

# Declaração de abertura do documento e planilha correta do sistema
planilha = openpyxl.load_workbook('Funcionarios.xlsx')
pagina = planilha['Funcionarios']

# Acessando a linha 2 da planilha escolhida, para que comece a navegar entre as células, coletando as informações preenchidas
for linha in pagina.iter_rows(min_row = 2):     # Coleta a informação a partir da segunda linha da tabela (a primeira não é necessária nesse caso pois é o cabeçalho do arquivo)
    time.sleep(2)

# ID
    pyautogui.click(715,432, duration = 1)
    funcionario_id = linha[0].value
    pyperclip.copy(funcionario_id)
    # pyautogui.press('tab', interval = 0.5)  # Troca para o próximo campo
    pyautogui.hotkey('ctrl', 'v')

# Nome
    funcionario_nome = linha[1].value
    pyperclip.copy(funcionario_nome)
    pyautogui.press('tab', interval = 0.5)  # Troca para o próximo campo
    pyautogui.hotkey('ctrl', 'v')

# Cargo
    funcionario_cargo = linha[2].value
    pyperclip.copy(funcionario_cargo)
    pyautogui.press('tab', interval = 0.5)  # Troca para o próximo campo
    pyautogui.hotkey('ctrl', 'v')

# Departamento
    funcionario_departamento = linha[3].value
    pyautogui.press('tab', interval = 0.5)
    pyperclip.copy(funcionario_departamento)
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

# Status
    funcionario_status = linha[7].value
    if funcionario_status == 'Sim':
        pyautogui.click(668,851, duration = 1)
    else:
        pyautogui.click(666,891, duration = 1)

    pyautogui.press('tab', interval = 0.5)
    pyautogui.press('enter', interval = 0.5)
    time.sleep(2)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.press('enter', interval = 0.5)

# Inserir os dados no arquivo "funcionarios.json"
# Playlist para entender como manipular arquivo.json
# https://www.youtube.com/watch?v=pM7EQKKs6Vg&list=PLZ6kIzk4n3uRmlJUAIwTLqMIIcgaR3uPa
    
# Limpando a tabela após realizar todos os cadastros
pagina.delete_rows(2, 100)  # Encontrar uma forma de automatizar a identificação da última linha da tabela, para que seja tudo deletado automaticamente a partir da primeira linha com dados na tabela (2), até a última linha com dados preenchidos
planilha.save('Funcionarios.xlsx')


# Realizar refino no código
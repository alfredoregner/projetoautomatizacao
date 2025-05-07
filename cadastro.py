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
<<<<<<< HEAD
    time.sleep(2)
=======
>>>>>>> 339b305623bd01ca742f42b7d44c064ad5e7f484
    # # Nome
    # funcionario_nome = linha[0].value       # Acessa o campo selecionado de acordo com o índice ([0] é o primeiro campo da linha 2)
    # pyperclip.copy(funcionario_nome)        # Copia a informação do campo selecionado, com a formatação original

    # # Click com o pc do SENAI
<<<<<<< HEAD
    # pyautogui.click(715,432, duration = 1)
=======
    # pyautogui.click(666,431, duration = 1)
>>>>>>> 339b305623bd01ca742f42b7d44c064ad5e7f484
    
    # # Clickcom o pc do trabalho
    # # pyautogui.click(735,444, duration = 1)  # Clica no local indicado pelas coordenadas
    # pyautogui.hotkey('ctrl', 'v')           # Cola a informação copiada anteriormente, no local selecionado
    
# ID
<<<<<<< HEAD
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

=======
    funcionario_id = linha[0].value
    pyperclip.copy(funcionario_id)
    pyautogui.press('tab', interval = 0.5)  # Troca para o próximo campo
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

>>>>>>> 339b305623bd01ca742f42b7d44c064ad5e7f484
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

    # # btn-Próximo
    # pyautogui.press('tab', interval = 1)
    # pyautogui.press('enter', interval = 1)  # Pressiona Enter para seguir para a próxima parte do formulário
    # time.sleep(2)

    # Data de Nascimento
    funcionario_data_nascimento = linha[5].value
    pyperclip.copy(funcionario_data_nascimento)
    pyautogui.press('tab', interval = 0.5)

<<<<<<< HEAD
=======
    # Data de Nascimento
    funcionario_data_nascimento = linha[5].value
    pyperclip.copy(funcionario_data_nascimento)
>>>>>>> 339b305623bd01ca742f42b7d44c064ad5e7f484
    # Click com o pc do SENAI
    # pyautogui.click(658,445, duration = 1)
    
    # Clickcom o pc do trabalho
    # pyautogui.click(783,471, duration = 1) 
    pyautogui.hotkey('ctrl', 'v')

    # E-mail
    funcionario_email = linha[6].value
    pyperclip.copy(funcionario_email)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.hotkey('ctrl', 'v')

    # Status
    funcionario_status = linha[7].value
<<<<<<< HEAD
    if funcionario_status == 'Sim':
        pyautogui.click(668,851, duration = 1)
    else:
        pyautogui.click(666,891, duration = 1)
=======
    pyperclip.copy(funcionario_status)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.hotkey('ctrl', 'v')
>>>>>>> 339b305623bd01ca742f42b7d44c064ad5e7f484

    # # btn-Próximo
    # pyautogui.press('tab', interval = 1)
    # pyautogui.press('tab', interval = 1)
    # pyautogui.press('enter', interval = 1)
    # time.sleep(2)

    #     # Clickcom o pc do trabalho
    #     # pyautogui.click(697,463, duration = 1)
    # elif funcionario_cargo == 'Supervisor':
    #     # Click com o pc do SENAI
    #     pyautogui.click(676,479, duration = 1)
    
    #     # Clickcom o pc do trabalho
    #     # pyautogui.click(696,503, duration = 1)
    # elif         funcionario_cargo == 'Repositor':
    #     # Click com o pc do SENAI
    #     pyautogui.click(678,519, duration = 1)
    
    #     # Clickcom o pc do trabalho
    #     # pyautogui.click(695,545, duration = 1)
    # else:
    #     # Click com o pc do SENAI
    #     pyautogui.click(675,563, duration = 1)
    
    #     # Clickcom o pc do trabalho
    #     # pyautogui.click(694,583, duration = 1)
        
    # # Salário
    # funcionario_salario = linha[12].value
    # pyperclip.copy(funcionario_salario)
    # pyautogui.press('tab', interval = 0.5)
    # pyautogui.hotkey('ctrl', 'v')

    # # btn-Enviar
    # pyautogui.press('tab', interval = 1)
    # pyautogui.press('tab', interval = 1)
    # pyautogui.press('enter', interval = 1)
    # time.sleep(2)

    # # Enviara outra resposta
    # # Click com o pc do SENAI
    # pyautogui.click(717,246, duration = 1)
    
    # Clickcom o pc do trabalho
    # pyautogui.click(741,273, duration = 1)
    # time.sleep(2)
<<<<<<< HEAD
    pyautogui.press('tab', interval = 0.5)
    pyautogui.press('enter', interval = 0.5)
    time.sleep(2)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.press('enter', interval = 0.5)
=======
>>>>>>> 339b305623bd01ca742f42b7d44c064ad5e7f484

# Limpando a tabela após realizar todos os cadastros
pagina.delete_rows(2, 100)  # Encontrar uma forma de automatizar a identificação da última linha da tabela, para que seja tudo deletado automaticamente a partir da primeira linha com dados na tabela (2), até a última linha com dados preenchidos
planilha.save('Funcionarios.xlsx')


# Realizar refino no código
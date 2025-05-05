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
    # Nome
    funcionario_nome = linha[0].value       # Acessa o campo selecionado de acordo com o índice ([0] é o primeiro campo da linha 2)
    pyperclip.copy(funcionario_nome)        # Copia a informação do campo selecionado, com a formatação original

    # Click com o pc do SENAI
    pyautogui.click(666,431, duration = 1)
    
    # Clickcom o pc do trabalho
    # pyautogui.click(735,444, duration = 1)  # Clica no local indicado pelas coordenadas
    pyautogui.hotkey('ctrl', 'v')           # Cola a informação copiada anteriormente, no local selecionado
    
    # CPF
    funcionario_cpf = linha[1].value
    pyperclip.copy(funcionario_cpf)
    pyautogui.press('tab', interval = 0.5)  # Troca para o próximo campo
    pyautogui.hotkey('ctrl', 'v')

    # E-mail
    funcionario_email = linha[2].value
    pyautogui.press('tab', interval = 0.5)
    pyperclip.copy(funcionario_email)
    pyautogui.hotkey('ctrl', 'v')

    # Telefone
    funcionario_telefone = linha[3].value
    pyperclip.copy(funcionario_telefone)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.hotkey('ctrl', 'v')

    # btn-Próximo
    pyautogui.press('tab', interval = 1)
    pyautogui.press('enter', interval = 1)  # Pressiona Enter para seguir para a próxima parte do formulário
    time.sleep(2)

    # CEP
    funcionario_cep = linha[4].value
    pyperclip.copy(funcionario_cep)
    # Click com o pc do SENAI
    pyautogui.click(658,445, duration = 1)
    
    # Clickcom o pc do trabalho
    # pyautogui.click(783,471, duration = 1) 
    pyautogui.hotkey('ctrl', 'v')

    # Rua
    funcionario_rua = linha[5].value
    pyperclip.copy(funcionario_rua)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.hotkey('ctrl', 'v')

    # Número
    funcionario_numero = linha[6].value
    pyperclip.copy(funcionario_numero)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.hotkey('ctrl', 'v')

    # Complemento
    funcionario_complemento = linha[7].value
    pyperclip.copy(funcionario_complemento)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.hotkey('ctrl', 'v')

    # Bairro
    funcionario_bairro = linha[8].value
    pyperclip.copy(funcionario_bairro)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.hotkey('ctrl', 'v')

    # Cidade
    funcionario_cidade = linha[9].value
    pyperclip.copy(funcionario_cidade)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.hotkey('ctrl', 'v')

    # Estado
    
    funcionario_estado = linha[10].value
    pyperclip.copy(funcionario_estado)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.hotkey('ctrl', 'v')

    # btn-Próximo
    pyautogui.press('tab', interval = 1)
    pyautogui.press('tab', interval = 1)
    pyautogui.press('enter', interval = 1)
    time.sleep(2)

    # Cargo
    funcionario_cargo = linha[11].value
    if funcionario_cargo == 'Gerente':         # Verifica o que está escrito na célula e compara com as condições para clicar na opção correta
        # Click com o pc do SENAI
        pyautogui.click(676,445, duration = 1)
    
        # Clickcom o pc do trabalho
        # pyautogui.click(697,463, duration = 1)
    elif funcionario_cargo == 'Supervisor':
        # Click com o pc do SENAI
        pyautogui.click(676,479, duration = 1)
    
        # Clickcom o pc do trabalho
        # pyautogui.click(696,503, duration = 1)
    elif funcionario_cargo == 'Repositor':
        # Click com o pc do SENAI
        pyautogui.click(678,519, duration = 1)
    
        # Clickcom o pc do trabalho
        # pyautogui.click(695,545, duration = 1)
    else:
        # Click com o pc do SENAI
        pyautogui.click(675,563, duration = 1)
    
        # Clickcom o pc do trabalho
        # pyautogui.click(694,583, duration = 1)
        
    # Salário
    funcionario_salario = linha[12].value
    pyperclip.copy(funcionario_salario)
    pyautogui.press('tab', interval = 0.5)
    pyautogui.hotkey('ctrl', 'v')

    # btn-Enviar
    pyautogui.press('tab', interval = 1)
    pyautogui.press('tab', interval = 1)
    pyautogui.press('enter', interval = 1)
    time.sleep(2)

    # Enviara outra resposta
    # Click com o pc do SENAI
    pyautogui.click(717,246, duration = 1)
    
    # Clickcom o pc do trabalho
    # pyautogui.click(741,273, duration = 1)
    time.sleep(2)

# Limpando a tabela após realizar todos os cadastros
pagina.delete_rows(2, 100)  # Encontrar uma forma de automatizar a identificação da última linha da tabela, para que seja tudo deletado automaticamente a partir da primeira linha com dados na tabela (2), até a última linha com dados preenchidos
planilha.save('Funcionarios.xlsx')


# Realizar refino no código
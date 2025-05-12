import time         # Biblioteca para configurar tempos de espera para a aplicação
import openpyxl     # Biblioteca para realizar ações com excel dentro do programa

# Inicio da conversa
print('Bem vindo(a). Eu serei sua assistente virtual, estou pronta para te ajudar!')
time.sleep(0.5)
print('Digite o número correspondente ao atendimento que deseja: ')
time.sleep(0.5)
print('1- Cadastro de Funcionário \n2- Edição de Cadastro \n3- Encerrar atendimento')
time.sleep(0.5)
atendimento = int(input('Insira sua opção: '))
time.sleep(1)

# Verifica se a resposta enviada foi correta para um dos atendimentos informados
while atendimento != 1 and atendimento != 2 and atendimento != 3:
    print('Opção inválida!')
    time.sleep(0.5)
    print('1- Cadastro de Funcionário \n2- Edição de Cadastro \n3- Encerrar atendimento')
    time.sleep(0.5)
    atendimento = int(input('Insira sua opção: '))
    time.sleep(1)

# Acessa a planilha com os dados dos funcionários
planilha = openpyxl.load_workbook('Funcionarios.xlsx')
pagina = planilha['Funcionarios']

# Cadastro de Funcionário
if atendimento == 1:
    # Coleta os dados que serão inseridos na planilha
    cadastro_id = input('Insira o ID do funcionário: ')
    time.sleep(0.5)
    cadastro_nome = input('Insira o Nome do funcionário: ')
    time.sleep(0.5)
    cadastro_cargo = input('Insira o Cargo do funcionário: ')
    time.sleep(0.5)
    cadastro_departamento = input('Insira o Departamento do funcionário: ')
    time.sleep(0.5)
    cadastro_idade = input('Insira a Idade do funcionário: ')
    time.sleep(0.5)
    cadastro_nascimento = input('Insira a Data de Nascimeno (dd/mm/aaaa) do funcionário: ')
    time.sleep(0.5)
    cadastro_email = input('Insira o Email do funcionário: ')
    time.sleep(0.5)
    cadastro_ativo = input('Insira "Ativo" ou "Inativo" para o status do funcionário ').lower()
    time.sleep(0.5)

    # Insere os dados na planilha
    pagina.append([cadastro_id, cadastro_nome, cadastro_cargo, cadastro_departamento, cadastro_idade, cadastro_nascimento, cadastro_email, cadastro_ativo])
    planilha.save('Funcionarios.xlsx') # Salva a planilha com os dados atualizados

    print(f'Dados do funcionário {cadastro_nome} cadastrados com sucesso!')

# Edição de Cadastro de Funcionário
elif atendimento == 2:
    # Inserir o ID do cadastro desejado para alteração
    id_usuario = input('Insira o ID do funcionário que deseja alterar: ')
    for rows in pagina.iter_rows(min_row=2, max_row=None):
        # Verificar se o ID é valido, e acessar a linha de cadastro desse funcionário
        if id_usuario == rows[0].value: 
            # print('O que você deseja alterar?')
            # print('1- Dados Básicos do Funcionário \n2- Endereço do Funcionário \n3- Função do Funcionário')
            # editar = int(input('Insira sua opção: '))
            
            # while editar < 1 and editar > 3: 
            #     print('Opção inválida!')
            #     editar = input('1- Dados Básicos do Funcionário \n2- Endereço do Funcionário \n3- Função do Funcionário')
            #     editar = int(input('Insira sua opção: '))

        # Realizar a alteração de todos dados do funcionário
        #if editar == 1:
            time.sleep(0.5)
            edicao_nome = input('Insira o Nome do funcionário: ')
            time.sleep(0.5)
            edicao_cargo = input('Insira o Cargo do funcionário: ')
            time.sleep(0.5)
            edicao_departamento = input('Insira o Departamento do funcionário: ')
            time.sleep(0.5)
            edicao_idade = input('Insira a Idade do funcionário: ')
            time.sleep(0.5)
            edicao_nascimento= input('Insira a Data de Nascimeno (dd/mm/aaaa) do funcionário: ')
            time.sleep(0.5)
            edicao_email = input('Insira o Email do funcionário: ')
            time.sleep(0.5)
            edicao_ativo = input('Insira "Ativo" ou "Inativo" para o status do funcionário ').lower()
          
            time.sleep(0.5)
            
            # Insere os novos dados na planilha
            rows[1].value = edicao_nome
            rows[2].value = edicao_cargo
            rows[3].value = edicao_departamento
            rows[4].value = edicao_idade
            rows[5].value = edicao_nascimento
            rows[6].value = edicao_email
            rows[7].value = edicao_ativo

            # Salva a planilha com os dados atualizados
            planilha.save('Funcionarios.xlsx') 

            print(f'Dados do funcionário {rows[1].value} alterados com sucesso!')
        
        # Caso o id do funcionário na planilha não seja identificado
        else:
            print('ID Inválido')
            time.sleep(1)

# Caso a opção desejada seja encerrar o contato
else :
    print('Agradecemos pelo seu contato. Tenha um ótimo dia!')
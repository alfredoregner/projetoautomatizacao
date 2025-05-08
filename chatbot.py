import time         # Biblioteca para configurar tempos de espera para a aplicação
import openpyxl     # Biblioteca para realizar ações com excel dentro do programa

print('Bem vindo ao Suporte da Maculados LTDA. Eu serei sua assistente virtual, estou pronta para te ajudar!')
print('Digite o número correspondente ao atendimento que deseja: ')
print('1- Cadastro de Funcionário \n2- Edição de Cadastro \n3- Encerrar atendimento')
atendimento = int(input('Insira sua opção: '))
time.sleep(1)

# Acessa a planilha com os dados dos funcionários
planilha = openpyxl.load_workbook('Funcionarios.xlsx')
pagina = planilha['Funcionarios']

while atendimento < 1 and atendimento > 3:
    print('Opção inválida!')
    print('1- Cadastro de Funcionário \n2- Edição de Cadastro \n3- Encerrar atendimento')
    atendimento = int(input('Insira sua opção: '))
    time.sleep(1)

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
    cadastro_idade = input('Insira o Idade do funcionário: ')
    time.sleep(0.5)
    cadastro_nascimento = input('Insira a Data de Nascimeno do funcionário: ')
    time.sleep(0.5)
    cadastro_email = input('Insira o Email do funcionário: ')
    time.sleep(0.5)
    cadastro_ativo = input('Insira "true" ou "false" para o status do funcionário ')
    time.sleep(0.5)

    # Insere os dados na planilha
    pagina.append([cadastro_id, cadastro_nome, cadastro_cargo, cadastro_departamento, cadastro_idade, cadastro_nascimento, cadastro_email, cadastro_ativo])
    planilha.save('Funcionarios.xlsx') # Salva a planilha com os dados atualizados

    print(f'Dados do funcionário {cadastro_nome} cadastrados com sucesso!')


elif atendimento == 2:
    # Inserir o CPF do cadastro desejado para alteração
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

            # Realizar a alteração de todos os campos de dados básicos do funcionário
        #if editar == 1:
            edicao_nome = input('Insira o Nome do funcionário: ')
            time.sleep(0.5)
            edicao_departamento = input('Insira o departamento do funcionário: ')
            time.sleep(0.5)
            edicao_email = input('Insira o E-mail do funcionário: ')
            time.sleep(0.5)
            edicao_cargo = input('Insira o cargo do funcionário: ')
            time.sleep(0.5)
            edicao_ativo = input('Insira o cargo do funcionário: ')
            time.sleep(0.5)
            edicao_idade = input('Insira o cargo do funcionário: ')
            time.sleep(0.5)
            edicao_nascimento= input('Insira o cargo do funcionário: ')
            time.sleep(0.5)
            
            # Insere os dados na planilha
            rows[1].value = edicao_nome
            rows[2].value = edicao_cargo
            rows[3].value = edicao_departamento
            rows[4].value = edicao_idade
            rows[5].value = edicao_nascimento
            rows[6].value = edicao_email
            rows[7].value = edicao_ativo
            
            planilha.save('Funcionarios.xlsx') # Salva a planilha com os dados atualizados

            print(f'Dados do funcionário {rows[1].value} alterados com sucesso!')

            # Realizar a alteração de todos os campos de endereço do funcionário
            # elif editar == 2:
            #     edicao_cep = input('Insira o CEP do endereço do funcionário: ')
            #     time.sleep(0.5)
            #     edicao_rua = input('Insira a Rua do endereço funcionário: ')
            #     time.sleep(0.5)
            #     edicao_numero = input('Insira o Número do endereço do funcionário: ')
            #     time.sleep(0.5)
            #     edicao_complemento = input('Insira o Complemento do endereço do funcionário: ')
            #     time.sleep(0.5)
            #     edicao_bairro = input('Insira o Bairro do endereço do funcionário: ')
            #     time.sleep(0.5)
            #     edicao_cidade = input('Insira a Cidade do endereço do funcionário: ')
            #     time.sleep(0.5)
            #     edicao_estado = input('Insira o Estado do endereço do funcionário: ')
            #     time.sleep(0.5)

                # Insere os dados na planilha
                # rows[4].value = edicao_cep
                # rows[5].value = edicao_rua
                # rows[6].value = edicao_numero
                # rows[7].value = edicao_complemento
                # rows[8].value = edicao_bairro
                # rows[9].value = edicao_cidade
                # rows[10].value = edicao_estado
                # planilha.save('Funcionarios.xlsx') # Salva a planilha com os dados atualizados

                # print(f'Dados de endereço do funcionário {rows[0].value} alterados com sucesso!')

            # Realizar a alteração de todos os campos de função do funcionário
        #     else:                    
        #         edicao_cargo = input('Insira o Cargo do funcionário: ')
        #         time.sleep(0.5)
        #         edicao_salario = input('Insira o Salário do funcionário: ')
        #         time.sleep(0.5)

            # # Insere os dados na planilha
            # rows[11].value = edicao_cargo
            # rows[12].value = edicao_salario
            # planilha.save('Funcionarios.xlsx') # Salva a planilha com os dados atualizados

            # print(f'Dados de função do funcionário {rows[0].value} alterados com sucesso!')
        
        else:
            print('ID Inválido')
            time.sleep(1)

else :
    print('Agradecemos pelo seu contato. Tenha um ótimo dia!')


# Realizar refino no código
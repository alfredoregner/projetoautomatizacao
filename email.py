import win32com.client as win32     # Biblioteca para identificar e acessar os arquivos e programas do sistema
import json                         # Biblioteca para identificar arquivos json
import datetime                     # Biblioteca para identificar data e hora

# instalar: pip install pywin32
# instalar: pip install datetime

# Criar a integração com o Outlook
outlook = win32.Dispatch('outlook.application')

# verifica a data atual, e transforma no formato utilizado no brasil (dd/mm/aaaa)
data_atual = datetime.datetime.now().date()
# Coleta o dia atual
dia_formatado = data_atual.strftime('%d')
# Coleta o mês atual
mes_formatado = data_atual.strftime('%m')

# Realiza a leitura do arquivo funcionarios.json para coletar os dados necessários
with open('funcionarios.json', 'r+', encoding='utf-8') as file:
    load = json.loads(file.read())
    

    # Realiza um loop para verificar todos os itens da lista
    for lista in load:
        
        # Trasnforma a data cadastrada do funcionario para coletar somente o dia do nascimento
        dia_funcionario = lista["data_nascimento"][:2]
        # Trasnforma a data cadastrada do funcionario para coletar somente o mês do nascimento
        mes_funcionario = lista["data_nascimento"][3:5]

        # Compara o dia e o mês cadastrado de cada funcionario com o dia e o mês da data atual
        if dia_formatado == dia_funcionario and mes_funcionario == mes_formatado:
            
            # Criar um email
            email = outlook.CreateItem(0)

            # Configura as Informações de envio do e-mail
            # Destinatário
            email.To = lista['email']   
            # Título da mensagem
            email.Subject = f"Feliz Aniversário {lista['nome']} "   
            # Corpo do Texto da Mensagem
            email.HTMLBody = f'''  
            <p>Prezado(a) {lista['nome']}! Hoje é o seu aniversário</p>
            <p>Hoje é um dia especial, e queremos aproveitar a oportunidade para lhe desejar um feliz aniversário!</p>
            <p>Que esta data seja repleta de alegrias, saúde, conquistas e momentos felizes ao lado das pessoas que você ama.</p>
            <p>Agradecemos pelo seu comprometimento, dedicação e por fazer parte da nossa equipe.</p>
            <p>Desejamos muito sucesso pessoal e profissional, hoje e sempre.</p>
            <p>Parabéns!</p>
            '''
            # Realiza o envio do E-mail
            email.Send()
            print("E-mail Enviado com sucesso!")

        else:
            print('Nenhum funcionário faz aniversário hoje')

    
    


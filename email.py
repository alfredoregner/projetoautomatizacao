import win32com.client as win32     # Biblioteca para identificar e acessar os arquivos e programas do sistema
import json                         # Biblioteca para identificar arquivos json
import datetime                     # Biblioteca para identificar data e hora

# instalar: pip install pywin32
# instalar: pip install datetime

# Criar a integração com o Outlook
outlook = win32.Dispatch('outlook.application')

# verifica a data atual, e transforma no formato utilizado no brasil (dd/mm/aaaa)
data_atual = datetime.datetime.now().date()
data_formatada = data_atual.strftime('%d/%m/%Y')

# Realiza a leitura do arquivo funcionarios.json para coletar os dados necessários
with open('funcionarios.json', 'r+', encoding='utf-8') as file:
    load = json.loads(file.read())

    # Realiza um loop para verificar todos os itens da lista
    for lista in load:
        # Compara a data de nascimento cadastrada de cada funcionario com a data atual formatada
        if lista["data_nascimento"] == data_formatada:
            
            # Criar um email
            email = outlook.CreateItem(0)

            # Configurar as Informações do seu e-mail
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

    
    


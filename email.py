import win32com.client as win32
import json
import datetime

# instalar: pip install pywin32

# Criar a integração com o Outlook
outlook = win32.Dispatch('outlook.application')

data_atual = datetime.datetime.now().date()

with open('funcionarios.json', 'r+', encoding='utf-8') as file:
    load = json.loads(file.read())

    for lista in load:
        if lista["data_nascimento"] == data_atual:
            # Criar um email
            email = outlook.CreateItem(0)

            # Configurar as Informações do seu e-mail
            email.To = lista['email']
            email.Subject = 'Teste de Disparo de E-mail Automatizado'
            email.HTMLBody = '''
            <p>Olá Alfredo! Aqui é o código Python</p>

            <p>O Faturamento da loja foi de R$1500
            Vendemos 10 produtos
            O ticket Médio foi de R$150</p>

            <p>Abs,</p>
            <p>Alfredo Tester</p>
            '''

            email.Send()
            print("E-mail Enviado com sucesso!")

    
    


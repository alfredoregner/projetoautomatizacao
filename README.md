# 🖥️ Automação de Tarefas Internas

Este projeto tem como objetivo implementar soluções de **automação de processos internos** para um ambiente de escritório. A aplicação integra manipulação de arquivos, automação de formulários e simulação de um assistente virtual para tornar o ambiente de trabalho mais eficiente.

---

## 👥 Equipe

- **Alfredo Walter Regner Neto** – PO  
- **Breno Luiz Raghianti Zein** – Scrum Master  
- **Carla Machado Miguel** – PO  
- **Matheus Rodrigues Antonio** – Dev Team  

---

## 🎯 Objetivos

- Reduzir tarefas repetitivas por meio de automação  
- Simular um ambiente corporativo com dados fictícios realistas  
- Treinar o uso de bibliotecas como `pyautogui`, `pyperclip`, `openpyxl` e `json`  
- Criar uma interface simples de atendimento virtual com Python  

---

## 🚀 Funcionalidades

- 📑 **Leitura e Escrita em Planilhas Excel** com `openpyxl`
- 🧠 **Automação de Formulários** com `pyautogui` e `pyperclip`
- 🤖 **Chatbot Simples** para interações com o usuário
- 👨‍💼 **Gerenciamento de Funcionários**
- 🧾 **Manipulação JSON** estruturada para cadastro e atualização de dados
- 📧 **Disparo de E-mail** com 'win32com.client'

---

## 🛠️ Tecnologias Utilizadas

- Python 3.13.3
- [openpyxl](https://pypi.org/project/openpyxl/)
- [pyautogui](https://pypi.org/project/PyAutoGUI/)
- [pyperclip](https://pypi.org/project/pyperclip/)
- Bibliotecas nativas: `time`, `json`, `datetime`, `os`
- Extensões utilizadas no Visual Studio Code: Python (Para executar arquivos Python), Python Debugger (Para fazer debug do código com melhor eficiência), Excel Viewer (Para visualizar a Planilha Excel utilizada dentro do Visual Studio)

---

## 🔧 Como Executar

1. Clone o repositório:

  ``https://github.com/alfredoregner/projetoautomatizacao.git``

2. Instale as dependências:
Deve ser executado os seguites comandos no prompt, para instalar as dependências necessárias para execução do projeto:
``pip install openpyxl``
``pip install pyautogui``
``pip install pyperclip``
``pip install pywin32``


3. Execute os scripts conforme a funcionalidade desejada:

``chatbot.py - Para realizar o cadastro de funcionários``
``cadastro.py - Para inserir os dados dentro do sistema``
``email.py - Para encaminhar mensagens de e-mail para os funcionários``


4. Alterar posicionamento do mouse (caso necessário):
Deve ser executado no prompt de comando, os seguintes comandos:
``python``
``from mouseinfo import mouseInfo``
``mouseInfo()``

Essa sequência de comandos abre um aplicativo nativo do python que exibe as informações do ponteiro do mouse. Pode ser utilizado para coletar os dados de coordenadas do ponteiro do mouse, para que seja atualizado no sistema e realizar as ações nos locais desejados

---

🧩 Organização e Arquitetura

O projeto foi modularizado em scripts com responsabilidades distintas:
Separação por arquivos: chatbot de cadastro, automação e disparo de e-mail independentes
Utilização de padrões simples: como Pub/Sub via arquivos externos (JSON/Excel)
Facilidade de manutenção e expansão futura

---

📌 Considerações Finais

Este projeto demonstra na prática como automação simples com Python pode trazer ganhos de produtividade, mesmo em ambientes pequenos.

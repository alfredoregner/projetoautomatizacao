# 🖥️ Automação de Tarefas Internas — *The Office Edition*

Este projeto tem como objetivo implementar soluções de **automação de processos internos** para um ambiente de escritório, utilizando os personagens da série *The Office* como base para simulação de dados e interações. A aplicação integra manipulação de arquivos, automação de formulários e simulação de um assistente virtual para tornar o ambiente de trabalho mais eficiente e divertido.

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
- 👨‍💼 **Gerenciamento de Funcionários** com dados baseados em personagens de *The Office*
- 🧾 **Manipulação JSON** estruturada para cadastro e atualização de dados

---

## 🛠️ Tecnologias Utilizadas

- Python 3.x
- [openpyxl](https://pypi.org/project/openpyxl/)
- [pyautogui](https://pypi.org/project/PyAutoGUI/)
- [pyperclip](https://pypi.org/project/pyperclip/)
- Bibliotecas nativas: `time`, `json`

---

## 🔧 Como Executar

1. Clone o repositório:

  ``git clone https://github.com/seu-usuario/seu-repositorio.git
   cd seu-repositorio``

3. Instale as dependências:
4. 
``pip install openpyxl pyautogui pyperclip``

5. Execute os scripts conforme a funcionalidade desejada:

``python cadastro.py
python chatbot.py
python manipulacao_personagens_the_office.py``

---

🧪 Exemplos de Comandos JSON
A partir do arquivo personagens_com_comandos.json, é possível realizar:

✅ Adição:

``personagens.push({
  id: 16,
  nome: "Toby Flenderson",
  cargo: "RH",
  ...
});``

---

✏️ Alteração:

``personagens.find(p => p.nome === "Pam Beesly").cargo = "Design Gráfico";``

❌ Remoção:

``personagens = personagens.filter(p => p.nome !== "Ryan Howard");``

---

🧩 Organização e Arquitetura

O projeto foi modularizado em scripts com responsabilidades distintas:
Separação por arquivos: automação, cadastro e chatbot independentes
Utilização de padrões simples: como Pub/Sub via arquivos externos (JSON/Excel)
Facilidade de manutenção e expansão futura

---

📌 Considerações Finais

Este projeto demonstra na prática como automação simples com Python pode trazer ganhos de produtividade, mesmo em ambientes pequenos. Além disso, a escolha de personagens de The Office como dados fictícios torna o processo mais divertido e pedagógico.

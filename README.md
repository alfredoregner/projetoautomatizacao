Questioário Apresentação

1. Como funciona o pip no Python e como ele foi utilizado no projeto?
   
O pip é o gerenciador de pacotes do Python. Ele permite instalar bibliotecas externas que não fazem parte da biblioteca padrão, facilitando a inclusão de funcionalidades específicas em projetos.

No projeto apresentado, o pip foi utilizado para instalar as seguintes bibliotecas externas:

openpyxl: para leitura, escrita e manipulação de arquivos Excel (.xlsx), usados como base de dados dos funcionários.

pyautogui: para automação de interações com mouse e teclado, simulando o preenchimento de formulários online.

pyperclip: para copiar e colar textos corretamente (inclusive com acentuação), integrando-se com o pyautogui.

Essas bibliotecas podem ser instaladas com: "pip install, openpyxl, pyautogui, pyperclip"

Além dessas, o projeto também utilizou a biblioteca time, que faz parte da biblioteca padrão do Python (ou seja, não precisa ser instalada com pip). Ela foi usada para controlar pausas no programa, criando esperas intencionais entre ações, como inputs do usuário e cliques automatizados, garantindo fluidez na interação e evitando falhas de sincronização.

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

2. Como foi feita a organização dos arquivos e pastas no projeto?
   
O projeto foi organizado com base nos seguintes elementos principais:

Arquivos .py: códigos-fonte do sistema de atendimento virtual e automação de cadastros (como o script com o menu interativo e o script com pyautogui).

Arquivo .xlsx: Funcionarios.xlsx, que armazena as informações dos funcionários (como ID, nome, cargo, etc.). Ele serve como base de dados local.

Arquivo .json/.py com personagens: representa uma simulação de manipulação de dados no formato de dicionários Python, útil para representar e testar operações CRUD (Criar, Ler, Atualizar e Deletar).

A estrutura pode ser representada assim:

/projeto


├── atendimento_virtual.py        # Script com menu de cadastro/edição

├── automacao_formulario.py       # Script com pyautogui para preencher formulário

├── personagens_com_comandos.json # Lista de personagens para manipulação de dados

└── Funcionarios.xlsx             # Planilha com os dados dos funcionários

----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

3. O que é arquitetura de software e como o grupo aplicou isso no projeto?
   
Arquitetura de software é a estrutura organizacional fundamental de um sistema, incluindo seus componentes, suas relações e os princípios de projeto que orientam sua evolução. 
Em projetos simples, isso pode ser aplicado com organização modular, separação de responsabilidades e reutilização de código.

No projeto, a arquitetura foi aplicada da seguinte forma:

Separação de responsabilidades:

Um script cuida do atendimento virtual com menu interativo e manipulação de planilhas (usando a biblioteca openpyxl).

Outro script realiza a automação da interface com mouse e teclado (usando pyautogui), preenchendo formulários online com os dados cadastrados.

O arquivo JSON com personagens foi usado como simulação para treinar a lógica de manipulação de dados estruturados, de forma semelhante ao tratamento feito na planilha Excel.

Organização lógica:

Funções de entrada, processamento e saída estão claramente separadas no fluxo do código.

Uso de loops e condicionais para validar interações com o usuário.

O código foi estruturado com foco na legibilidade e modularidade, pensando em futuras melhorias — como a automatização da exclusão da última linha da planilha.

Tipo de arquitetura adotada: Pub/Sub (Publicador/Assinante):

O sistema foi projetado com base no padrão Pub/Sub, onde ações como o cadastro de um funcionário funcionam como eventos publicados, e a automação da inserção de dados no formulário atua como assinante desses eventos.

A planilha Funcionarios.xlsx serve como canal de comunicação indireto, permitindo que componentes fiquem desacoplados e possam evoluir separadamente.

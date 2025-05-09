Nomes dos Integrantes

Alfredo Walter Regner Neto - PO

Breno Luiz Raghianti Zein - Srum Master

Carla Machado Miguel - PO

Mathues Rodrigues Antonio - Dev Time

-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Questioário Apresentação

1. Como funciona o pip no Python e como ele foi utilizado no projeto?
O pip é o gerenciador de pacotes do Python, usado para instalar bibliotecas externas. No projeto, ele instalou as bibliotecas openpyxl, pyautogui e pyperclip, essenciais para manipular planilhas e automatizar formulários. A biblioteca time, por ser nativa, não precisou de instalação.

2. Como foi feita a organização dos arquivos e pastas no projeto?
O projeto foi organizado com scripts separados para o menu (atendimento_virtual.py) e automação (automacao_formulario.py), uma planilha Excel com os dados dos funcionários e um arquivo JSON de teste. Essa estrutura facilita a manutenção e a compreensão do código.

3. O que é arquitetura de software e como o grupo aplicou isso no projeto?
Arquitetura de software é a estrutura e organização do sistema. O grupo aplicou isso com separação de responsabilidades entre scripts, modularidade e uso do padrão Pub/Sub, onde a planilha age como meio de comunicação entre os módulos de cadastro e automação.


Introdução do Projeto – Automação de Tarefas Internas

Este projeto tem como objetivo implementar soluções de automação para otimizar processos internos em um ambiente de escritório pequeno. A iniciativa busca reduzir tarefas manuais repetitivas, aumentar a eficiência operacional e melhorar a comunicação entre os colaboradores. Como ponto de partida, será desenvolvido um sistema simples de envio automático de lembretes, que notificará os funcionários sobre datas importantes e outros compromissos recorrentes, garantindo maior organização e pontualidade no dia a dia da empresa.
O que essa aplicação faz?
Este programa foi feito para automatizar o trabalho de buscar faturas em um site, baixar os arquivos dessas faturas e salvar em uma planilha do Excel. Ele usa as bibliotecas Selenium, OpenPyxl e Requests para navegar na internet, criar a planilha e fazer o download dos arquivos.

--------------------------------------------

Por que o código é funcional?
- O programa espera que os elementos do site (como tabelas e botões) fiquem prontos antes de tentar usá-los. Isso evita erros de "elemento não encontrado".
- Ele verifica a data de cada fatura e só baixa aquelas com data igual ou maior que hoje, economizando tempo e recursos.
- Para fazer o download das faturas, ele usa a biblioteca requests, que é mais rápida e leve do que baixar pelo navegador.
- Os dados das faturas (como número, data e link) são salvos diretamente em uma planilha Excel. Isso é feito de forma simples e eficiente.
- Se a pasta onde os arquivos precisam ser salvos não existe, o programa cria ela sozinho.

--------------------------------------

Como usar o programa?
Antes de começar:
Instale o Python 3.9 (ou mais recente).
Instale as bibliotecas necessárias com este comando no terminal:
> pip install selenium openpyxl requests

Baixe o WebDriver do Google Chrome

No código, ajuste o valor da variável pasta_saida para o local onde você quer salvar os arquivos baixados e a planilha.

No terminal ou no editor de código, execute o script:
> python nome_do_arquivo.py

O programa vai abrir o navegador, acessar o site das faturas, pegar as informações e fazer os downloads.

No final, ele cria uma planilha chamada Faturas.xlsx com:

O número da fatura, data da fatura e o link para o arquivo.
Além disso, os arquivos de fatura serão salvos como .jpg na pasta que você configurou.

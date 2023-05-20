# Correios Atende - Extrator de Dados de Postagem

## Descrição

Este é um aplicativo em Python desenvolvido para resolver o problema de falta de geração de relatório no sistema Correios Atende. O sistema não fornece uma funcionalidade para gerar relatórios de postagem, o que dificulta o controle e análise dos dados. O aplicativo extrai os dados de postagem do sistema, permitindo que você os salve em um arquivo Excel para fins de controle e análise.

## Tecnologias Utilizadas

O aplicativo foi desenvolvido utilizando as seguintes tecnologias:

- **Python**: Linguagem de programação utilizada para desenvolver a aplicação.
- **Tkinter**: Biblioteca gráfica do Python para criar a interface do usuário.
- **Pandas**: Biblioteca do Python para manipulação e análise de dados.
- **tkcalendar**: Biblioteca que fornece um widget de calendário para seleção da data.
- **PIL (Python Imaging Library)**: Biblioteca do Python para manipulação de imagens.
- **Selenium**: Framework para automação de testes web.
- **ChromeDriver**: Driver necessário para a execução do Selenium com o navegador Google Chrome.

## Funcionalidades

O aplicativo possui as seguintes funcionalidades:

1. **Seleção de Data**: Permite selecionar a data de pesquisa dos atendimentos de postagem.
2. **Login no Sistema**: Realiza o login no sistema Correios Atende utilizando as credenciais fornecidas.
3. **Pesquisa de Atendimentos**: Realiza a pesquisa dos atendimentos de postagem na data selecionada.
4. **Extração de Dados**: Extrai os dados de postagem de cada atendimento e cria uma lista com as informações relevantes.
5. **Geração do Relatório**: Salva os dados de postagem em um arquivo Excel, na pasta correspondente à data de pesquisa.
6. **Exibição de Resultados**: Apresenta os resultados da extração na interface do aplicativo, incluindo a quantidade total de atendimentos e objetos postados.

## Problema Resolvido

O sistema Correios Atende não gera relatórios de postagem, dificultando o controle e análise dos dados por parte dos usuários. Com o aplicativo Correios Atende - Extrator de Dados de Postagem, é possível coletar os dados diretamente do sistema e salvá-los em um formato estruturado, como um arquivo Excel. Isso permite que os usuários tenham maior controle sobre as postagens realizadas, além de possibilitar análises e tomadas de decisão mais eficientes com base nos dados coletados.



**Observação**: Foi criado um executável da aplicação, para utilização outras máquinas. O arquivo `chromedriver.exe` foi incluído no build juntamente com as dependências necessárias.

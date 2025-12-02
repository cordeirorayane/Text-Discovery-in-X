# Text-Discovery-in-X
O algoritmo TDiX foi desenvolvido com VBA e Selenium utilizando a técnica de Web Scraping para coletar a parte textual das URLs do arquivo .txt de entrada. O código foi criado para sanar um problema de extração de dados citado no artigo **Análise de desempenho do modelo Cassiopeia na clusterização de dados da rede social X**. O TDiX oferece uma solução eficiente e automatizada para a coleta dos textos das URLs provenientes das publicações no X, tornando o processo ideal para atividades como a mineração de dados onde em pesquisas acadêmicas tem-se a construção de uma corpora. 

## Conceitos

Entenda o que é VBA, Selenium, WebDriver e as diferenças entre eles. 


### O que é VBA (Visual Basic for Applications)?

VBA é uma linguagem de programação incorporada em ferramentas da Microsoft, como Excel, Word, Outlook e Access.
Ela permite automatizar tarefas, criar macros, desenvolver funções personalizadas e integrar o Office com outras aplicações.

#### Características principais

* Linguagem baseada no Visual Basic.
* Executada dentro do ambiente do Microsoft Office.
* Muito utilizada para automação de planilhas, manipulação de dados e integração com APIs ou sistemas externos.
* No Excel, por exemplo, pode controlar células, planilhas, gráficos, arquivos e até navegar na web.

---

### O que é o Selenium?

Selenium é um conjunto de ferramentas voltado para automação de navegadores web.
Ele permite que você programe um navegador para abrir páginas, clicar em botões, preencher formulários, extrair dados, etc.

#### Características principais

* Usado principalmente para testes automatizados e web scraping.
* Não é dependente de uma única linguagem: funciona com Python, Java, C#, JavaScript, Ruby e também pode ser integrado ao VBA.
* Simula o comportamento de um usuário humano navegando na web.

---

### O que é o WebDriver?

WebDriver é o componente do Selenium que controla o navegador de verdade.
É ele que envia comandos ao browser e retorna o resultado para o seu código.

#### Como funciona

* Para cada navegador existe um WebDriver específico:

  * ChromeDriver (Google Chrome)
  * GeckoDriver (Firefox)
  * EdgeDriver (Microsoft Edge)
* Ele atua como um “mensageiro”: recebe instruções do Selenium e executa no navegador.

#### Exemplo do fluxo

Seu código → Selenium → WebDriver → Navegador → Resposta

---

### Qual a diferença entre eles?

| Tecnologia    | O que é                                                 | Para que serve                                                                        |
| ------------- | ------------------------------------------------------- | ------------------------------------------------------------------------------------- |
| **VBA**       | Linguagem de programação do Microsoft Office            | Automatizar tarefas dentro do Excel/Outlook/etc.; pode usar Selenium como complemento |
| **Selenium**  | Ferramenta/framework de automação web                   | Criar rotinas que controlam navegadores; permite web scraping e testes                |
| **WebDriver** | Mecanismo que o Selenium usa para controlar o navegador | Executar ações no browser (abrir página, clicar, enviar texto)                        |

#### Resumo 

* VBA é a linguagem que você usa para programar dentro do Office.
* Selenium é o framework que permite automatizar navegadores.
* WebDriver é o “motor” que realmente executa os comandos no browser.

Ou seja:

* Você escreve código em VBA,
* Que chama o Selenium,
* Que usa um WebDriver para controlar o navegador.


## Como rodar o código?

No navegador:

Instalando o Selenium Basic:
1) Faça o download do Selenium no repositório abaixo, clicando no arquivo .exe: 
   https://github.com/florentbr/SeleniumBasic/releases/tag/v2.0.9.0
2) Execute o aplicativo clicando 2 vezes, aperte next e aceite os termos.
3) Clique em Full installation e selecione a opção Compact installation.
4) Clique em next > install e guarde onde ele será instalado.
5) Feche o aplicativo e procure onde ele foi instalado. No windows 11, clique no símbolo do windows no teclado e selecione todos, desça a pesquisa até achar a pasta.
6) Estando na pasta Selenium Basic, iremos agora baixar o WebDriver do navegador em uso para assim adicionar ele a este caminho.
   
Instalando o WebDriver:
1) Baixe a versão estável do WebDriver para o seu navegador. <br>
   Link para as versões do Chrome Driver do Google Chrome: https://googlechromelabs.github.io/chrome-for-testing/ <br>
   Para outros navegores, embaixo do título Platforms Supported by Selenium, clique em Browsers > documentation do navegador que você utiliza e instale o programa. 
   Link: https://www.selenium.dev/downloads/
2) Caso haja uma pasta compactada, descompacte o arquivo.
3) Copie o executável do WebDriver do seu navegador para a pasta onde o Selenium Basic foi baixado.
4) Volte para o Excel.

----------------------------------------------------------------

No Excel:
1) Abra o Excel, clique em Desenvolvedor > Visual Basic.
2) Com a página do Microsoft Visual Basic for Applications aberta, clique em Ferramentas > Referências > marque a caixa Selenium Type Library e aperte em OK.
4) Clique em Pesquisador de Objeto (F2) > Todas as bibliotecas e verifique se o Selenium aparece, caso contrário o código não irá funcionar.
5) Com a página do Microsoft Visual Basic for Applications aberta, clique no nome VBAProject do lado esquerdo da tela.
6) Aperte o lado direito do mouse e selecione inserir > módulo. 
7) Copie o arquivo TDiX.txt e cole na guia aberta do módulo.
8) Adicione no código o caminho onde se encontra o arquivo que contem o .txt com as URLs no seu computador. <br>
   Ex: C:\user\joao\downloads\urls-bradesco-323.txt
9) Adicione no código o caminho da pasta onde irá ser salvo os arquivos com a parte textual das publicações no seu computador. Lembre-se de colocar a barra no final. <br>
   Ex: C:\users\joao\downloads\pastaArquivos\
11) Adicione no código o usuário e senha da conta que você irá utilizar para fazer o Web Scraping.
12) Para executar o código clique no símbolo de play (Executar Sub/UserForm - f9)
13) Se aparecer a mensagem **erro em tempo de execução "-2146232576 (80131700)"** siga os passos do vídeo abaixo: <br>
    Link: https://www.youtube.com/watch?v=6Q9JyUxTw-Y
14) Se ao executar o arquivo aparecer a mensagem **Could not log you in now. Please try again later. g;176464577526156229:-1764645783544:68eMdkWqUk4RXqDydHGOg5vn:1** durante o teste automatizado ao logar com as credenciais na rede social X comente as seguintes linhas: <br>

    driver.Get "https://x.com/i/flow/login" <br>
    Application.Wait (Now + TimeValue("00:00:10")) <br>
    Set usernameInput = driver.FindElementByName("text") <br>
    usernameInput.SendKeys " " <br>
    Application.Wait (Now + TimeValue("00:00:35")) <br>
    Set passwordInput = driver.FindElementByName("password") <br>
    passwordInput.SendKeys " " <br>
    Application.Wait (Now + TimeValue("00:00:04")) <br>
    
Com as alterações acima realizadas o teste irá abrir diretamente as URLs do arquivo .txt ao invés de logar primeiro. Ao realizar testes automatizados, o X pode bloquear automaticamente o acesso por conta de  comportamentos suspeitos como detecção de automação (bot detection) ou user-agent suspeito. Foi utilizado no código a inserção das credenciais, pelo fato de que algumas publicações não são visíveis a pessoas que não estão logadas no X, o que limita a quantidade de dados que podem ser extraídos pelo teste. 

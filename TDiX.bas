Attribute VB_Name = "Módulo1"
Sub PegarConteudoDeURLsDoTxt()
    ' Declaração do objeto WebDriver
    Dim driver As Object
    Dim usernameInput As Object
    Dim passwordInput As Object
    
    ' Declaração das variáveis
    Dim tweetUrl As String
    Dim tweetContent As String
    Dim errorContent As String
    Dim filePath As String
    Dim savePath As String
    Dim fileNum As Integer
    Dim individualFileNum As Integer
    Dim lineText As String
    Dim i As Long
    Dim xpathExpression As String
    Dim errorXPath As String
    Dim individualFilePath As String
    
    ' Caminho completo para o arquivo .txt de URLs
    filePath = " "
    
    ' Caminho da pasta para salvar os tweets (verifique se a barra está no final da pasta)
    savePath = " "
    
    ' Expressão XPath para localizar o conteúdo do tweet
    xpathExpression = "/html/body/div[1]/div/div/div[2]/main/div/div/div/div/div/section/div/div/div[1]/div/div/article/div/div/div[3]/div[1]/div/div"
    
    ' Expressão XPath para detectar a mensagem de erro
    errorXPath = "/html/body/div/div/div/div[2]/main/div/div/div/div/div/div[3]/div/span"
    
    ' Abre o arquivo .txt para leitura
    fileNum = FreeFile
    Open filePath For Input As fileNum
    
    ' Iniciar o Selenium WebDriver
    Set driver = CreateObject("Selenium.ChromeDriver")
    driver.Start "Chrome", tweetUrl
    
    ' Loop para ler cada linha (URL) do arquivo .txt
    i = 1
    Do Until EOF(fileNum)
        Line Input #fileNum, lineText
        tweetUrl = lineText
        
        ' Navega para a URL do tweet
        driver.Get tweetUrl
        
        ' Aguarda o carregamento da página antes de capturar as informações
        Application.Wait Now + TimeValue("00:00:07")
        
        ' Verifica se a mensagem de erro está presente
        On Error Resume Next
        errorContent = driver.FindElementByXPath(errorXPath).Text
        On Error GoTo 0
        
        ' Se a mensagem de erro for encontrada, pula para a próxima URL
        If errorContent <> "" Then
            i = i + 1
        Else
            ' Tenta localizar a div pelo XPath e capturar o texto do tweet
            On Error Resume Next
            tweetContent = driver.FindElementByXPath(xpathExpression).Text
            On Error GoTo 0
            
            ' Se o conteúdo do tweet for encontrado, armazena-o em um arquivo individual
            If tweetContent <> "" Then
                ' Cria o caminho completo para o arquivo individual
                individualFilePath = savePath & "tweet_" & i & ".txt"
                
                ' Atribui um número de arquivo para o arquivo individual e o abre para gravação
                individualFileNum = FreeFile
                Open individualFilePath For Output As individualFileNum
                Print #individualFileNum, tweetContent
                Close individualFileNum
                
                ' Incrementa o contador para o próximo arquivo
                i = i + 1
                tweetContent = ""
            End If
        End If
    Loop
    
    ' Fecha o arquivo de URLs e o navegador depois de ter coletado tudo
    Close fileNum
    driver.Quit
    
    ' Mensagem ao terminar o processo
    MsgBox "Coleta de conteúdo concluída para todas as URLs."
End Sub

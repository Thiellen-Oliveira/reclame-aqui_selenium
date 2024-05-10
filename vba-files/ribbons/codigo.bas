'Attribute VB_Name = "codigo"

Sub extrairDadosReclameAqui()
    
    Dim driveChrome As New ChromeDriver
    
    'Pesquisa da URL do site
    driveChrome.Get "https://www.reclameaqui.com.br/"
    
    Call limparDados
    
        
    'Pesquisa o campo Name para escrever a palavra do produto
    driveChrome.FindElementById("search-input").SendKeys Sheets(1).Range("K1")
    
    Dim teclaTeclado As New Selenium.Keys
    
    'Demos um Enter para pesquisar
    driveChrome.FindElementById("search-input").SendKeys teclaTeclado.Enter
    
            
    'Demos um Enter para clicar no nome da empresa
    driveChrome.FindElementByClass("avatar-letter").SendKeys teclaTeclado.Enter
    
                   
    
    'Demos um Enter para clicar no nome da empresa
    driveChrome.FindElementByXPath("//*[@id=""menu""]/ul/li[2]/a").SendKeys teclaTeclado.Enter
    
    
    
    'Aguarda 8 segundos para que o computador processe as informações e o site também
    Application.Wait (Now + TimeValue("00:00:08"))
    
    'Número de reclamações
    Dim numReclamacoes As String
    ' Extrair o número de reclamações da Amazon nos últimos 6 meses
    numReclamacoes = driveChrome.FindElementByXPath("//*[@id=""newPerformanceCard""]/div[2]/div[1]").Text
    
    
    'Número de reclamações
    'Problemas
    'Título do Problema
    'Status do Problema
    'Clicar no problema
    'Cidade do Problema
    'Data
    'Relacionar qual problema pertence a reclamação
    
    
    Dim linha As Integer
    'Dim numReclamacoes As String
    Dim tituloProblema As String
    Dim stProblema As String
    Dim cidade As String
    Dim data As String
    Dim descricao As String
    Dim problema As String
    
    
    
    linha = 2
    
     'Lista de reclamações
Set listaProblemas = driveChrome.FindElementsByClass("sc-1pe7b5t-0")

' Iterar sobre todos os elementos na lista
For i = 1 To listaProblemas.Count

    ' Recupere a lista de problemas novamente a cada iteração para evitar referências obsoletas
    Set listaProblemas = driveChrome.FindElementsByClass("sc-1pe7b5t-0")
    
    ' Pegue o i-ésimo elemento da lista
    Set informacoesListaProblemas = listaProblemas.Item(i)
   
    Application.Wait (Now + TimeValue("00:00:15"))
   
    'Extrair o número de reclamações da Amazon nos últimos 6 meses
    numReclamacoes = informacoesListaProblemas.FindElementByXPath("//*[@id=""newPerformanceCard""]/div[2]/div[1]/span").Text
    
    ' Extrair o título do problema
    tituloProblema = informacoesListaProblemas.FindElementByClass("sc-1pe7b5t-1").Text
    
    ' Extrair o status do problema
    stProblema = informacoesListaProblemas.FindElementByClass("sc-1pe7b5t-4").Text
    
    ' Clique no título do problema para acessar a página
    informacoesListaProblemas.FindElementByClass("sc-1pe7b5t-1").Click
    
    ' Aguarde um curto período para a página carregar
    Application.Wait (Now + TimeValue("00:00:05"))
    
    ' Extrair a cidade da página do problema
    cidade = driveChrome.FindElementByXPath("//*[@id=""__next""]/div[1]/div[1]/div[3]/main/div/div[2]/div[1]/div[1]/div[3]/div[1]/section/div[1]/span").Text
    
    ' Extrair a data e hora da reclamação
    data = driveChrome.FindElementByXPath("//*[@id=""__next""]/div[1]/div[1]/div[3]/main/div/div[2]/div[1]/div[1]/div[3]/div[1]/section/div[2]/span").Text
    
    ' Extrair a descrição do problema detalhado pelo cliente
    descricao = driveChrome.FindElementByClass("sc-lzlu7c-17").Text
    
    ' Extrair a classificação do problema com o qual o cliente relacionou no momento da reclamação
    problema = driveChrome.FindElementByClass("sc-1dmxdqs-0").Text

    ' Volte para a página anterior
    driveChrome.GoBack     
    
   
        
        
        
    
    'Título
    [A1] = "Numero de Reclamacoes"
    [B1] = "Titulo do Problema"
    [C1] = "Status do Problema"
    [D1] = "Cidade"
    [E1] = "Data e Hora"
    [F1] = "Descricao"
    [G1] = "Problema"
    
    Sheets(1).Cells(linha, 1) = numReclamacoes
    Sheets(1).Cells(linha, 2) = tituloProblema
    Sheets(1).Cells(linha, 3) = stProblema
    Sheets(1).Cells(linha, 4) = cidade
    Sheets(1).Cells(linha, 5) = data
    Sheets(1).Cells(linha, 6) = descricao
    Sheets(1).Cells(linha, 7) = problema
    
    ' Avança para a próxima linha na planilha
    linha = linha + 1
    
    
    
    Next i  
       
    
    
    
    
    Stop
    
    
    
    MsgBox "Dados Extraídos com Sucesso!!!"
    

    
    
End Sub


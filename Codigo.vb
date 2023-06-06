Sub loginsite()
    ' Importar biblioteca Selenium
    Dim cd As New Selenium.ChromeDriver
    
    On Error Resume Next ' Ignorar erros caso o arquivo não exista
    Kill ("C:\Users\%%\Downloads\Routes by vehicle.csv") 'deleta o arquivo
    On Error GoTo 0 ' Voltar ao tratamento normal de erros
   
    'cd.AddArgument "--headless"
    ' Iniciar o driver do Chrome
    cd.Start
    
    
    ' Navegar para a página de login
    cd.Get "https://app.linxio.com/login"

    ' Preencher as informações de login no formulário
    cd.FindElementByCss("[autocomplete='username']").SendKeys "*****"
    cd.FindElementByCss("[type='password']").SendKeys "******"
    cd.FindElementByCss("[type='submit']").Click

    ' Aguardar até que a página seja carregada
    Application.Wait (Now + TimeValue("0:00:04"))

    ' Navegar para a página do relatório
    cd.Get "https://app.linxio.com/client/reports/routes_details/routes_by_vehicle"

    ' Aguardar até que a página seja carregada
    Application.Wait (Now + TimeValue("0:00:08"))

    ' Clicar no botão para abrir o filtro de data
    cd.FindElementByCss("[class='ng-untouched ng-pristine ng-valid']").Click

    ' Aguardar até que o filtro de data seja exibido
    Application.Wait (Now + TimeValue("0:00:02"))
    
    ' Obter a data anterior do dia de hoje
    Dim dataAnterior As Date
    dataAnterior = Date - 1
    
    ' Formatar a data no formato desejado
    Dim dataInicial As String
    Dim dataFinal As String
    dataInicial = Format(dataAnterior, "dd/mm/yyyy") & " 00:00"
    dataFinal = Format(dataAnterior, "dd/mm/yyyy") & " 23:59"
    
    ' Clicar no campo de data para ativá-lo
    cd.FindElementByCss("[id='cdkOverlayTrigger']").Click
    ' Selecionar o texto existente no campo
    ' Injetar o script JavaScript para atualizar o valor do elemento
    Dim script As String
    script = "document.getElementById('cdkOverlayTrigger').value = '" & dataInicial & " - " & dataFinal & "';"
    cd.ExecuteScript script
    Application.Wait (Now + TimeValue("0:00:04"))
    cd.FindElementByCss("[id='cdkOverlayTrigger']").SendKeys "{ENTER}"
    Application.Speech.SpeakCellOnEnter = True

    ' Aguardar até que a página seja carregada
    Application.Wait (Now + TimeValue("0:00:04"))
    
    
    cd.FindElementByCss("[type='submit']").Click
    Application.Wait (Now + TimeValue("0:00:08"))
    
    cd.FindElementByXPath("/html/body/lin-root/mat-sidenav-container/mat-sidenav-content/div/div/div[2]/lin-content/lin-client/div/lin-routes-details/lin-layout-section/div/div[2]/lin-tab-nav-bar/lin-routes-by-vehicle-list/lin-table-settings/button").Click
    Application.Wait (Now + TimeValue("0:00:04"))
    cd.FindElementByXPath("/html/body/lin-root/mat-sidenav-container/mat-sidenav-content/div/div/div[2]/lin-content/lin-client/div/lin-routes-details/lin-layout-section/div/div[2]/lin-tab-nav-bar/lin-routes-by-vehicle-list/lin-table-settings/div[2]/div/div[1]/button[1]").Click
    Application.Wait (Now + TimeValue("0:00:04"))
    

    Dim filePath As String
    filePath = "C:\Users\%%\Downloads\Routes by vehicle.csv" ' Substitua pelo caminho completo do arquivo CSV
    
    ' Abrir o arquivo CSV no Excel
    Workbooks.Open filePath
    
    ' Referenciar a planilha aberta
    Dim wb As Workbook
    Set wb = ActiveWorkbook ' A planilha aberta se torna o workbook ativo
    
    ' Referenciar a primeira planilha no workbook aberto
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1) ' Substitua o número da planilha pelo índice correto
    
    Dim wslimpar As Worksheet
    Set wslimpar = ThisWorkbook.Sheets("Planilha1")
    
    'ws.Cells.Clear ' Limpa todos os dados, incluindo conteúdo, formatação e valores
    
    ' Se você quiser limpar apenas o conteúdo, mantendo a formatação e as fórmulas, use:
     wslimpar.UsedRange.ClearContents
    
    ' Se você quiser limpar apenas os valores, mantendo a formatação e as fórmulas, use:
    ' ws.UsedRange.Value = ""
    
    ' Copiar os dados da planilha aberta para sua planilha atual
    Dim wsAtual As Worksheet
    Set wsAtual = ThisWorkbook.Sheets("Planilha1") ' Substitua pelo nome da sua planilha atual
    
    ws.UsedRange.Copy wsAtual.Range("A1") ' Copiar a partir da célula A1 da planilha atual
    
    ' Fechar o arquivo CSV aberto
    wb.Close SaveChanges:=False
    
    ' Liberar a memória dos objetos
    Set wsAtual = Nothing
    Set ws = Nothing
    Set wb = Nothing
    

End Sub
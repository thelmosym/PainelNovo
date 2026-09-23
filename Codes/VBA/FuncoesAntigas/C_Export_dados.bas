Attribute VB_Name = "C_Export_dados"
Sub ExportarDadosParaNovoArquivoFinal()




 MsgBox "Todos os dados serão exportados para um novo arquivo, selecione o caminho a ser salvo.", vbExclamation
 
    Dim wsMC As Worksheet
    Dim wsARM As Worksheet
    Dim wsDados As Worksheet
    Dim wsFretes As Worksheet
    Dim novoArquivo As Workbook
    Dim caminhoSalvar As String
    Dim nomeArquivo As String
    Dim dataHoraAtual As String
    Dim valorD5 As Variant
    Dim valorJ5 As Variant
    Dim valorG6 As Variant
    Dim valorE6 As Variant
    Dim wsPlanilha1 As Worksheet
    Dim wsNovoMC As Worksheet
    Dim cel As Range
    Dim linksExternos As Variant
    Dim i As Long
    Dim oldUpdateLinks As Boolean

    ' Desativa a atualização de links para evitar caixas de diálogo
    oldUpdateLinks = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    On Error GoTo CleanExit

    ' Referências para as planilhas no arquivo atual
    Set wsMC = ThisWorkbook.Sheets("MC")
    Set wsARM = ThisWorkbook.Sheets("ARM")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    Set wsFretes = ThisWorkbook.Sheets("Fretes")
    
    ' Captura os valores necessários das células
    valorD5 = wsMC.Range("D5").Value
    valorJ5 = wsMC.Range("J5").Value
    valorG6 = ThisWorkbook.Sheets(1).Range("G6").Value ' Assumindo que G6 está na primeira planilha
    valorE6 = ThisWorkbook.Sheets(1).Range("E6").Value
    
    ' Formatar a data e hora no formato especificado
    dataHoraAtual = Format(Now, "yyyy-mm-dd_hh\hmm\mss\s")
    
    ' Nome do arquivo
    nomeArquivo = valorD5 & "-PLA-MemoriaPetrobras-" & valorG6 & valorE6 & "-" & valorJ5 & "#" & dataHoraAtual
    
    ' Caixa de diálogo para o usuário escolher o local para salvar o arquivo
    caminhoSalvar = Application.GetSaveAsFilename(InitialFileName:=nomeArquivo, _
                                                  FileFilter:="Excel Files (*.xlsx), *.xlsx")
    
    ' Verifica se o usuário cancelou a caixa de diálogo
    If caminhoSalvar = "False" Then
        MsgBox "Operação cancelada.", vbInformation, "Cancelado"
        GoTo CleanExit
    End If
    
    ' Cria um novo arquivo Excel
    Set novoArquivo = Workbooks.Add
    
    ' Copia o conteúdo das planilhas para o novo arquivo (preservando formatação e fórmulas)
    wsMC.Copy After:=novoArquivo.Sheets(novoArquivo.Sheets.Count)
    wsARM.Copy After:=novoArquivo.Sheets(novoArquivo.Sheets.Count)
    wsDados.Copy After:=novoArquivo.Sheets(novoArquivo.Sheets.Count)
    wsFretes.Copy After:=novoArquivo.Sheets(novoArquivo.Sheets.Count)
    
    ' Ajusta as fórmulas da planilha "MC" para apontar para as guias do novo arquivo
    Set wsNovoMC = novoArquivo.Sheets("MC")
    For Each cel In wsNovoMC.UsedRange
        If cel.HasFormula Then
            ' Substitui as referências externas (se houver) para apontar para o novo arquivo
            cel.Formula = Replace(cel.Formula, "[" & ThisWorkbook.Name & "]", "")
        End If
    Next cel

    ' Colar apenas os valores das células D5 e J5
    With wsNovoMC
        .Range("D5").Value = valorD5
        .Range("J5").Value = valorJ5
    End With
    
    ' Quebrar todos os vínculos externos no novo arquivo
    On Error Resume Next
    linksExternos = novoArquivo.LinkSources(xlExcelLinks)
    If Not IsEmpty(linksExternos) Then
        For i = LBound(linksExternos) To UBound(linksExternos)
            novoArquivo.BreakLink Name:=linksExternos(i), Type:=xlLinkTypeExcelLinks
        Next i
    End If
    On Error GoTo 0

    ' Verificar e excluir a planilha "Planilha1" se existir
    On Error Resume Next
    Set wsPlanilha1 = novoArquivo.Sheets("Planilha1")
    On Error GoTo 0
    If Not wsPlanilha1 Is Nothing Then
        Application.DisplayAlerts = False ' Desativa os alertas para evitar a confirmação de exclusão
        wsPlanilha1.Delete
        Application.DisplayAlerts = True ' Reativa os alertas
    End If
    
    ' Mover a planilha "MC" para a primeira posição
    wsNovoMC.Move Before:=novoArquivo.Sheets(1)
    
    ' Salva o novo arquivo com o nome e caminho escolhidos
    novoArquivo.SaveAs Filename:=caminhoSalvar, FileFormat:=xlOpenXMLWorkbook
    
    ' Fecha o novo arquivo
    novoArquivo.Close
    
    ' Mensagem de confirmação
    MsgBox "Arquivo exportado e salvo com sucesso!", vbInformation, "Sucesso"

CleanExit:
    ' Reativa as configurações originais
    Application.DisplayAlerts = oldUpdateLinks
    Application.AskToUpdateLinks = True
End Sub






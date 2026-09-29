Attribute VB_Name = "M?dulo1"

' Exporta as principais guias da medicao para um novo arquivo Excel.
' O arquivo gerado recebe um nome baseado nos dados da medicao e e salvo
' no local escolhido pelo usuario.
Sub ExportarDadosParaNovoArquivoAdaptado()
    Dim wsMC As Worksheet
    Dim wsARM As Worksheet
    Dim wsDados As Worksheet
    Dim wsFretes As Worksheet
    Dim novoArquivo As Workbook
    Dim caminhoSalvar As String
    Dim nomeArquivo As String
    Dim dataHoraAtual As String
    Dim valorD5 As String
    Dim valorJ5 As String
    Dim valorG6 As String
    Dim valorE6 As String
    Dim wsPlanilha1 As Worksheet
    Dim wsNovoMC As Worksheet
    Dim cel As Range

    ' Define as guias de origem que serao copiadas para o novo arquivo.
    Set wsMC = ThisWorkbook.Sheets("MC")
    Set wsARM = ThisWorkbook.Sheets("ARM")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    Set wsFretes = ThisWorkbook.Sheets("Fretes")
    
    ' Le os valores usados para montar o nome do arquivo exportado.
    valorD5 = wsMC.Range("D5").Value
    valorJ5 = wsMC.Range("J5").Value
    valorG6 = ThisWorkbook.Sheets(1).Range("G6").Value ' Assumindo que G6 est? na primeira planilha
    valorE6 = ThisWorkbook.Sheets(1).Range("E6").Value
    
    ' Gera um identificador de data e hora para evitar nomes repetidos.
    dataHoraAtual = Format(Now, "yyyy-mm-dd_hh\hmm\mss\s")
    
    ' Monta o nome final usando contrato, periodo e informacoes da primeira guia.
    nomeArquivo = valorD5 & "-PLA-MemoriaPetrobras-" & valorG6 & valorE6 & "-" & valorJ5 & "#" & dataHoraAtual
    
    ' Permite ao usuario escolher a pasta e confirmar o nome do arquivo.
    caminhoSalvar = Application.GetSaveAsFilename(InitialFileName:=nomeArquivo, _
                                                  FileFilter:="Excel Files (*.xlsx), *.xlsx")
    
    ' Encerra o procedimento sem criar arquivo quando a operacao e cancelada.
    If caminhoSalvar = "False" Then
        MsgBox "Opera??o cancelada.", vbInformation, "Cancelado"
        Exit Sub
    End If
    
    ' Cria o workbook que recebera as copias das guias de origem.
    Set novoArquivo = Workbooks.Add
    
    ' Copia MC, ARM, Dados e Fretes preservando formulas e formatacao.
    wsMC.Copy After:=novoArquivo.Sheets(novoArquivo.Sheets.Count)
    wsARM.Copy After:=novoArquivo.Sheets(novoArquivo.Sheets.Count)
    wsDados.Copy After:=novoArquivo.Sheets(novoArquivo.Sheets.Count)
    wsFretes.Copy After:=novoArquivo.Sheets(novoArquivo.Sheets.Count)
    
    ' Remove das formulas da guia MC a referencia ao arquivo original.
    ' Assim, as formulas passam a apontar para as guias copiadas no novo workbook.
    Set wsNovoMC = novoArquivo.Sheets("MC")
    For Each cel In wsNovoMC.UsedRange
        If cel.HasFormula Then
            ' Converte uma referencia externa em referencia interna ao novo arquivo.
            cel.Formula = Replace(cel.Formula, "[" & ThisWorkbook.Name & "]", "")
        End If
    Next cel
    
    ' Remove a guia vazia criada automaticamente pelo Workbooks.Add, se existir.
    On Error Resume Next
    Set wsPlanilha1 = novoArquivo.Sheets("Planilha1")
    On Error GoTo 0
    If Not wsPlanilha1 Is Nothing Then
        Application.DisplayAlerts = False ' Desativa os alertas para evitar a confirma??o de exclus?o
        wsPlanilha1.Delete
        Application.DisplayAlerts = True ' Reativa os alertas
    End If
    
    ' Deixa a guia principal MC como a primeira guia do arquivo exportado.
    wsNovoMC.Move Before:=novoArquivo.Sheets(1)
    
    ' Salva no formato XLSX, sem macros, usando o caminho escolhido pelo usuario.
    novoArquivo.SaveAs Filename:=caminhoSalvar, FileFormat:=xlOpenXMLWorkbook
    
    ' Fecha o arquivo exportado depois que o salvamento foi concluido.
    novoArquivo.Close
    
    ' Confirma ao usuario que a exportacao terminou com sucesso.
    MsgBox "Arquivo exportado e salvo com sucesso!", vbInformation, "Sucesso"
End Sub



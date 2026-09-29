Attribute VB_Name = "modGerarArquivos"
Option Explicit

'====================================================================================================
' MÓDULO: modGerarArquivos
' OBJETIVO: Gera 4 arquivos finais limpos (um por contrato: PA-LT2, IRON-LT1-RJ, IRON-LT1-SP,
'           IRON-LT1-ES), cada um com 4 abas fixas (MC, ARM, DADOS, FRETE), copiando os dados
'           das respectivas abas de origem da planilha principal ("{TIPO}_{CONTRATO}").
'
' PADRÕES DE ENGENHARIA APLICADOS:
'   - Tipagem estrita 64-bit com variáveis Long para todas as linhas e índices de planilha.
'   - Isolamento e restauração garantida de estado do Excel (Calculation, EnableEvents, ScreenUpdating).
'   - Cópia atômica de larguras de coluna (xlPasteColumnWidths), eliminando loops lentos coluna a coluna.
'   - Conversão de fórmulas em valores na aba MC para eliminação de vínculos externos.
'   - Gravação de Hyperlinks na aba Painel (K7:K10) com carimbo de data/hora no nome do arquivo.
'====================================================================================================

Private Const MSG_TITULO As String = "Geração de Arquivos por Contrato"
Private Const NOME_PLANILHA_LINKS As String = "Painel"
Private Const COLUNA_CELULA_LINK As String = "K"
Private Const LINHA_INICIAL_LINK As Long = 7
Private Const NOME_SUBPASTA_DESTINO As String = "MEDIÇÃO"

' GUID do rótulo de confidencialidade "Pública" do tenant da Petrobras
Private Const PUBLICO_LABEL_ID As String = "140b9f7d-8e3a-482f-9702-4b7ffc40985a"
Private Const MSO_ASSIGNMENT_METHOD_STANDARD As Long = 1

Private Function ListaTipos() As Variant
    ListaTipos = Array("MC", "ARM", "DADOS", "FRETE")
End Function

Private Function ListaContratos() As Variant
    ListaContratos = Array("PA-LT2", "IRON-LT1-RJ", "IRON-LT1-SP", "IRON-LT1-ES")
End Function

Private Function ListaNumerosContrato() As Variant
    ListaNumerosContrato = Array("4600687006", "4600686987", "4600686987", "4600686987")
End Function

Private Function ListaSufixosEstado() As Variant
    ListaSufixosEstado = Array("BA", "", "", "")
End Function

'----------------------------------------------------------------------------------------------------
' PONTO DE ENTRADA PÚBLICO (Associado ao botão "Gerar Arquivos" no Painel)
'----------------------------------------------------------------------------------------------------
Public Sub GerarArquivosPorContrato()
    Dim wbOrigem        As Workbook
    Dim wsLinks         As Worksheet
    Dim sPastaDestino   As String
    Dim vContratos      As Variant
    Dim vTipos          As Variant
    Dim vNumeros        As Variant
    Dim vSufixos        As Variant
    Dim iContrato       As Long
    Dim iArquivosOK     As Long
    Dim sResumoFinal    As String
    
    Dim appCalc         As XlCalculation
    Dim bEvents         As Boolean
    Dim bScreen         As Boolean
    Dim bAlerts         As Boolean
    
    Set wbOrigem = ThisWorkbook
    
    If Len(wbOrigem.Path) = 0 Then
        MsgBox "Procedimento interrompido." & vbNewLine & vbNewLine & _
               "Este arquivo ainda não foi salvo em disco, portanto não é possível " & _
               "determinar a pasta de destino automaticamente." & vbNewLine & _
               "Salve este arquivo (Ctrl+S) e tente novamente.", _
               vbExclamation, MSG_TITULO
        Exit Sub
    End If
    
    vContratos = ListaContratos()
    vTipos = ListaTipos()
    vNumeros = ListaNumerosContrato()
    vSufixos = ListaSufixosEstado()
    
    sPastaDestino = ObterOuCriarSubpasta(wbOrigem.Path, NOME_SUBPASTA_DESTINO)
    If Len(sPastaDestino) = 0 Then
        MsgBox "Procedimento interrompido. Não foi possível criar ou acessar " & _
               "a pasta '" & NOME_SUBPASTA_DESTINO & "'.", vbCritical, MSG_TITULO
        Exit Sub
    End If
    
    Set wsLinks = ObterOuCriarPlanilha(wbOrigem, NOME_PLANILHA_LINKS)
    
    If MsgBox("Serão gerados " & (UBound(vContratos) + 1) & " cadernos contratuais na pasta:" & vbNewLine & _
              sPastaDestino & vbNewLine & vbNewLine & _
              "Contratos: " & Join(vContratos, ", ") & vbNewLine & _
              "Abas por arquivo: " & Join(vTipos, ", ") & vbNewLine & vbNewLine & _
              "Deseja iniciar a geração agora?", _
              vbYesNo + vbQuestion, MSG_TITULO) = vbNo Then
        MsgBox "Procedimento cancelado.", vbInformation, MSG_TITULO
        Exit Sub
    End If
    
    '-- Captura e isola o estado do ambiente do Excel
    appCalc = Application.Calculation
    bEvents = Application.EnableEvents
    bScreen = Application.ScreenUpdating
    bAlerts = Application.DisplayAlerts
    
    On Error GoTo TratarErro
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    iArquivosOK = 0
    sResumoFinal = ""
    
    For iContrato = LBound(vContratos) To UBound(vContratos)
        Dim sContrato As String, sNumero As String, sSufixo As String
        Dim sCaminhoGerado As String
        
        sContrato = CStr(vContratos(iContrato))
        sNumero = CStr(vNumeros(iContrato))
        sSufixo = CStr(vSufixos(iContrato))
        
        Application.StatusBar = "Gerando caderno " & (iContrato + 1) & " de " & _
                                (UBound(vContratos) + 1) & ": " & sContrato & "..."
        
        sCaminhoGerado = GerarArquivoDoContrato(wbOrigem, sContrato, sNumero, sSufixo, vTipos, sPastaDestino)
        
        If Len(sCaminhoGerado) > 0 Then
            iArquivosOK = iArquivosOK + 1
            sResumoFinal = sResumoFinal & "[OK] " & sContrato & " -> " & sCaminhoGerado & vbNewLine
            EscreverLinkArquivo wsLinks, iContrato, sCaminhoGerado
        Else
            sResumoFinal = sResumoFinal & "[FALHA] " & sContrato & " (Falha na geração)" & vbNewLine
            EscreverErroLink wsLinks, iContrato, sContrato
        End If
    Next iContrato

SairRotina:
    Application.Calculation = appCalc
    Application.EnableEvents = bEvents
    Application.DisplayAlerts = bAlerts
    Application.ScreenUpdating = bScreen
    Application.StatusBar = False
    
    If iArquivosOK > 0 Then
        MsgBox "Processamento concluído com sucesso." & vbNewLine & vbNewLine & _
               "Arquivos gerados: " & iArquivosOK & " de " & (UBound(vContratos) + 1) & vbNewLine & vbNewLine & _
               sResumoFinal & vbNewLine & _
               "Links gravados em: " & wsLinks.Name & "!" & COLUNA_CELULA_LINK & LINHA_INICIAL_LINK & _
               ":" & COLUNA_CELULA_LINK & (LINHA_INICIAL_LINK + UBound(vContratos)), _
               vbInformation, MSG_TITULO
    End If
    Exit Sub

TratarErro:
    MsgBox "Erro inesperado durante a geração dos arquivos:" & vbNewLine & _
           Err.Number & " - " & Err.Description, vbCritical, MSG_TITULO
    Resume SairRotina
End Sub

'----------------------------------------------------------------------------------------------------
' FUNÇÃO DE GERAÇÃO INDIVIDUAL DE ARQUIVO POR CONTRATO
'----------------------------------------------------------------------------------------------------
Private Function GerarArquivoDoContrato(ByVal wbOrigem As Workbook, ByVal sContrato As String, _
                                        ByVal sNumero As String, ByVal sSufixo As String, _
                                        ByVal vTipos As Variant, ByVal sPastaDestino As String) As String
    Dim wbDestino       As Workbook
    Dim wsOrigem        As Worksheet
    Dim wsDestino       As Worksheet
    Dim iTipo           As Long
    Dim sTipo           As String
    Dim sNomeAbaOrig    As String
    Dim iCopiadas       As Long
    Dim sFaltantes      As String
    Dim sNomeArquivo    As String
    Dim sCaminhoCompleto As String
    Dim sSufixoComHifen As String
    Dim rngArea         As Range
    
    On Error GoTo TratarErroLocal
    GerarArquivoDoContrato = ""
    
    Set wbDestino = Workbooks.Add(xlWBATWorksheet)
    Do While wbDestino.Worksheets.Count > 1
        wbDestino.Worksheets(wbDestino.Worksheets.Count).Delete
    Loop
    
    iCopiadas = 0
    sFaltantes = ""
    
    For iTipo = LBound(vTipos) To UBound(vTipos)
        sTipo = CStr(vTipos(iTipo))
        sNomeAbaOrig = sTipo & "_" & sContrato
        
        Set wsOrigem = ObterPlanilha(wbOrigem, sNomeAbaOrig)
        
        If iTipo = LBound(vTipos) Then
            Set wsDestino = wbDestino.Worksheets(1)
        Else
            Set wsDestino = wbDestino.Worksheets.Add(After:=wbDestino.Worksheets(wbDestino.Worksheets.Count))
        End If
        
        wsDestino.Name = sTipo
        
        If wsOrigem Is Nothing Then
            sFaltantes = sFaltantes & " - Aba '" & sNomeAbaOrig & "' não encontrada." & vbNewLine
            wsDestino.Range("A1").Value = "Aba de origem '" & sNomeAbaOrig & "' não encontrada."
        Else
            If sTipo = "MC" Then
                CopiarComFormatacaoCompleta wsOrigem.UsedRange, wsDestino
            Else
                Set rngArea = ObterAreaRealDados(wsOrigem)
                CopiarValoresEFormatos rngArea, wsDestino
            End If
            iCopiadas = iCopiadas + 1
        End If
    Next iTipo
    
    If iCopiadas = 0 Then
        wbDestino.Close SaveChanges:=False
        Exit Function
    End If
    
    AplicarPerfilPublico wbDestino
    
    If Len(sSufixo) > 0 Then
        sSufixoComHifen = "-" & sSufixo
    Else
        sSufixoComHifen = ""
    End If
    
    sNomeArquivo = sNumero & "-PLA-MemoriaPetrobras-" & sContrato & sSufixoComHifen & "_" & _
                   Format(Now, "ddmmyyyy_hhmmss") & ".xlsx"
                   
    sCaminhoCompleto = sPastaDestino & Application.PathSeparator & sNomeArquivo
    
    wbDestino.SaveAs Filename:=sCaminhoCompleto, FileFormat:=xlOpenXMLWorkbook
    wbDestino.Close SaveChanges:=False
    
    GerarArquivoDoContrato = sCaminhoCompleto
    Exit Function

TratarErroLocal:
    On Error Resume Next
    If Not wbDestino Is Nothing Then wbDestino.Close SaveChanges:=False
    GerarArquivoDoContrato = ""
End Function

'----------------------------------------------------------------------------------------------------
' ROTINAS DE CÓPIA OTIMIZADA E FORMATAÇÃO
'----------------------------------------------------------------------------------------------------
Private Sub CopiarValoresEFormatos(ByVal rngOrigem As Range, ByVal wsDestino As Worksheet)
    If rngOrigem Is Nothing Then Exit Sub
    On Error GoTo TratarErroLocal
    
    rngOrigem.Copy
    wsDestino.Range("A1").PasteSpecial Paste:=xlPasteValues
    wsDestino.Range("A1").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    ReplicarLargurasColunas rngOrigem, wsDestino
    Exit Sub

TratarErroLocal:
    Application.CutCopyMode = False
End Sub

Private Sub CopiarComFormatacaoCompleta(ByVal rngOrigem As Range, ByVal wsDestino As Worksheet)
    Dim rngDestinoFinal As Range
    If rngOrigem Is Nothing Then Exit Sub
    On Error GoTo TratarErroLocal
    
    rngOrigem.Copy
    wsDestino.Range("A1").PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False
    
    ' Congela fórmulas em valores estáticos
    Set rngDestinoFinal = wsDestino.Range("A1").Resize(rngOrigem.Rows.Count, rngOrigem.Columns.Count)
    rngDestinoFinal.Value = rngDestinoFinal.Value
    
    ReplicarLargurasColunas rngOrigem, wsDestino
    ReplicarAlturasLinhas rngOrigem, wsDestino
    Exit Sub

TratarErroLocal:
    Application.CutCopyMode = False
    On Error Resume Next
    rngOrigem.Copy
    wsDestino.Range("A1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    On Error GoTo 0
End Sub

Private Sub ReplicarLargurasColunas(ByVal rngOrigem As Range, ByVal wsDestino As Worksheet)
    On Error Resume Next
    ' Cópia atômica via clipboard interno do Excel (muito mais rápida que loop de colunas)
    rngOrigem.Rows(1).Copy
    wsDestino.Range("A1").PasteSpecial Paste:=xlPasteColumnWidths
    Application.CutCopyMode = False
    
    ' Fallback caso a cópia atômica não seja suportada
    If Err.Number <> 0 Then
        Err.Clear
        Dim iCol As Long
        For iCol = 1 To rngOrigem.Columns.Count
            wsDestino.Columns(iCol).ColumnWidth = rngOrigem.Worksheet.Columns( _
                rngOrigem.Column + iCol - 1).ColumnWidth
        Next iCol
    End If
    On Error GoTo 0
End Sub

Private Sub ReplicarAlturasLinhas(ByVal rngOrigem As Range, ByVal wsDestino As Worksheet)
    Dim iLin As Long
    For iLin = 1 To rngOrigem.Rows.Count
        wsDestino.Rows(iLin).RowHeight = rngOrigem.Worksheet.Rows( _
            rngOrigem.Row + iLin - 1).RowHeight
    Next iLin
End Sub

Private Function ObterAreaRealDados(ByVal ws As Worksheet) As Range
    Dim ultimaColuna As Long
    Dim ultimaLinha As Long
    
    If ws.Cells(1, ws.Columns.Count).Value <> "" Then
        ultimaColuna = ws.Columns.Count
    Else
        ultimaColuna = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    End If
    
    If ws.Cells(ws.Rows.Count, 1).Value <> "" Then
        ultimaLinha = ws.Rows.Count
    Else
        ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    End If
    
    If ultimaColuna < 1 Then ultimaColuna = 1
    If ultimaLinha < 1 Then ultimaLinha = 1
    
    Set ObterAreaRealDados = ws.Range(ws.Cells(1, 1), ws.Cells(ultimaLinha, ultimaColuna))
End Function

'----------------------------------------------------------------------------------------------------
' GERENCIAMENTO DE LINKS NA ABA PAINEL
'----------------------------------------------------------------------------------------------------
Private Function ObterLinhaCelulaLink(ByVal iIndiceContrato As Long) As Long
    ObterLinhaCelulaLink = LINHA_INICIAL_LINK + iIndiceContrato
End Function

Private Sub EscreverLinkArquivo(ByVal ws As Worksheet, ByVal iIndiceContrato As Long, _
                                ByVal sCaminhoCompleto As String)
    Dim iLinha As Long
    Dim rngCelula As Range
    Dim sNomeArq As String
    
    On Error GoTo TratarErroLocal
    iLinha = ObterLinhaCelulaLink(iIndiceContrato)
    Set rngCelula = ws.Range(COLUNA_CELULA_LINK & iLinha)
    sNomeArq = Mid(sCaminhoCompleto, InStrRev(sCaminhoCompleto, Application.PathSeparator) + 1)
    
    rngCelula.Hyperlinks.Delete
    rngCelula.ClearContents
    rngCelula.Interior.ColorIndex = xlColorIndexNone
    rngCelula.Font.ColorIndex = xlColorIndexAutomatic
    
    ws.Hyperlinks.Add Anchor:=rngCelula, Address:=sCaminhoCompleto, TextToDisplay:=sNomeArq
    Exit Sub

TratarErroLocal:
    On Error Resume Next
    rngCelula.Value = sCaminhoCompleto
    On Error GoTo 0
End Sub

Private Sub EscreverErroLink(ByVal ws As Worksheet, ByVal iIndiceContrato As Long, ByVal sContrato As String)
    Dim iLinha As Long
    Dim rngCelula As Range
    On Error Resume Next
    
    iLinha = ObterLinhaCelulaLink(iIndiceContrato)
    Set rngCelula = ws.Range(COLUNA_CELULA_LINK & iLinha)
    
    rngCelula.Hyperlinks.Delete
    rngCelula.Value = "Falha ao gerar arquivo de " & sContrato
    rngCelula.Font.Color = RGB(220, 38, 38)
    On Error GoTo 0
End Sub

'----------------------------------------------------------------------------------------------------
' UTILITÁRIOS DE SISTEMA DE ARQUIVOS E PLANILHAS
'----------------------------------------------------------------------------------------------------
Private Function ObterOuCriarSubpasta(ByVal sPastaBase As String, ByVal sNomeSubpasta As String) As String
    Dim sCaminhoCompleto As String
    On Error GoTo TratarErroLocal
    
    sCaminhoCompleto = sPastaBase & Application.PathSeparator & sNomeSubpasta
    If Len(Dir(sCaminhoCompleto, vbDirectory)) = 0 Then
        MkDir sCaminhoCompleto
    End If
    
    ObterOuCriarSubpasta = sCaminhoCompleto
    Exit Function

TratarErroLocal:
    ObterOuCriarSubpasta = ""
End Function

Private Function ObterOuCriarPlanilha(ByVal wb As Workbook, ByVal nomeAba As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nomeAba)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = nomeAba
    End If
    
    Set ObterOuCriarPlanilha = ws
End Function

Private Function ObterPlanilha(ByVal wb As Workbook, ByVal nomeAba As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nomeAba)
    On Error GoTo 0
    Set ObterPlanilha = ws
End Function

Private Sub AplicarPerfilPublico(ByVal wb As Workbook)
    If Len(PUBLICO_LABEL_ID) = 0 Then Exit Sub
    On Error GoTo TratarErroLocal
    
    Dim lbl As Object
    Set lbl = wb.SensitivityLabel.CreateLabelInfo
    lbl.LabelId = PUBLICO_LABEL_ID
    lbl.AssignmentMethod = MSO_ASSIGNMENT_METHOD_STANDARD
    lbl.Justification = "Rótulo aplicado automaticamente na geração do arquivo."
    
    wb.SensitivityLabel.SetLabel lbl, lbl
    Exit Sub

TratarErroLocal:
    ' Falha não-bloqueante caso o ambiente Office não possua AIP/MIP habilitado
End Sub

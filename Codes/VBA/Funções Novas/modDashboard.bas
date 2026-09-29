Attribute VB_Name = "modDashboard"
Option Explicit

'====================================================================================================
' MÓDULO: modDashboard
' OBJETIVO: Constrói e atualiza dinamicamente a aba executiva 'DASHBOARD' com cards de indicadores,
'           resumo volumétrico consolidado por contrato e alertas de faturamento.
' PADRÕES:  Processamento otimizado com ScreenUpdating/Calculation controlado e tratamento seguro.
'====================================================================================================

Private Const NOME_DASHBOARD    As String = "DASHBOARD"
Private Const NOME_DADOS        As String = "DADOS"
Private Const NOME_PAINEL       As String = "Painel"

' Cores corporativas para identidade visual
Private Const COR_VERDE         As Long = 54784
Private Const COR_VERDE_ESCURO  As Long = 37888
Private Const COR_VERDE_CLARO   As Long = 13434828
Private Const COR_CINZA         As Long = 15921906
Private Const COR_CINZA_TEXTO   As Long = 5921370
Private Const COR_AMBAR         As Long = 49407
Private Const COR_VERMELHO      As Long = 255

'----------------------------------------------------------------------------------------------------
' PONTOS DE ENTRADA PÚBLICOS
'----------------------------------------------------------------------------------------------------
Public Sub CriarDashboard()
    Dim ws              As Worksheet
    Dim wsDados         As Worksheet
    Dim ultimaLinha     As Long
    Dim colContrato     As Long, colSolicitada As Long, colAtendida As Long
    Dim colQExec        As Long, colPrazo As Long, colFechamento As Long
    Dim totalRegistros  As Long
    Dim totalSolicitada As Double, totalAtendida As Double, totalQExec As Double
    Dim totalForaPrazo  As Long
    
    Dim appCalc         As XlCalculation
    Dim bEvents         As Boolean
    Dim bScreen         As Boolean

    On Error GoTo TratarErro
    
    appCalc = Application.Calculation
    bEvents = Application.EnableEvents
    bScreen = Application.ScreenUpdating
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set ws = ObterOuCriarDashboard()
    Set wsDados = ObterPlanilha(NOME_DADOS)

    PrepararFolha ws
    ConstruirCabecalho ws
    ConstruirAcoes ws
    ConstruirFiltros ws
    ConstruirCards ws
    ConstruirResumoContratos ws
    ConstruirAlertas ws
    ConstruirArquivos ws
    ConstruirRodape ws

    If Not wsDados Is Nothing Then
        ultimaLinha = UltimaLinhaDados(wsDados)
        colContrato = EncontrarCabecalho(wsDados, "Contrato")
        colSolicitada = EncontrarCabecalho(wsDados, "Qtd. Solicitada", "Qtd Solicitada", "Qtd._x000a_Solicitada")
        colAtendida = EncontrarCabecalho(wsDados, "Qtd. Atendida", "Qtd Atendida", "Qtd._x000a_Atendida")
        colQExec = EncontrarCabecalho(wsDados, "QExec")
        colPrazo = EncontrarCabecalho(wsDados, "Data SLA", "Prazo SLA")
        colFechamento = EncontrarCabecalho(wsDados, "D. fechamento", "Data fechamento", "Data Fechamento")

        If ultimaLinha > 1 Then
            totalRegistros = ultimaLinha - 1
            totalSolicitada = SomarColuna(wsDados, colSolicitada, 2, ultimaLinha)
            totalAtendida = SomarColuna(wsDados, colAtendida, 2, ultimaLinha)
            totalQExec = SomarColuna(wsDados, colQExec, 2, ultimaLinha)
            totalForaPrazo = ContarForaPrazo(wsDados, colPrazo, colFechamento, 2, ultimaLinha)
        End If
    End If

    ws.Range("B10").Value = totalRegistros
    ws.Range("D10").Value = totalSolicitada
    ws.Range("F10").Value = totalAtendida
    ws.Range("H10").Value = totalQExec
    ws.Range("J10").Value = "4 de 4"
    ws.Range("L10").Value = totalForaPrazo
    AtualizarResumoContratos ws, wsDados
    AtualizarAlertas ws, totalRegistros, totalSolicitada - totalAtendida, totalForaPrazo

SairRotina:
    Application.Calculation = appCalc
    Application.EnableEvents = bEvents
    Application.ScreenUpdating = bScreen
    
    On Error Resume Next
    ws.Activate
    On Error GoTo 0
    
    MsgBox "A aba DASHBOARD foi criada/atualizada com sucesso.", vbInformation, "Dashboard"
    Exit Sub

TratarErro:
    MsgBox "Não foi possível criar o dashboard:" & vbCrLf & Err.Number & " - " & Err.Description, vbCritical, "Dashboard"
    Resume SairRotina
End Sub

Public Sub AtualizarDashboard()
    CriarDashboardSemMensagem
End Sub

Private Sub CriarDashboardSemMensagem()
    Dim ws              As Worksheet
    Dim wsDados         As Worksheet
    Dim ultimaLinha     As Long
    Dim colSolicitada   As Long, colAtendida As Long, colQExec As Long
    Dim colPrazo        As Long, colFechamento As Long
    Dim totalRegistros  As Long
    Dim totalSolicitada As Double, totalAtendida As Double, totalQExec As Double
    Dim totalForaPrazo  As Long
    
    Dim appCalc         As XlCalculation
    Dim bEvents         As Boolean
    Dim bScreen         As Boolean

    Set ws = ObterOuCriarDashboard()
    Set wsDados = ObterPlanilha(NOME_DADOS)
    If wsDados Is Nothing Then Exit Sub

    appCalc = Application.Calculation
    bEvents = Application.EnableEvents
    bScreen = Application.ScreenUpdating
    
    On Error GoTo SairSilencioso
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ultimaLinha = UltimaLinhaDados(wsDados)
    colSolicitada = EncontrarCabecalho(wsDados, "Qtd. Solicitada", "Qtd Solicitada", "Qtd._x000a_Solicitada")
    colAtendida = EncontrarCabecalho(wsDados, "Qtd. Atendida", "Qtd Atendida", "Qtd._x000a_Atendida")
    colQExec = EncontrarCabecalho(wsDados, "QExec")
    colPrazo = EncontrarCabecalho(wsDados, "Data SLA", "Prazo SLA")
    colFechamento = EncontrarCabecalho(wsDados, "D. fechamento", "Data fechamento", "Data Fechamento")

    If ultimaLinha > 1 Then
        totalRegistros = ultimaLinha - 1
        totalSolicitada = SomarColuna(wsDados, colSolicitada, 2, ultimaLinha)
        totalAtendida = SomarColuna(wsDados, colAtendida, 2, ultimaLinha)
        totalQExec = SomarColuna(wsDados, colQExec, 2, ultimaLinha)
        totalForaPrazo = ContarForaPrazo(wsDados, colPrazo, colFechamento, 2, ultimaLinha)
    End If

    ws.Range("B10").Value = totalRegistros
    ws.Range("D10").Value = totalSolicitada
    ws.Range("F10").Value = totalAtendida
    ws.Range("H10").Value = totalQExec
    ws.Range("L10").Value = totalForaPrazo
    AtualizarResumoContratos ws, wsDados
    AtualizarAlertas ws, totalRegistros, totalSolicitada - totalAtendida, totalForaPrazo
    ws.Range("B4").Value = "Periodo de medicao: " & TextoPeriodo() & "    |    Ultima atualizacao: " & Format(Now, "dd/mm/yyyy hh:mm")

SairSilencioso:
    Application.Calculation = appCalc
    Application.EnableEvents = bEvents
    Application.ScreenUpdating = bScreen
End Sub

'----------------------------------------------------------------------------------------------------
' ROTINAS DE CONSTRUÇÃO DE LAYOUT
'----------------------------------------------------------------------------------------------------
Private Function ObterOuCriarDashboard() As Worksheet
    On Error Resume Next
    Set ObterOuCriarDashboard = ThisWorkbook.Worksheets(NOME_DASHBOARD)
    On Error GoTo 0

    If ObterOuCriarDashboard Is Nothing Then
        Set ObterOuCriarDashboard = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ObterOuCriarDashboard.Name = NOME_DASHBOARD
    End If
End Function

Private Function ObterPlanilha(ByVal nome As String) As Worksheet
    On Error Resume Next
    Set ObterPlanilha = ThisWorkbook.Worksheets(nome)
    On Error GoTo 0
End Function

Private Sub PrepararFolha(ByVal ws As Worksheet)
    Dim shp As Shape

    ws.Cells.Clear
    For Each shp In ws.Shapes
        shp.Delete
    Next shp

    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10
    ws.Cells.Interior.Color = RGB(244, 246, 248)
    
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 3
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 3
    ws.Columns("F").ColumnWidth = 18
    ws.Columns("G").ColumnWidth = 3
    ws.Columns("H").ColumnWidth = 18
    ws.Columns("I").ColumnWidth = 3
    ws.Columns("J").ColumnWidth = 18
    ws.Columns("K").ColumnWidth = 3
    ws.Columns("L").ColumnWidth = 18
    ws.Columns("M").ColumnWidth = 3
    ws.Columns("N").ColumnWidth = 20
    
    ws.Rows("1:40").RowHeight = 18
    ws.Rows("1:3").RowHeight = 24
    ws.Range("A1:N40").VerticalAlignment = xlCenter
    
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
End Sub

Private Sub ConstruirCabecalho(ByVal ws As Worksheet)
    With ws.Range("B1:N3")
        .Merge
        .Interior.Color = RGB(0, 133, 66)
        .Font.Color = vbWhite
        .Font.Bold = True
        .Font.Size = 16
        .Value = "PAINEL DE CONTROLE | MEMORIA PETROBRAS"
        .HorizontalAlignment = xlLeft
    End With

    ws.Range("B4:N4").Merge
    ws.Range("B4").Value = "Periodo de medicao: " & TextoPeriodo() & "    |    Ultima atualizacao: " & Format(Now, "dd/mm/yyyy hh:mm")
    ws.Range("B4").Font.Color = RGB(51, 65, 85)
    ws.Range("B4").Font.Bold = True
    ws.Range("B4:N4").Interior.Color = RGB(226, 247, 235)
    ws.Range("B4:N4").Borders.LineStyle = xlContinuous
    ws.Range("B4:N4").Borders.Color = RGB(187, 222, 196)
End Sub

Private Sub ConstruirAcoes(ByVal ws As Worksheet)
    CriarBotao ws, "B6", "Atualizar painel", "AtualizarTodasConsultasPowerQuery", RGB(0, 133, 66), 125
    CriarBotao ws, "D6", "Gerar arquivos", "GerarArquivosPorContrato", RGB(51, 65, 85), 115
    CriarBotao ws, "F6", "Atualizar indicadores", "AtualizarDashboard", RGB(100, 116, 139), 135
    CriarBotao ws, "H6", "Abrir pasta MEDICAO", "AbrirPastaMedicao", RGB(100, 116, 139), 145
    CriarBotao ws, "J6", "Ver detalhes", "MostrarResumoDashboard", RGB(217, 119, 6), 105
End Sub

Private Sub ConstruirFiltros(ByVal ws As Worksheet)
    ws.Range("B8:N8").Merge
    ws.Range("B8").Value = "FILTROS E CONTEXTO OPERACIONAL"
    EstilizarTituloSecao ws.Range("B8:N8")

    ws.Range("B9").Value = "Periodo"
    ws.Range("D9").Value = "Contrato"
    ws.Range("F9").Value = "Situacao"
    ws.Range("H9").Value = "FDM"
    ws.Range("J9").Value = "Atividade"
    ws.Range("L9").Value = "Localidade"
    ws.Range("B9:N9").Font.Bold = True
    ws.Range("B9:N9").Font.Color = COR_CINZA_TEXTO
    ws.Range("B9:N9").HorizontalAlignment = xlCenter
End Sub

Private Sub ConstruirCards(ByVal ws As Worksheet)
    Dim rotulos As Variant
    Dim colunas As Variant
    Dim i       As Long

    rotulos = Array("REGISTROS", "QTD. SOLICITADA", "QTD. ATENDIDA", "QEXEC TOTAL", "CONTRATOS", "FORA DO PRAZO")
    colunas = Array("B", "D", "F", "H", "J", "L")

    For i = LBound(rotulos) To UBound(rotulos)
        With ws.Range(colunas(i) & "10:" & colunas(i) & "12")
            .Merge
            .Interior.Color = IIf(i = 5, RGB(255, 247, 224), vbWhite)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = IIf(i = 5, RGB(245, 158, 11), RGB(203, 213, 225))
        End With
        ws.Range(colunas(i) & "9").Value = rotulos(i)
        ws.Range(colunas(i) & "10").HorizontalAlignment = xlCenter
        ws.Range(colunas(i) & "10").Font.Size = 18
        ws.Range(colunas(i) & "10").Font.Bold = True
        ws.Range(colunas(i) & "10").NumberFormat = "#,##0.00"
    Next i

    ws.Range("B10").NumberFormat = "#,##0"
    ws.Range("J10").NumberFormat = "@"
    ws.Range("L10").NumberFormat = "#,##0"
    ws.Range("B9:N9").Interior.Color = RGB(248, 250, 252)
End Sub

Private Sub ConstruirResumoContratos(ByVal ws As Worksheet)
    ws.Range("B14:J14").Merge
    ws.Range("B14").Value = "RESUMO CONSOLIDADO POR CONTRATO"
    EstilizarTituloSecao ws.Range("B14:J14")

    EscreverCabecalhoResumo ws
    ws.Range("B15:J15").Font.Bold = True
    ws.Range("B15:J15").Interior.Color = RGB(226, 232, 240)
    ws.Range("B15:J15").Borders.LineStyle = xlContinuous
    ws.Range("B16:B19").Value = Application.Transpose(Array("PA-LT2", "IRON-LT1-RJ", "IRON-LT1-SP", "IRON-LT1-ES"))
    ws.Range("B20").Value = "TOTAL GERAL"
    ws.Range("C20:I20").Value = 0
    ws.Range("B16:J20").Borders.LineStyle = xlContinuous
    ws.Range("B20:J20").Font.Bold = True
    ws.Range("B20:J20").Interior.Color = RGB(248, 250, 252)
    ws.Range("B20:J20").Borders(xlEdgeBottom).LineStyle = xlDouble
End Sub

Private Sub ConstruirAlertas(ByVal ws As Worksheet)
    ws.Range("L14:N14").Merge
    ws.Range("L14").Value = "ALERTAS E PENDENCIAS"
    EstilizarTituloSecao ws.Range("L14:N14")
    ws.Range("L15").Value = "Sem alertas calculados"
    ws.Range("L15:N20").Merge
    ws.Range("L15").WrapText = True
    ws.Range("L15").VerticalAlignment = xlTop
    ws.Range("L15:N20").Interior.Color = RGB(255, 247, 224)
    ws.Range("L15:N20").Borders.LineStyle = xlContinuous
End Sub

Private Sub ConstruirArquivos(ByVal ws As Worksheet)
    ws.Range("B23:N23").Merge
    ws.Range("B23").Value = "ARQUIVOS DE MEMORIA DE CALCULO GERADOS"
    EstilizarTituloSecao ws.Range("B23:N23")
    EscreverCabecalhoArquivos ws
    ws.Range("B24:N24").Font.Bold = True
    ws.Range("B24:N24").Interior.Color = RGB(226, 232, 240)
    ws.Range("B25:B28").Value = Application.Transpose(Array("PA-LT2", "IRON-LT1-RJ", "IRON-LT1-SP", "IRON-LT1-ES"))
    EscreverLinkArquivoDashboard ws, 25, 7
    EscreverLinkArquivoDashboard ws, 26, 8
    EscreverLinkArquivoDashboard ws, 27, 9
    EscreverLinkArquivoDashboard ws, 28, 10
    ws.Range("D25:D28").Value = ""
    ws.Range("E25:E28").Value = "Verificar link"
    ws.Range("B24:N28").Borders.LineStyle = xlContinuous
    ws.Range("B25:N28").Font.Size = 9
    ws.Range("C25:C28").Font.Color = RGB(0, 102, 204)
End Sub

Private Sub ConstruirRodape(ByVal ws As Worksheet)
    ws.Range("B31:N32").Merge
    ws.Range("B31").Value = "Fonte principal: DADOS | Atualizacao: Power Query M | Exportacao: VBA | Aba original PAINEL preservada"
    ws.Range("B31").Font.Color = COR_CINZA_TEXTO
    ws.Range("B31").Font.Italic = True
    ws.Range("B31:N32").Interior.Color = RGB(241, 245, 249)
End Sub

Private Sub EscreverCabecalhoResumo(ByVal ws As Worksheet)
    ws.Range("B15").Value = "Contrato"
    ws.Range("C15").Value = "Registros"
    ws.Range("D15").Value = "Qtd. Solicitada"
    ws.Range("E15").Value = "Qtd. Atendida"
    ws.Range("F15").Value = "QExec"
    ws.Range("G15").Value = "No Prazo"
    ws.Range("H15").Value = "Fora Prazo"
    ws.Range("I15").Value = "Status"
    ws.Range("J15").Value = ""
End Sub

Private Sub EscreverCabecalhoArquivos(ByVal ws As Worksheet)
    ws.Range("B24").Value = "Contrato"
    ws.Range("C24").Value = "Nome do arquivo"
    ws.Range("D24").Value = "Data de geracao"
    ws.Range("E24").Value = "Status"
    ws.Range("F24").Value = "Acao"
End Sub

Private Sub AtualizarResumoContratos(ByVal wsDash As Worksheet, ByVal wsDados As Worksheet)
    Dim contratos   As Variant
    Dim i           As Long
    Dim colContrato As Long
    Dim ultimaLinha As Long
    Dim linha       As Long
    Dim contrato    As String
    Dim qtd         As Long

    If wsDados Is Nothing Then Exit Sub
    contratos = Array("PA-LT2", "IRON-LT1-RJ", "IRON-LT1-SP", "IRON-LT1-ES")
    colContrato = EncontrarCabecalho(wsDados, "Contrato")
    ultimaLinha = UltimaLinhaDados(wsDados)
    If colContrato = 0 Or ultimaLinha < 2 Then Exit Sub

    For i = LBound(contratos) To UBound(contratos)
        linha = 16 + i
        contrato = CStr(contratos(i))
        qtd = ContarContrato(wsDados, colContrato, contrato, 2, ultimaLinha)
        wsDash.Cells(linha, 3).Value = qtd
        wsDash.Cells(linha, 9).Value = IIf(qtd > 0, "Pronto", "Sem dados")
    Next i
End Sub

Private Sub AtualizarAlertas(ByVal ws As Worksheet, ByVal totalRegistros As Long, ByVal diferenca As Double, ByVal foraPrazo As Long)
    Dim texto As String
    texto = "Registros totais: " & Format(totalRegistros, "#,##0") & vbCrLf
    texto = texto & "Diferenca solicitado x atendido: " & Format(diferenca, "#,##0.00") & vbCrLf
    texto = texto & "Fora do prazo: " & Format(foraPrazo, "#,##0") & vbCrLf
    texto = texto & "Consulte a aba DADOS para auditoria detalhada."
    ws.Range("L15").Value = texto
End Sub

Private Function UltimaLinhaDados(ByVal ws As Worksheet) As Long
    UltimaLinhaDados = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If UltimaLinhaDados < 2 Then UltimaLinhaDados = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
End Function

Private Function EncontrarCabecalho(ByVal ws As Worksheet, ParamArray nomes()) As Long
    Dim nome        As Variant
    Dim ultimaColuna As Long
    Dim coluna      As Long
    Dim texto       As String

    ultimaColuna = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For Each nome In nomes
        For coluna = 1 To ultimaColuna
            texto = LCase$(Trim$(CStr(ws.Cells(1, coluna).Value)))
            If texto = LCase$(CStr(nome)) Or InStr(1, texto, LCase$(CStr(nome)), vbTextCompare) > 0 Then
                EncontrarCabecalho = coluna
                Exit Function
            End If
        Next coluna
    Next nome
End Function

Private Function SomarColuna(ByVal ws As Worksheet, ByVal coluna As Long, ByVal primeira As Long, ByVal ultima As Long) As Double
    Dim i As Long
    If coluna = 0 Then Exit Function
    For i = primeira To ultima
        If IsNumeric(ws.Cells(i, coluna).Value) Then SomarColuna = SomarColuna + CDbl(ws.Cells(i, coluna).Value)
    Next i
End Function

Private Function ContarContrato(ByVal ws As Worksheet, ByVal coluna As Long, ByVal contrato As String, ByVal primeira As Long, ByVal ultima As Long) As Long
    Dim i As Long
    For i = primeira To ultima
        If UCase$(Trim$(CStr(ws.Cells(i, coluna).Value))) = UCase$(contrato) Then ContarContrato = ContarContrato + 1
    Next i
End Function

Private Function ContarForaPrazo(ByVal ws As Worksheet, ByVal colPrazo As Long, ByVal colFechamento As Long, ByVal primeira As Long, ByVal ultima As Long) As Long
    Dim i As Long
    If colPrazo = 0 Or colFechamento = 0 Then Exit Function
    For i = primeira To ultima
        If IsDate(ws.Cells(i, colPrazo).Value) And IsDate(ws.Cells(i, colFechamento).Value) Then
            If CDate(ws.Cells(i, colFechamento).Value) > CDate(ws.Cells(i, colPrazo).Value) Then ContarForaPrazo = ContarForaPrazo + 1
        End If
    Next i
End Function

Private Function TextoPeriodo() As String
    On Error GoTo SemPeriodo
    TextoPeriodo = CStr(ThisWorkbook.Names("PeriodoMes").RefersToRange.Value) & "/" & CStr(ThisWorkbook.Names("PeriodoAno").RefersToRange.Value)
    Exit Function
SemPeriodo:
    TextoPeriodo = "Parametros nao encontrados"
End Function

Private Sub CriarBotao(ByVal ws As Worksheet, ByVal endereco As String, ByVal texto As String, ByVal macro As String, ByVal cor As Long, ByVal largura As Single)
    Dim shp As Shape
    Dim alvo As Range
    Set alvo = ws.Range(endereco)
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, alvo.Left, alvo.Top, largura, 25)
    shp.TextFrame2.TextRange.Characters.Text = texto
    shp.TextFrame2.TextRange.Font.Size = 10
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbWhite
    shp.Fill.ForeColor.RGB = cor
    shp.Line.ForeColor.RGB = cor
    shp.OnAction = macro
End Sub

Private Sub EstilizarTituloSecao(ByVal rng As Range)
    rng.Interior.Color = RGB(226, 232, 240)
    rng.Font.Bold = True
    rng.Font.Color = RGB(30, 41, 59)
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Color = RGB(203, 213, 225)
End Sub

Private Sub EscreverLinkArquivoDashboard(ByVal ws As Worksheet, ByVal linha As Long, ByVal linhaPainel As Long)
    ws.Cells(linha, 3).Formula = "=IFERROR(HYPERLINK(Painel!K" & linhaPainel & ",Painel!K" & linhaPainel & "),""Nao gerado"")"
End Sub

Public Sub AbrirPastaMedicao()
    Dim caminho As String
    caminho = ThisWorkbook.Path & Application.PathSeparator & "MEDIÇÃO"
    If Dir(caminho, vbDirectory) = "" Then MkDir caminho
    Shell "explorer.exe """ & caminho & """", vbNormalFocus
End Sub

Public Sub MostrarResumoDashboard()
    MsgBox "Dashboard atualizado." & vbCrLf & _
           "Consulte os cards, o resumo por contrato e a area de alertas.", _
           vbInformation, "Dashboard"
End Sub

Attribute VB_Name = "B_Calc_ARM"
Sub importar_e_calcular_ARM()






 MsgBox "Selecione o arquivo com os dados dos armazenamentos fornecido pela SOS DOCS.", vbExclamation


 Dim wsInicial As Worksheet
    Dim wsTabelas As Worksheet
    Dim wsMC As Worksheet
    Dim valorConcatenado As String
    Dim valorPeriodo As String
    Dim linhaEncontrada As Long


    Sheets("Tabelas").Visible = True
    
    ' Definir as guias
    Set wsInicial = ThisWorkbook.Sheets("Inicial")
    Set wsTabelas = ThisWorkbook.Sheets("Tabelas")
    Set wsMC = ThisWorkbook.Sheets("MC")

    ' Passo 1: Concatenar valores de F6 e G6 com " / "
    valorConcatenado = wsInicial.Range("F6").Value & " / " & wsInicial.Range("G6").Value

    ' Passo 2: Buscar o valor concatenado na coluna AP da guia "Tabelas"
    On Error Resume Next
    linhaEncontrada = Application.Match(valorConcatenado, wsTabelas.Range("AP:AP"), 0)
    On Error GoTo 0

    ' Passo 3: Se o valor for encontrado, obter o valor da coluna AQ na mesma linha
    If linhaEncontrada > 0 Then
        valorPeriodo = wsTabelas.Cells(linhaEncontrada, "AQ").Value
        ' Passo 4: Inserir o valor na célula J7 da guia "MC"
        wsMC.Range("J7").Value = valorPeriodo
    Else
        MsgBox "Período não encontrado na guia Tabelas.", vbExclamation
    End If

    Sheets("ARM").Select
    Range("A2:m10000").Select
    Selection.ClearContents
   ' MsgBox "Os dados foram apagados com sucesso!", vbInformation
    
    Call importar_dados_arm
    
End Sub


Sub importar_dados_arm()
    ' Declarar variáveis
    Dim wbOrigem As Workbook
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaLinhaOrigem As Long
    Dim caminhoArquivo As String

    ' Referenciar a guia "ARM" no workbook atual
    Set wsDestino = ThisWorkbook.Sheets("ARM")

    ' Solicitar ao usuário para selecionar o arquivo Excel a ser aberto
    On Error Resume Next
    caminhoArquivo = Application.GetOpenFilename("Arquivos do Excel (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm")
    On Error GoTo 0

    If caminhoArquivo = "False" Then Exit Sub ' Caso o usuário cancele a seleção

    ' Desabilitar a atualização automática de vínculos externos
    Application.AskToUpdateLinks = False

    ' Abrir o arquivo selecionado
    Set wbOrigem = Workbooks.Open(caminhoArquivo)

    ' Referenciar a guia "Consolidado"
    On Error Resume Next
    Set wsOrigem = wbOrigem.Sheets("Consolidado")
    On Error GoTo 0

    If wsOrigem Is Nothing Then
        MsgBox "A guia 'Consolidado' não foi encontrada no arquivo selecionado.", vbExclamation
        wbOrigem.Close SaveChanges:=False
        Exit Sub
    End If

    ' Copiar os dados para a guia "ARM"
    Application.ScreenUpdating = False

    ' Copiar A4:A (dados) para B2
    ultimaLinhaOrigem = wsOrigem.Cells(wsOrigem.Rows.Count, "A").End(xlUp).Row
    wsOrigem.Range("A4:A" & ultimaLinhaOrigem).Copy
    wsDestino.Range("B2").PasteSpecial xlPasteValues
    
    ' Copiar B4:B (dados) para C2
    ultimaLinhaOrigem = wsOrigem.Cells(wsOrigem.Rows.Count, "B").End(xlUp).Row
    wsOrigem.Range("B4:B" & ultimaLinhaOrigem).Copy
    wsDestino.Range("C2").PasteSpecial xlPasteValues

    ' Copiar J4:J (dados) para D2
    ultimaLinhaOrigem = wsOrigem.Cells(wsOrigem.Rows.Count, "J").End(xlUp).Row
    wsOrigem.Range("J4:J" & ultimaLinhaOrigem).Copy
    wsDestino.Range("D2").PasteSpecial xlPasteValues

    ' Copiar K4:K (dados) para E2
    ultimaLinhaOrigem = wsOrigem.Cells(wsOrigem.Rows.Count, "K").End(xlUp).Row
    wsOrigem.Range("K4:K" & ultimaLinhaOrigem).Copy
    wsDestino.Range("E2").PasteSpecial xlPasteValues

    ' Copiar L4:L (dados) para F2
    ultimaLinhaOrigem = wsOrigem.Cells(wsOrigem.Rows.Count, "L").End(xlUp).Row
    wsOrigem.Range("L4:L" & ultimaLinhaOrigem).Copy
    wsDestino.Range("F2").PasteSpecial xlPasteValues

    ' Copiar M4:M (dados) para G2
    ultimaLinhaOrigem = wsOrigem.Cells(wsOrigem.Rows.Count, "M").End(xlUp).Row
    wsOrigem.Range("M4:M" & ultimaLinhaOrigem).Copy
    wsDestino.Range("G2").PasteSpecial xlPasteValues

    ' Copiar I4:I (dados) para I2
    ultimaLinhaOrigem = wsOrigem.Cells(wsOrigem.Rows.Count, "I").End(xlUp).Row
    wsOrigem.Range("I4:I" & ultimaLinhaOrigem).Copy
    wsDestino.Range("I2").PasteSpecial xlPasteValues

    ' Finalizar
    Application.CutCopyMode = False
    Application.ScreenUpdating = True

    ' Fechar o arquivo de origem sem salvar alterações
    wbOrigem.Close SaveChanges:=False

    ' Informar ao usuário que o processo foi concluído
   ' MsgBox "Os dados foram importados com sucesso para a guia 'ARM'.", vbInformation

    ' Chamar a próxima função
    Call definir_item_e_Linha_de_Serviço
End Sub

Sub definir_item_e_Linha_de_Serviço()
    Dim wsARM As Worksheet
    Dim wsTabelas As Worksheet
    Dim lastRowARM As Long
    Dim i As Long
    Dim matchRow As Long
    Dim valorConcat As String
    Dim valorColunaC As String
    Dim intervaloPesquisa As Range
    Dim celulaAchada As Range

    ' Definir as guias
    Set wsARM = ThisWorkbook.Sheets("ARM")
    Set wsTabelas = ThisWorkbook.Sheets("Tabelas")

    ' Encontrar a última linha com dados na guia ARM
    lastRowARM = wsARM.Cells(wsARM.Rows.Count, "B").End(xlUp).Row

    ' Passo 1: Concatenar coluna B com " & " e coluna C, colocar resultado na coluna L
    For i = 2 To lastRowARM ' Começar na linha 2 para ignorar cabeçalho
        wsARM.Cells(i, "L").Value = wsARM.Cells(i, "B").Value & " & " & wsARM.Cells(i, "C").Value
    Next i

    ' Passo 2: Procurar valores na coluna L na guia "Tabelas" (coluna AC) e pegar os valores das colunas AA
    For i = 2 To lastRowARM
        valorConcat = wsARM.Cells(i, "L").Value
        
        ' Procurar na coluna AC da guia "Tabelas"
        matchRow = 0
        On Error Resume Next
        matchRow = Application.Match(valorConcat, wsTabelas.Columns("AC"), 0)
        On Error GoTo 0
        
        If matchRow > 0 Then
            ' Colar valor da coluna AA na coluna M da guia ARM
            wsARM.Cells(i, "M").Value = wsTabelas.Cells(matchRow, "AA").Value
        Else
            ' Caso não encontre correspondência, deixar em branco
            wsARM.Cells(i, "M").Value = ""
        End If
    Next i

    ' Passo 3: Substituir valores da coluna C pelos valores da coluna M
    For i = 2 To lastRowARM
        wsARM.Cells(i, "C").Value = wsARM.Cells(i, "M").Value
    Next i

    ' Passo 4: Nova funcionalidade - Pesquisar valor da coluna C no intervalo H2:K24
    Set intervaloPesquisa = wsTabelas.Range("I2:K24")
    
    For i = 2 To lastRowARM
        valorColunaC = wsARM.Cells(i, "C").Value
        
        ' Procurar o valor na coluna H do intervalo H2:H24
        Set celulaAchada = Nothing
        On Error Resume Next
        Set celulaAchada = intervaloPesquisa.Columns(1).Find(What:=valorColunaC, LookIn:=xlValues, LookAt:=xlWhole)
        On Error GoTo 0
        
        If Not celulaAchada Is Nothing Then
            ' Se encontrado, pegar o valor da coluna K na mesma linha
            wsARM.Cells(i, "K").Value = celulaAchada.Offset(0, 2).Value ' Coluna K é 2 colunas à direita de I
        Else
            ' Se não encontrado, deixar em branco
            wsARM.Cells(i, "K").Value = ""
        End If
    Next i

    ' Passo 5: Calcular a coluna J (H * I)
    For i = 2 To lastRowARM
        ' Verificar se as colunas H e I têm valores numéricos
        If IsNumeric(wsARM.Cells(i, "H").Value) And IsNumeric(wsARM.Cells(i, "I").Value) Then
            wsARM.Cells(i, "J").Value = wsARM.Cells(i, "H").Value * wsARM.Cells(i, "I").Value
        Else
            wsARM.Cells(i, "J").Value = ""
        End If
    Next i

    ' Passo 6: Apagar as colunas L e M
    wsARM.Columns("L:M").ClearContents

   ' MsgBox "Macro concluída com sucesso!", vbInformation
    
    Call Calcular_QUAm_e_QUA
End Sub

Sub Calcular_QUAm_e_QUA()
    Dim wsARM As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim valorD As Double, valorE As Double, valorF As Double, valorG As Double
    Dim valorH As Double, valorI As Double

    ' Definir a guia ARM
    Set wsARM = ThisWorkbook.Sheets("ARM")

    ' Encontrar a última linha preenchida na coluna D
    lastRow = wsARM.Cells(wsARM.Rows.Count, "D").End(xlUp).Row

    ' Iterar pelas linhas a partir da linha 2
    For i = 2 To lastRow
        ' Calcular coluna H: D - (E + F) + G
        valorD = wsARM.Cells(i, "D").Value
        valorE = wsARM.Cells(i, "E").Value
        valorF = wsARM.Cells(i, "F").Value
        valorG = wsARM.Cells(i, "G").Value
        
        ' Evitar erros em células vazias
        If IsNumeric(valorD) And IsNumeric(valorE) And IsNumeric(valorF) And IsNumeric(valorG) Then
            wsARM.Cells(i, "H").Value = valorD - (valorE + valorF) + valorG
        Else
            wsARM.Cells(i, "H").Value = ""
        End If

        ' Calcular coluna J: H * I
        valorH = wsARM.Cells(i, "H").Value
        valorI = wsARM.Cells(i, "I").Value
        
        If IsNumeric(valorH) And IsNumeric(valorI) Then
            wsARM.Cells(i, "J").Value = valorH * valorI
        Else
            wsARM.Cells(i, "J").Value = ""
        End If
    Next i

   ' MsgBox "Cálculos concluídos com sucesso!", vbInformation
    Call Preencher_contrato_Reg
    
End Sub

Sub Preencher_contrato_Reg()
    Dim wsMC As Worksheet
    Dim wsTabelas As Worksheet
    Dim wsARM As Worksheet
    Dim valorJ5 As String
    Dim ultimaLinhaARM As Long
    Dim linhaEncontrada As Long
    Dim valorAL As Variant
    Dim i As Long

    ' Definir as planilhas
    Set wsMC = ThisWorkbook.Sheets("MC")
    Set wsTabelas = ThisWorkbook.Sheets("Tabelas")
    Set wsARM = ThisWorkbook.Sheets("ARM")

    ' Obter o valor da célula J5 na guia MC
    valorJ5 = wsMC.Cells(5, "J").Value

    ' Encontrar a última linha preenchida na coluna B da guia ARM
    ultimaLinhaARM = wsARM.Cells(wsARM.Rows.Count, "B").End(xlUp).Row

    ' Loop para percorrer as linhas da guia ARM e preencher a coluna A com o valor encontrado na guia Tabelas
    For i = 2 To ultimaLinhaARM ' Começar na linha 2 da guia ARM
        ' Procurar o valor de valorJ5 na coluna AJ da guia Tabelas
        linhaEncontrada = 0
        On Error Resume Next
        linhaEncontrada = Application.Match(valorJ5, wsTabelas.Columns("AJ"), 0)
        On Error GoTo 0
        
        If linhaEncontrada > 0 Then
            ' Obter o valor da coluna AL da mesma linha encontrada
            valorAL = wsTabelas.Cells(linhaEncontrada, "AL").Value
            ' Preencher a coluna A da guia ARM com o valor encontrado
            wsARM.Cells(i, "A").Value = valorAL
        End If
    Next i
   
    Range("A1").Select
    ' Exibir mensagem de sucesso
    MsgBox "Valores preenchidos na guia ARM.", vbInformation
    
    Call preencher_valores_MC
    
End Sub


Sub preencher_valores_MC()
    ' Declarar variáveis
    Dim wsMC As Worksheet
    Dim lote As String
    
    ' Definir a guia MC
    Set wsMC = ThisWorkbook.Sheets("MC")
    
    ' Obter o valor de J5 para identificar o lote
    lote = wsMC.Range("J5").Value
    
    ' Se for Lote 1 (RJ, SP, DF, ES)
    If lote = "Lote 1 - RJ" Or lote = "Lote 1 - SP" Or lote = "Lote 1 - DF" Or lote = "Lote 1 - ES" Then
        ' Preencher as células de L13 a L52 para Lote 1
        wsMC.Range("L13").Value = 0.93
        wsMC.Range("L16").Value = 3.11
        wsMC.Range("L19").Value = 0.48
        wsMC.Range("L22").Value = 49.18
        wsMC.Range("L25").Value = 0.13
        wsMC.Range("L28").Value = 0.09
        wsMC.Range("L30").Value = 25
        wsMC.Range("L31").Value = 150
        wsMC.Range("L32").Value = 0.25
        wsMC.Range("L33").Value = 0.3
        wsMC.Range("L34").Value = 1.58
        wsMC.Range("L35").Value = 2.8
        wsMC.Range("L36").Value = 532747.3
        wsMC.Range("L37").Value = 1.18
        wsMC.Range("L39").Value = 0.44
        wsMC.Range("L40").Value = 1.12
        wsMC.Range("L41").Value = 2.01
        wsMC.Range("L42").Value = 20
        wsMC.Range("L43").Value = 20
        wsMC.Range("L44").Value = 5.56
        wsMC.Range("L45").Value = 0.08
        wsMC.Range("L46").Value = 0.08
        wsMC.Range("L47").Value = 0.12
        wsMC.Range("L48").Value = 0.26
        wsMC.Range("L49").Value = 99.96
        wsMC.Range("L50").Value = 30.13
        wsMC.Range("L51").Value = 9.94
        wsMC.Range("L52").Value = 2.02
        
    ' Se for Lote 2
    ElseIf lote = "Lote 2" Then
        ' Preencher as células de L13 a L52 para Lote 2
        wsMC.Range("L13").Value = 0.93
        wsMC.Range("L16").Value = 4.53
        wsMC.Range("L19").Value = 0.74
        wsMC.Range("L22").Value = 64.74
        wsMC.Range("L25").Value = 0.23
        wsMC.Range("L28").Value = 0.15
        wsMC.Range("L30").Value = 25
        wsMC.Range("L31").Value = 150
        wsMC.Range("L32").Value = 0.25
        wsMC.Range("L33").Value = 0.3
        wsMC.Range("L34").Value = 2.15
        wsMC.Range("L35").Value = 2.8
        wsMC.Range("L36").Value = 591107.26
        wsMC.Range("L37").Value = 1.18
        wsMC.Range("L39").Value = 0.44
        wsMC.Range("L40").Value = 1.12
        wsMC.Range("L41").Value = 2.01
        wsMC.Range("L42").Value = 20
        wsMC.Range("L43").Value = 20
        wsMC.Range("L44").Value = 5.56
        wsMC.Range("L45").Value = 0.14
        wsMC.Range("L46").Value = 0.23
        wsMC.Range("L47").Value = 0.14
        wsMC.Range("L48").Value = 0.7
        wsMC.Range("L49").Value = 99.96
        wsMC.Range("L50").Value = 30.13
        wsMC.Range("L51").Value = 9.94
        wsMC.Range("L52").Value = 0
    End If
    
    ' Aplicar o formato de número nas células de L13 a L52
    wsMC.Range("L13:L52").NumberFormat = "0.00"
    
    'formatar valores para exibição de forma contábil
    Sheets("MC").Select
    Range("L13:M52").Select
    Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Sheets("Tabelas").Select
    ActiveWindow.SelectedSheets.Visible = False
      
      Sheets("MC").Select
    Range("C10").Select
        
        ActiveWorkbook.Save
    ' Informar ao usuário que o processo foi concluído
   ' MsgBox "Valores preenchidos com sucesso para " & lote & ".", vbInformation
    
    
    
    
    Call ExportarDadosParaNovoArquivoFinal
    
End Sub

Attribute VB_Name = "ImportarTabelaB"
'===================================================================================================================
' Módulo: ImportarTabelaB
' Finalidade: Importação e sincronização dos dados de referência técnica do arquivo externo TabelaB-V2.xlsx
' Aplicação: Monitoramento Individual (A4UU, DPBR, GPZ1, GQ6S, S2IJ)
' Codificação: Windows-1252 (CP1252 / ANSI) - Padrão VBE Excel com acentuação em português
'===================================================================================================================
Option Explicit

Private Const MSG_TAB1 As String = "Monitoramento - Importação tabelaB"

Public Sub btImportacaoTabelaB()
' Copia dados de cada guia/planilha do arquivo "Tabelas" para o monitoramento > planilha tabelaB
' e registra o caminho atualizado do arquivo na célula C27 da aba Menu

    Dim wsTabelaB As Worksheet
    Dim wsMenu As Worksheet

    Dim sDirTabelaB As String
    Dim wbTabelaB As Workbook

    Dim iWSCont As Long
    Dim iColA As Long
    Dim iPlan As Long
    Dim wsPlan As Worksheet
    Dim sNomePlan As String
    Dim c As Long
    Dim iULin As Long
    Dim numCols As Long

    Dim vHdr1 As Variant
    Dim vHdr2 As Variant
    Dim vMerged() As Variant

    Dim fso As Object
    Dim sDirAuto As String
    Dim sDirPlan As String
    Dim appCalc As XlCalculation

    On Error GoTo TratarErro

    '=== 0. IDENTIFICAÇÃO DAS PLANILHAS (Por Nome da Aba ou CodeName) ===
    On Error Resume Next
    Set wsTabelaB = ThisWorkbook.Sheets("TabelaB")
    If wsTabelaB Is Nothing Then Set wsTabelaB = Planilha5
    
    Set wsMenu = ThisWorkbook.Sheets("Menu")
    If wsMenu Is Nothing Then Set wsMenu = Planilha3
    If wsMenu Is Nothing Then Set wsMenu = ThisWorkbook.Sheets("TabelaA")
    On Error GoTo TratarErro

    If wsTabelaB Is Nothing Then
        MsgBox "Aba 'TabelaB' não encontrada na pasta de trabalho.", vbCritical, MSG_TAB1
        Exit Sub
    End If

    If wsMenu Is Nothing Then
        MsgBox "Aba 'Menu' não encontrada na pasta de trabalho.", vbCritical, MSG_TAB1
        Exit Sub
    End If

    '=== 1. RESOLUÇÃO AUTOMÁTICA E REGISTRO DO CAMINHO NA CÉLULA C27 DA ABA MENU ===
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Tenta resolver o caminho relativo padrão: ..\Tabelas Especificas\TabelaB-V2.xlsx
    sDirAuto = fso.GetAbsolutePathName(ThisWorkbook.Path & "\..\Tabelas Especificas\TabelaB-V2.xlsx")
    sDirPlan = Trim(wsMenu.Range("C27").Value)
    
    If Dir(sDirAuto) <> "" Then
        sDirTabelaB = sDirAuto
        ' Salva o caminho atualizado na célula C27 da aba Menu
        wsMenu.Range("C27").Value = sDirAuto
    ElseIf sDirPlan <> "" And Dir(sDirPlan) <> "" Then
        sDirTabelaB = sDirPlan
        wsMenu.Range("C27").Value = sDirPlan
    Else
        ' Fallback interativo: caso o arquivo tenha sido movido ou renomeado
        MsgBox "O arquivo TabelaB-V2.xlsx não foi localizado automaticamente em:" & vbCrLf & _
               sDirAuto & vbCrLf & vbCrLf & _
               "Por favor, selecione o arquivo TabelaB na janela a seguir.", vbInformation, MSG_TAB1
               
        With Application.FileDialog(3) ' 3 = msoFileDialogFilePicker
            .Title = "Selecione o arquivo TabelaB-V2.xlsx"
            .Filters.Clear
            .Filters.Add "Pastas de Trabalho do Excel", "*.xlsx; *.xlsm; *.xlsb"
            .InitialFileName = ThisWorkbook.Path & "\..\"
            If .Show = -1 Then
                sDirTabelaB = .SelectedItems(1)
                ' Salva o caminho selecionado na célula C27 da aba Menu
                wsMenu.Range("C27").Value = sDirTabelaB
            Else
                MsgBox "Procedimento cancelado pelo usuário.", vbExclamation, MSG_TAB1
                Set fso = Nothing
                Exit Sub
            End If
        End With
    End If
    Set fso = Nothing

    '=== 2. PADRÃO DE ALTA PERFORMANCE (Desabilitar eventos e atualização) ===
    appCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    '=== 3. LIMPEZA DOS DADOS ANTERIORES ===
    wsTabelaB.Range("A1:BZ10000").ClearContents
    wsTabelaB.Columns("A:BZ").ColumnWidth = 9

    '=== 4. ABERTURA E LEITURA DA TABELAB ===
    Set wbTabelaB = Workbooks.Open(Filename:=sDirTabelaB, ReadOnly:=True)

    iWSCont = wbTabelaB.Worksheets.Count
    iColA = 1

    For iPlan = 1 To iWSCont
        Set wsPlan = wbTabelaB.Worksheets(iPlan)
        sNomePlan = wsPlan.Name
        c = 1
        
        Do While wsPlan.Cells(1, c).Value <> ""
            iULin = wsPlan.Cells(wsPlan.Rows.Count, c).End(xlUp).Row
            If iULin < 1 Then iULin = 1
            
            ' Grava o nome da guia no cabeçalho superior
            wsTabelaB.Cells(1, iColA).Value = sNomePlan
            
            ' Transferência direta de valores em bloco (sem uso de área de transferência / clipboard)
            wsTabelaB.Range(wsTabelaB.Cells(2, iColA), wsTabelaB.Cells(iULin + 1, iColA)).Value = _
                wsPlan.Range(wsPlan.Cells(1, c), wsPlan.Cells(iULin, c)).Value
            
            iColA = iColA + 1
            c = c + 1
        Loop
    Next iPlan

    wbTabelaB.Close SaveChanges:=False
    Set wbTabelaB = Nothing

    '=== 5. CONCATENAÇÃO DE CABEÇALHOS EM MEMÓRIA (NomePlan#Campo) ===
    numCols = iColA - 1
    If numCols >= 1 Then
        vHdr1 = wsTabelaB.Range(wsTabelaB.Cells(1, 1), wsTabelaB.Cells(1, numCols)).Value
        vHdr2 = wsTabelaB.Range(wsTabelaB.Cells(2, 1), wsTabelaB.Cells(2, numCols)).Value
        
        ReDim vMerged(1 To 1, 1 To numCols)
        For c = 1 To numCols
            vMerged(1, c) = CStr(vHdr1(1, c)) & "#" & CStr(vHdr2(1, c))
        Next c
        
        wsTabelaB.Range(wsTabelaB.Cells(1, 1), wsTabelaB.Cells(1, numCols)).Value = vMerged
        wsTabelaB.Rows(2).Delete
        wsTabelaB.Columns("A:BZ").AutoFit
    End If

    ' Grava o caminho atualizado na célula C27 e o carimbo de data/hora na aba Menu
    wsMenu.Range("C27").Value = sDirTabelaB
    wsMenu.Range("G16").Value = Now

    '=== 6. RESTAURAÇÃO DE AMBIENTE E FEEDBACK ===
    Application.Calculation = appCalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Importação da TabelaB realizada com sucesso!", vbInformation, MSG_TAB1
    Exit Sub

TratarErro:
    If Not wbTabelaB Is Nothing Then
        On Error Resume Next
        wbTabelaB.Close SaveChanges:=False
        On Error GoTo 0
    End If
    Application.Calculation = appCalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Erro durante a importação da TabelaB:" & vbCrLf & Err.Description, vbCritical, MSG_TAB1
End Sub

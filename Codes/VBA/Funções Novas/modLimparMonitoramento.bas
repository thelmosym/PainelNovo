Attribute VB_Name = "modLimparMonitoramento"
'===================================================================================================================
' MÓDULO: modLimparMonitoramento
' OBJETIVO: Limpeza completa de dados e formatações da aba "Monitoramento", mantendo exatamente 100 linhas
'           na tabela estruturada (ListObject), removendo linhas em excesso e preservando 100% das regras
'           de validação de dados (Data Validation / listas suspensas) em todas as células.
' APLICAÇÃO: Planilha de Monitoramento Individual (A4UU, DPBR, GPZ1, GQ6S, S2IJ, etc.)
'
' CORREÇÕES E BLINDAGEM CONTRA ERRO 1004:
'   1. LIMPEZA PRÉVIA DE FILTROS: Desativa filtros ativos na planilha e na tabela antes de qualquer ação.
'   2. REDIMENSIONAMENTO NATIVO (Resize): Substitui .Delete xlShiftUp por tbl.Resize, eliminando o erro 1004
'      de tentativa de deslocar células dentro de uma tabela estruturada.
'   3. LIMPEZA SEGURA DE COLUNAS RESIDUAIS: Utiliza .Clear em vez de .Delete xlToLeft ao lado da tabela.
'   4. POSICIONAMENTO SEGURO: Utiliza Application.Goto após reabilitar ScreenUpdating.
'   5. RASTREAMENTO DE ETAPAS: Identifica a etapa exata em caso de exceção.
'
' CODIFICAÇÃO: Windows-1252 (CP1252 / ANSI) com quebras de linha CRLF para compatibilidade VBE.
'===================================================================================================================
Option Explicit

Private Const MSG_TITULO As String = "Limpeza da Aba Monitoramento"
Private Const QTD_LINHAS_PADRAO As Long = 100

' Estrutura para armazenamento seguro das propriedades de validação
Private Type InfoValidacao
    TemValidacao As Boolean
    Tipo As Long
    EstiloAlerta As Long
    Operador As Long
    Formula1 As String
    Formula2 As String
    IgnorarBranco As Boolean
    ListaSuspensa As Boolean
    ExibirEntrada As Boolean
    ExibirErro As Boolean
    TituloEntrada As String
    MsgEntrada As String
    TituloErro As String
    MsgErro As String
End Type

'-------------------------------------------------------------------------------------------------------------------
' PROCEDIMENTO PRINCIPAL (PONTO DE ENTRADA DO USUÁRIO)
' Vinculável a botões na aba "Menu" ou "Monitoramento", ou executável via Alt+F8.
'-------------------------------------------------------------------------------------------------------------------
Public Sub LimparMonitoramento()
    LimparAbaMonitoramento100Linhas False
End Sub

'-------------------------------------------------------------------------------------------------------------------
' PROCEDIMENTO PARAMETRIZADO
' Permite execução silenciosa (bSilencioso = True) para rotinas automatizadas e scripts batch.
'-------------------------------------------------------------------------------------------------------------------
Public Sub LimparAbaMonitoramento100Linhas(Optional ByVal bSilencioso As Boolean = False)
    Dim sEtapaAtual As String
    sEtapaAtual = "Inicialização"

    On Error GoTo TratarErro

    Dim wsMonit As Worksheet
    Dim tbl As ListObject
    Dim appCalc As XlCalculation
    Dim resposta As VbMsgBoxResult
    Dim lLinhaFimTabela As Long
    Dim lUltimaLinhaPlan As Long
    Dim lColFimTabela As Long
    Dim lUltimaColPlan As Long
    Dim rngNovoTamanho As Range
    Dim vValidacoes() As InfoValidacao
    Dim sNomeTabela As String

    ' 1. Confirmação do Usuário (se não estiver em modo silencioso)
    If Not bSilencioso Then
        resposta = MsgBox("ATENÇÃO: Este procedimento irá:" & vbNewLine & vbNewLine & _
                          "1. Limpar todos os dados digitados na aba 'Monitoramento'." & vbNewLine & _
                          "2. Remover formatações manuais e regras condicionais fragmentadas." & vbNewLine & _
                          "3. Ajustar a tabela para conter exatamente 100 linhas (A3:AK102), removendo o excesso." & vbNewLine & _
                          "4. Manter e revalidar todas as listas suspensas (validações de dados)." & vbNewLine & vbNewLine & _
                          "Deseja realmente continuar?", vbQuestion + vbYesNo + vbDefaultButton2, MSG_TITULO)
        If resposta <> vbYes Then Exit Sub
    End If

    ' 2. Salvar e Desativar Recursos de Interface para Máxima Performance
    sEtapaAtual = "Configuração do ambiente"
    appCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' 3. Localizar Planilha 'Monitoramento' (ActiveWorkbook com fallback para ThisWorkbook)
    sEtapaAtual = "Localização da aba Monitoramento"
    Set wsMonit = ObterAbaMonitoramento()
    If wsMonit Is Nothing Then
        MsgBox "[ERRO] A aba 'Monitoramento' não foi encontrada na pasta de trabalho ativa.", vbCritical, MSG_TITULO
        GoTo SairRotina
    End If

    ' 4. Localizar Tabela Estruturada (ListObject)
    sEtapaAtual = "Localização da tabela estruturada"
    Set tbl = ObterTabelaMonitoramento(wsMonit)
    If tbl Is Nothing Then
        MsgBox "[ERRO] Nenhuma tabela estruturada encontrada na aba 'Monitoramento'.", vbCritical, MSG_TITULO
        GoTo SairRotina
    End If

    sNomeTabela = tbl.Name

    ' 5. Desativar quaisquer filtros ativos (crucial para evitar Erro 1004)
    sEtapaAtual = "Limpeza de filtros ativos"
    On Error Resume Next
    If wsMonit.FilterMode Then wsMonit.ShowAllData
    If Not tbl.AutoFilter Is Nothing Then
        If tbl.AutoFilter.FilterMode Then tbl.AutoFilter.ShowAllData
    End If
    On Error GoTo TratarErro

    ' 6. Mapear e Salvar Validações de Dados das Colunas Existentes
    sEtapaAtual = "Mapeamento das validações de dados"
    ReDim vValidacoes(1 To tbl.ListColumns.Count)
    CapturarValidacoesTabela tbl, vValidacoes

    ' 7. Redimensionamento Seguro da Tabela para 100 Linhas de Dados (1 Cabeçalho + 100 Dados = 101 Linhas)
    sEtapaAtual = "Redimensionamento da tabela para 100 linhas"
    Set rngNovoTamanho = wsMonit.Range(tbl.Range.Cells(1, 1), _
                                      tbl.Range.Cells(QTD_LINHAS_PADRAO + 1, tbl.Range.Columns.Count))
    tbl.Resize rngNovoTamanho

    ' 8. Limpeza dos Dados e Formatações na Tabela (100 Linhas)
    sEtapaAtual = "Limpeza de dados e formatações da tabela"
    If Not tbl.DataBodyRange Is Nothing Then
        With tbl.DataBodyRange
            .ClearContents            ' Limpa dados e textos
            .FormatConditions.Delete  ' Elimina regras de formatação condicional duplicadas/fragmentadas
            .ClearFormats             ' Limpa formatações manuais e cores sobrepostas
            .RowHeight = 20           ' Altura confortável de linha
            .VerticalAlignment = xlCenter
        End With
    End If

    ' 9. Excluir Linhas Residuais Abaixo da Tabela na Planilha (linha 103 em diante)
    sEtapaAtual = "Exclusão de linhas residuais abaixo da tabela"
    lLinhaFimTabela = tbl.Range.Row + tbl.Range.Rows.Count - 1
    lUltimaLinhaPlan = wsMonit.Cells.SpecialCells(xlCellTypeLastCell).Row
    If lUltimaLinhaPlan > lLinhaFimTabela Then
        wsMonit.Rows((lLinhaFimTabela + 1) & ":" & lUltimaLinhaPlan).Delete xlShiftUp
    End If

    ' 10. Limpar Células Residuais à Direita da Tabela (sem usar Delete para evitar Erro 1004)
    sEtapaAtual = "Limpeza de células à direita da tabela"
    lColFimTabela = tbl.Range.Column + tbl.Range.Columns.Count - 1
    lUltimaColPlan = wsMonit.Cells.SpecialCells(xlCellTypeLastCell).Column
    If lUltimaColPlan > lColFimTabela Then
        wsMonit.Range(wsMonit.Cells(1, lColFimTabela + 1), _
                      wsMonit.Cells(wsMonit.Rows.Count, lUltimaColPlan)).Clear
    End If

    ' 11. Restaurar e Garantir Validação de Dados nas 100 Linhas
    sEtapaAtual = "Reaplicação das validações de dados"
    ReaplicarValidacoesTabela tbl, vValidacoes, wsMonit

    ' 12. Formatação Padronizada de Tipos de Colunas Específicas
    sEtapaAtual = "Padronização de formatos e alinhamentos"
    FormatarColunasPadrao tbl

    ' 13. Garantir Ativação dos Estilos Nativos da Tabela e Filtros
    On Error Resume Next
    tbl.ShowTableStyleRowStripes = True
    tbl.ShowAutoFilter = True
    On Error GoTo TratarErro

    ' 14. Restaurar Ambiente e Posicionar Cursor de Forma Segura na Célula A3
    sEtapaAtual = "Finalização e posicionamento"
    Application.ScreenUpdating = True
    On Error Resume Next
    Application.Goto wsMonit.Range("A3"), True
    On Error GoTo TratarErro

    ' 15. Mensagem de Sucesso (se não estiver em modo silencioso)
    If Not bSilencioso Then
        MsgBox "[OK] Limpeza da aba 'Monitoramento' concluída com sucesso!" & vbNewLine & vbNewLine & _
               "- Tabela '" & sNomeTabela & "' redimensionada para exatamente 100 linhas (A3:AK102)." & vbNewLine & _
               "- Todos os dados anteriores e formatações manuais foram limpos." & vbNewLine & _
               "- Linhas e resquícios excedentes eliminados da planilha." & vbNewLine & _
               "- Validações de dados (listas suspensas) 100% preservadas e ativas.", _
               vbInformation, MSG_TITULO
    End If

SairRotina:
    ' Restauração estrita dos estados da aplicação
    On Error Resume Next
    Application.Calculation = appCalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

TratarErro:
    MsgBox "[ERRO] Ocorreu um erro durante a limpeza da aba 'Monitoramento':" & vbNewLine & vbNewLine & _
           "Etapa: " & sEtapaAtual & vbNewLine & _
           "Código: " & Err.Number & vbNewLine & _
           "Descrição: " & Err.Description, vbCritical, MSG_TITULO
    Resume SairRotina
End Sub

'-------------------------------------------------------------------------------------------------------------------
' FUNÇÃO AUXILIAR: Localiza a aba 'Monitoramento' com segurança
'-------------------------------------------------------------------------------------------------------------------
Private Function ObterAbaMonitoramento() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets("Monitoramento")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets("Monitoramento")
    End If
    On Error GoTo 0
    Set ObterAbaMonitoramento = ws
End Function

'-------------------------------------------------------------------------------------------------------------------
' FUNÇÃO AUXILIAR: Localiza o ListObject da aba 'Monitoramento'
'-------------------------------------------------------------------------------------------------------------------
Private Function ObterTabelaMonitoramento(ByVal ws As Worksheet) As ListObject
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ws.ListObjects("Monitoramento")
    If tbl Is Nothing And ws.ListObjects.Count > 0 Then
        Set tbl = ws.ListObjects(1)
    End If
    On Error GoTo 0
    Set ObterTabelaMonitoramento = tbl
End Function

'-------------------------------------------------------------------------------------------------------------------
' SUBROTINA AUXILIAR: Mapeia as regras de validação presentes em cada coluna
'-------------------------------------------------------------------------------------------------------------------
Private Sub CapturarValidacoesTabela(ByVal tbl As ListObject, ByRef vValidacoes() As InfoValidacao)
    Dim c As Long
    Dim r As Long
    Dim rngCol As Range
    Dim cellTest As Range
    Dim vObj As Validation
    Dim lMaxLinhasTeste As Long

    For c = 1 To tbl.ListColumns.Count
        vValidacoes(c).TemValidacao = False
        Set rngCol = tbl.ListColumns(c).DataBodyRange

        If Not rngCol Is Nothing Then
            lMaxLinhasTeste = Application.WorksheetFunction.Min(rngCol.Rows.Count, 5)
            For r = 1 To lMaxLinhasTeste
                Set cellTest = rngCol.Cells(r, 1)
                On Error Resume Next
                Set vObj = cellTest.Validation
                If Not vObj Is Nothing Then
                    If vObj.Type <> 0 Then
                        vValidacoes(c).TemValidacao = True
                        vValidacoes(c).Tipo = vObj.Type
                        vValidacoes(c).EstiloAlerta = vObj.AlertStyle
                        vValidacoes(c).Operador = vObj.Operator
                        vValidacoes(c).Formula1 = vObj.Formula1
                        vValidacoes(c).Formula2 = ""
                        vValidacoes(c).Formula2 = vObj.Formula2
                        vValidacoes(c).IgnorarBranco = vObj.IgnoreBlank
                        vValidacoes(c).ListaSuspensa = vObj.InCellDropdown
                        vValidacoes(c).ExibirEntrada = vObj.ShowInput
                        vValidacoes(c).ExibirErro = vObj.ShowError
                        vValidacoes(c).TituloEntrada = vObj.InputTitle
                        vValidacoes(c).MsgEntrada = vObj.InputMessage
                        vValidacoes(c).TituloErro = vObj.ErrorTitle
                        vValidacoes(c).MsgErro = vObj.ErrorMessage
                        On Error GoTo 0
                        Exit For
                    End If
                End If
                On Error GoTo 0
            Next r
        End If
    Next c
End Sub

'-------------------------------------------------------------------------------------------------------------------
' SUBROTINA AUXILIAR: Reaplica as validações de dados capturadas e injeta regras de fallback
'-------------------------------------------------------------------------------------------------------------------
Private Sub ReaplicarValidacoesTabela(ByVal tbl As ListObject, ByRef vValidacoes() As InfoValidacao, ByVal wsMonit As Worksheet)
    Dim c As Long
    Dim rngCol As Range
    Dim sFormulaFbk As String
    Dim bTemTabelaA As Boolean

    ' Verifica se a aba de apoio TabelaA existe na pasta de trabalho
    On Error Resume Next
    bTemTabelaA = Not (wsMonit.Parent.Sheets("TabelaA") Is Nothing)
    On Error GoTo 0

    For c = 1 To tbl.ListColumns.Count
        Set rngCol = tbl.ListColumns(c).DataBodyRange
        If Not rngCol Is Nothing Then
            If vValidacoes(c).TemValidacao Then
                ' Reaplica a validação capturada previamente
                AplicarValidacaoNoRange rngCol, vValidacoes(c)
            ElseIf bTemTabelaA Then
                ' Fallback inteligente: se a coluna perdeu a validação, recupera do padrão contratual da TabelaA
                sFormulaFbk = ObterFormulaFallbackPorIndice(c)
                If Len(sFormulaFbk) > 0 Then
                    Dim infoFbk As InfoValidacao
                    infoFbk.TemValidacao = True
                    infoFbk.Tipo = 3 ' xlValidateList
                    infoFbk.EstiloAlerta = 1 ' xlValidAlertStop
                    infoFbk.Operador = 1 ' xlBetween
                    infoFbk.Formula1 = sFormulaFbk
                    infoFbk.IgnorarBranco = True
                    infoFbk.ListaSuspensa = True
                    infoFbk.ExibirEntrada = True
                    infoFbk.ExibirErro = True
                    AplicarValidacaoNoRange rngCol, infoFbk
                End If
            End If
        End If
    Next c
End Sub

'-------------------------------------------------------------------------------------------------------------------
' SUBROTINA AUXILIAR: Aplica uma estrutura de validação em um determinado Range
'-------------------------------------------------------------------------------------------------------------------
Private Sub AplicarValidacaoNoRange(ByVal rngAlvo As Range, ByRef info As InfoValidacao)
    On Error Resume Next
    rngAlvo.Validation.Delete
    If Len(info.Formula2) > 0 Then
        rngAlvo.Validation.Add Type:=info.Tipo, AlertStyle:=info.EstiloAlerta, _
                               Operator:=info.Operador, Formula1:=info.Formula1, Formula2:=info.Formula2
    Else
        rngAlvo.Validation.Add Type:=info.Tipo, AlertStyle:=info.EstiloAlerta, _
                               Operator:=info.Operador, Formula1:=info.Formula1
    End If
    rngAlvo.Validation.IgnoreBlank = info.IgnorarBranco
    rngAlvo.Validation.InCellDropdown = info.ListaSuspensa
    rngAlvo.Validation.ShowInput = info.ExibirEntrada
    rngAlvo.Validation.ShowError = info.ExibirErro
    If Len(info.TituloEntrada) > 0 Then rngAlvo.Validation.InputTitle = info.TituloEntrada
    If Len(info.MsgEntrada) > 0 Then rngAlvo.Validation.InputMessage = info.MsgEntrada
    If Len(info.TituloErro) > 0 Then rngAlvo.Validation.ErrorTitle = info.TituloErro
    If Len(info.MsgErro) > 0 Then rngAlvo.Validation.ErrorMessage = info.MsgErro
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------------------------------------------
' FUNÇÃO AUXILIAR: Fórmulas de fallback para as colunas padrão de Monitoramento Individual
'-------------------------------------------------------------------------------------------------------------------
Private Function ObterFormulaFallbackPorIndice(ByVal c As Long) As String
    Select Case c
        Case 3  ' Coluna C: Situação
            ObterFormulaFallbackPorIndice = "=TabelaA!$B$2:$B$5"
        Case 4  ' Coluna D: Contrato
            ObterFormulaFallbackPorIndice = "=TabelaA!$E$2:$E$6"
        Case 5  ' Coluna E: UF
            ObterFormulaFallbackPorIndice = "=TabelaA!$AD$2:$AD$57"
        Case 6  ' Coluna F: Descrição da atividade
            ObterFormulaFallbackPorIndice = "=TabelaA!$AO$2:$AO$35"
        Case 7  ' Coluna G: Item
            ObterFormulaFallbackPorIndice = "=TabelaA!$AQ$2:$AQ$53"
        Case 8  ' Coluna H: Aplicação
            ObterFormulaFallbackPorIndice = "=TabelaA!$AH$2:$AH$3"
        Case 25 ' Coluna Y: OS disponibilizada? / Faturamento Direto
            ObterFormulaFallbackPorIndice = "=TabelaA!$AW$2:$AW$3"
        Case Else
            ObterFormulaFallbackPorIndice = ""
    End Select
End Function

'-------------------------------------------------------------------------------------------------------------------
' SUBROTINA AUXILIAR: Formata colunas de data/hora e alinhamentos básicos
'-------------------------------------------------------------------------------------------------------------------
Private Sub FormatarColunasPadrao(ByVal tbl As ListObject)
    On Error Resume Next
    ' Coluna B: Período faturamento
    If tbl.ListColumns.Count >= 2 Then
        tbl.ListColumns(2).DataBodyRange.NumberFormat = "dd/mm/aaaa (ddd)"
        tbl.ListColumns(2).DataBodyRange.HorizontalAlignment = xlCenter
    End If

    ' Colunas de Códigos e Siglas Curtas (Centralizadas)
    If tbl.ListColumns.Count >= 3 Then tbl.ListColumns(3).DataBodyRange.HorizontalAlignment = xlCenter ' Situação
    If tbl.ListColumns.Count >= 4 Then tbl.ListColumns(4).DataBodyRange.HorizontalAlignment = xlCenter ' Contrato
    If tbl.ListColumns.Count >= 5 Then tbl.ListColumns(5).DataBodyRange.HorizontalAlignment = xlCenter ' UF
    If tbl.ListColumns.Count >= 8 Then tbl.ListColumns(8).DataBodyRange.HorizontalAlignment = xlCenter ' Aplicação
    If tbl.ListColumns.Count >= 25 Then tbl.ListColumns(25).DataBodyRange.HorizontalAlignment = xlCenter ' OS disponibilizada

    ' Colunas de Data/Hora (dd/mm/aaaa hh:mm)
    If tbl.ListColumns.Count >= 12 Then tbl.ListColumns(12).DataBodyRange.NumberFormat = "dd/mm/aaaa hh:mm" ' D. solicitação
    If tbl.ListColumns.Count >= 17 Then tbl.ListColumns(17).DataBodyRange.NumberFormat = "dd/mm/aaaa hh:mm" ' D. abertura
    If tbl.ListColumns.Count >= 18 Then tbl.ListColumns(18).DataBodyRange.NumberFormat = "dd/mm/aaaa hh:mm" ' D. fechamento

    ' Colunas de Quantidades (Alinhadas ao centro)
    If tbl.ListColumns.Count >= 15 Then tbl.ListColumns(15).DataBodyRange.HorizontalAlignment = xlCenter ' Qtd. Solicitada
    If tbl.ListColumns.Count >= 20 Then tbl.ListColumns(20).DataBodyRange.HorizontalAlignment = xlCenter ' Qtd. Atendida
    On Error GoTo 0
End Sub

Attribute VB_Name = "modAtualizarConsultas"
Option Explicit

'====================================================================================================
' MÓDULO: modAtualizarConsultas
' OBJETIVO: Força a atualização SÍNCRONA de todas as conexões e consultas Power Query do Workbook,
'           exibindo o progresso em tempo real na célula F9 da aba 'Painel' com cores de status:
'             [AMARELO] AMARELO  - Atualização em andamento
'             [VERDE] VERDE    - Concluído com sucesso (registra data/hora e tempo decorrido)
'             [VERMELHO] VERMELHO - Concluído com interrupção fatal
'
' PADRÕES DE ENGENHARIA APLICADOS:
'   - Tipagem 64-bit Long para contadores de conexões.
'   - Preservação e restauração estrita do estado da aplicação (Calculation, EnableEvents, ScreenUpdating).
'   - Desativação explícita de BackgroundQuery para prevenir concorrência em pipelines dependentes.
'   - Tratamento resiliente de conexões órfãs ou inválidas sem abortar o processamento das demais.
'   - Execução de CalculateUntilAsyncQueriesDone para consolidação de cálculos pendentes.
'====================================================================================================

Private Const MSG_TITULO_ATUALIZACAO As String = "Atualização de Consultas Power Query"

' Aba e célula onde o status visual será exibido
Private Const NOME_PLANILHA_STATUS As String = "Painel"
Private Const ENDERECO_CELULA_STATUS As String = "F9"

' Paleta de cores corporativa para status visual (RGB)
Private Const COR_AMARELO As Long = 65535        ' RGB(255, 255, 0)
Private Const COR_VERDE   As Long = 5296274      ' RGB(146, 208, 80) - Verde suave
Private Const COR_VERMELHO As Long = 255         ' RGB(255, 0, 0)

'----------------------------------------------------------------------------------------------------
' PONTO DE ENTRADA PÚBLICO (Associado ao botão "Atualizar Consultas" no Painel)
'----------------------------------------------------------------------------------------------------
Public Sub AtualizarTodasConsultasPowerQuery()
    Dim wb          As Workbook
    Dim cn          As WorkbookConnection
    Dim wsStatus    As Worksheet
    Dim rngStatus   As Range
    Dim dtInicio    As Date
    Dim dtFim       As Date
    Dim iTotal      As Long
    Dim iAtual      As Long
    Dim iComErro    As Long
    Dim sErros      As String
    Dim sResumo     As String
    
    Dim appCalc     As XlCalculation
    Dim bEvents     As Boolean
    Dim bScreen     As Boolean
    
    Set wb = ThisWorkbook
    iTotal = wb.Connections.Count
    
    If iTotal = 0 Then
        MsgBox "Nenhuma conexão ou consulta Power Query foi encontrada neste arquivo.", _
               vbInformation, MSG_TITULO_ATUALIZACAO
        Exit Sub
    End If
    
    Set wsStatus = ObterPlanilhaStatus(wb, NOME_PLANILHA_STATUS)
    Set rngStatus = wsStatus.Range(ENDERECO_CELULA_STATUS)
    
    If MsgBox("Serão atualizadas " & iTotal & " conexão(ões)/consulta(s) Power Query de forma síncrona." & vbNewLine & vbNewLine & _
              "O progresso será refletido na célula " & wsStatus.Name & "!" & ENDERECO_CELULA_STATUS & "." & vbNewLine & vbNewLine & _
              "Deseja iniciar a atualização?", vbYesNo + vbQuestion, MSG_TITULO_ATUALIZACAO) = vbNo Then
        MsgBox "Atualização cancelada.", vbInformation, MSG_TITULO_ATUALIZACAO
        Exit Sub
    End If
    
    '-- Captura o estado prévio da aplicação
    appCalc = Application.Calculation
    bEvents = Application.EnableEvents
    bScreen = Application.ScreenUpdating
    
    On Error GoTo TratarErro
    
    ' Mantemos ScreenUpdating ativo para que o redesenho da célula F9 seja visível
    Application.EnableEvents = False
    
    dtInicio = Now
    iAtual = 0
    iComErro = 0
    sErros = ""
    
    '-- Sinaliza início do processamento ([AMARELO] Amarelo)
    With rngStatus
        .Interior.Color = COR_AMARELO
        .Value = "Iniciando atualização de " & iTotal & " consulta(s)..."
    End With
    DoEvents
    
    '-- Percorre e atualiza cada conexão individualmente
    For Each cn In wb.Connections
        iAtual = iAtual + 1
        
        rngStatus.Value = "[" & iAtual & "/" & iTotal & "] Atualizando: " & cn.Name & "..."
        Application.StatusBar = rngStatus.Value
        DoEvents
        
        On Error Resume Next
        Err.Clear
        
        ' Garante execução síncrona para conexões OLEDB / Power Query
        If cn.Type = xlConnectionTypeOLEDB Then
            cn.OLEDBConnection.BackgroundQuery = False
        End If
        
        cn.Refresh
        
        If Err.Number <> 0 Then
            iComErro = iComErro + 1
            sErros = sErros & "  • " & cn.Name & ": " & Err.Description & vbNewLine
            Err.Clear
        End If
        On Error GoTo TratarErro
    Next cn
    
    rngStatus.Value = "Finalizando (aguardando conclusão dos cálculos assíncronos)..."
    DoEvents
    
    On Error Resume Next
    Application.CalculateUntilAsyncQueriesDone
    On Error GoTo TratarErro
    
    dtFim = Now
    
    '-- Sinaliza conclusão ([VERDE] Verde ou [AMARELO] Amarelo com advertência)
    With rngStatus
        If iComErro = 0 Then
            .Interior.Color = COR_VERDE
            .Value = "[OK] Concluído em " & Format(dtFim, "dd/mm/yyyy hh:mm:ss") & _
                     " — " & iTotal & " consulta(s) 100% atualizada(s)."
        Else
            .Interior.Color = COR_AMARELO
            .Value = "[AVISO] Concluído em " & Format(dtFim, "dd/mm/yyyy hh:mm:ss") & _
                     " — " & (iTotal - iComErro) & " OK, " & iComErro & " com pendência(s)."
        End If
    End With

SairRotina:
    Application.StatusBar = False
    Application.EnableEvents = bEvents
    Application.ScreenUpdating = bScreen
    Application.Calculation = appCalc
    
    ' Monta resumo informativo
    sResumo = "Atualização de consultas finalizada." & vbNewLine & vbNewLine & _
              "Total de conexões: " & iTotal & vbNewLine & _
              "Atualizadas com sucesso: " & (iTotal - iComErro) & vbNewLine & _
              "Com pendências/erros: " & iComErro & vbNewLine & vbNewLine & _
              "Início: " & Format(dtInicio, "dd/mm/yyyy hh:mm:ss") & vbNewLine & _
              "Término: " & Format(dtFim, "dd/mm/yyyy hh:mm:ss") & vbNewLine & _
              "Tempo total: " & Format(dtFim - dtInicio, "hh:mm:ss")
              
    If iComErro > 0 Then
        sResumo = sResumo & vbNewLine & vbNewLine & _
                  "Detalhes das pendências (possíveis conexões órfãs ou fontes offline):" & vbNewLine & sErros
        MsgBox sResumo, vbExclamation, MSG_TITULO_ATUALIZACAO
    Else
        MsgBox sResumo, vbInformation, MSG_TITULO_ATUALIZACAO
    End If
    Exit Sub

TratarErro:
    On Error Resume Next
    With rngStatus
        .Interior.Color = COR_VERMELHO
        .Value = "[ERRO] Erro fatal em " & Format(Now, "dd/mm/yyyy hh:mm:ss") & ": " & Err.Description
    End With
    On Error GoTo 0
    
    MsgBox "Erro inesperado durante a atualização das consultas:" & vbNewLine & _
           Err.Number & " - " & Err.Description, vbCritical, MSG_TITULO_ATUALIZACAO
    Resume SairRotina
End Sub

Private Function ObterPlanilhaStatus(ByVal wb As Workbook, ByVal nomeAba As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nomeAba)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = nomeAba
    End If
    
    Set ObterPlanilhaStatus = ws
End Function

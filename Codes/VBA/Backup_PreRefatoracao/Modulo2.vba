Option Explicit

'-----------------------------------------------------------------------
' MÓDULO: modAtualizarConsultas
'---------------------------------------------------------------------
' Força a atualização SÍNCRONA de todas as conexões/consultas Power
' Query do workbook (uma por vez), exibindo o progresso em uma CÉLULA
' específica da planilha — com a cor de fundo mudando conforme o
' status:
'   ?? AMARELO  ? atualização em andamento
'   ?? VERDE    ? concluído com sucesso (mostra data/hora de conclusão)
'   ?? VERMELHO ? concluído com erro fatal (interrompeu antes do fim)
' Ao final, exibe também uma mensagem (MsgBox) com o resumo completo.
'
' Associe esta macro (AtualizarTodasConsultasPowerQuery) a um botão
' via: Inserir > Formas/Botão > Botão direito > Atribuir Macro.
'-----------------------------------------------------------------------

Private Const MSG_TITULO_ATUALIZACAO As String = "Atualização de Consultas Power Query"

' ?? AJUSTE AQUI — aba e célula onde o status será exibido
Private Const NOME_PLANILHA_STATUS As String = "Painel"   ' nome da aba
Private Const ENDERECO_CELULA_STATUS As String = "B12"   ' endereço da célula

' ?? Cores de status (RGB) — ajuste os tons se preferir outra tonalidade
Private Const COR_AMARELO As Long = 65535        ' RGB(255,255,0)
Private Const COR_VERDE   As Long = 5296274      ' RGB(146, 208, 80) — verde suave
Private Const COR_VERMELHO As Long = 255         ' RGB(255,0,0)
Private Const COR_PADRAO_SEM_PREENCHIMENTO As Long = -4142 ' xlColorIndexNone equivalente


'-----------------------------------------------------------------------
' PROCEDIMENTO PRINCIPAL — associar este ao botão
'-----------------------------------------------------------------------
Public Sub AtualizarTodasConsultasPowerQuery()

    Dim wb          As Workbook
    Dim cn          As WorkbookConnection
    Dim wsStatus    As Worksheet
    Dim rngStatus   As Range
    Dim dtInicio    As Date
    Dim dtFim       As Date
    Dim iTotal      As Integer
    Dim iAtual      As Integer
    Dim iComErro    As Integer
    Dim sErros      As String
    Dim sResumo     As String

    On Error GoTo TratarErro

    Set wb = ThisWorkbook
    iTotal = wb.Connections.Count

    If iTotal = 0 Then
        MsgBox "Nenhuma conexão/consulta Power Query foi encontrada neste arquivo.", _
               vbInformation, MSG_TITULO_ATUALIZACAO
        Exit Sub
    End If

    '-- Localiza (ou cria, se não existir) a célula de status ------------
    Set wsStatus = ObterPlanilhaStatus(wb, NOME_PLANILHA_STATUS)
    Set rngStatus = wsStatus.Range(ENDERECO_CELULA_STATUS)

    '-- Confirmação antes de iniciar --------------------------------------
    If MsgBox("Serão atualizadas " & iTotal & " conexão(ões)/consulta(s)." & _
              vbNewLine & vbNewLine & "O progresso será exibido na célula " & _
              wsStatus.Name & "!" & ENDERECO_CELULA_STATUS & "." & vbNewLine & vbNewLine & _
              "INICIAR?", vbYesNo + vbQuestion, MSG_TITULO_ATUALIZACAO) = vbNo Then
        MsgBox "Atualização cancelada.", vbInformation, MSG_TITULO_ATUALIZACAO
        Exit Sub
    End If

    ' Mantemos ScreenUpdating = True para que a cor/texto da célula
    ' sejam realmente redesenhados na tela durante o processo.
    Application.EnableEvents = False

    dtInicio = Now
    iAtual = 0
    iComErro = 0

    '-- ?? AMARELO — marca início do processamento -----------------------
    With rngStatus
        .Interior.Color = COR_AMARELO
        .Value = "? Iniciando atualização de " & iTotal & " consulta(s)..."
    End With
    DoEvents

    '-- Percorre e atualiza cada conexão, uma por vez (modo síncrono) ----
    For Each cn In wb.Connections
        iAtual = iAtual + 1

        rngStatus.Value = "? Atualizando " & iAtual & " de " & iTotal & _
                           ": " & cn.Name & "..."
        Application.StatusBar = rngStatus.Value
        DoEvents ' força o redesenho da célula/tela imediatamente

        On Error Resume Next
        Err.Clear

        cn.OLEDBConnection.BackgroundQuery = False
        Err.Clear

        cn.Refresh

        If Err.Number <> 0 Then
            iComErro = iComErro + 1
            sErros = sErros & "  - " & cn.Name & ": " & Err.Description & vbNewLine
            Err.Clear
        End If

        On Error GoTo TratarErro
    Next cn

    rngStatus.Value = "? Finalizando (aguardando cálculos pendentes)..."
    DoEvents

    On Error Resume Next
    Application.CalculateUntilAsyncQueriesDone
    On Error GoTo TratarErro

    dtFim = Now

    '-- ?? VERDE — concluído (mesmo que com erros parciais registrados) --
    With rngStatus
        .Interior.Color = COR_VERDE
        If iComErro = 0 Then
            .Value = "? Concluído em " & Format(dtFim, "dd/mm/yyyy hh:mm:ss") & _
                      " — " & iTotal & " consulta(s) atualizada(s) com sucesso."
        Else
            .Value = "? Concluído em " & Format(dtFim, "dd/mm/yyyy hh:mm:ss") & _
                      " — " & (iTotal - iComErro) & " OK, " & iComErro & " com erro."
        End If
    End With

    Application.StatusBar = False
    Application.EnableEvents = True

    '-- Monta a mensagem final (MsgBox) -----------------------------------
    sResumo = "Atualização de consultas concluída." & vbNewLine & vbNewLine & _
              "Total de conexões processadas: " & iTotal & vbNewLine & _
              "Concluídas com sucesso: " & (iTotal - iComErro) & vbNewLine & _
              "Com erro: " & iComErro & vbNewLine & vbNewLine & _
              "Início: " & Format(dtInicio, "dd/mm/yyyy hh:mm:ss") & vbNewLine & _
              "Término: " & Format(dtFim, "dd/mm/yyyy hh:mm:ss") & vbNewLine & _
              "Duração: " & Format(dtFim - dtInicio, "hh:mm:ss")

    If sErros <> "" Then
        sResumo = sResumo & vbNewLine & vbNewLine & _
                  "? Detalhes dos erros (dica: 'Query does not exist' indica " & _
                  "uma conexão órfã, sem consulta Power Query correspondente):" & _
                  vbNewLine & sErros
        MsgBox sResumo, vbExclamation, MSG_TITULO_ATUALIZACAO
    Else
        MsgBox sResumo, vbInformation, MSG_TITULO_ATUALIZACAO
    End If

    Exit Sub

TratarErro:
    Application.StatusBar = False
    Application.EnableEvents = True

    '-- ?? VERMELHO — erro fatal, processo interrompido antes do fim -----
    On Error Resume Next
    With rngStatus
        .Interior.Color = COR_VERMELHO
        .Value = "? Erro em " & Format(Now, "dd/mm/yyyy hh:mm:ss") & _
                  ": " & Err.Description
    End With
    On Error GoTo 0

    MsgBox "Erro inesperado durante a atualização das consultas:" & vbNewLine & _
           Err.Number & " - " & Err.Description, vbCritical, MSG_TITULO_ATUALIZACAO
End Sub


'-----------------------------------------------------------------------
' Retorna a Worksheet informada; se não existir, CRIA uma nova aba com
' esse nome (para garantir que sempre haja um local válido para o
' status, mesmo que a aba "Menu" ainda não exista no arquivo).
'-----------------------------------------------------------------------
Private Function ObterPlanilhaStatus(wb As Workbook, nomeAba As String) As Worksheet
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



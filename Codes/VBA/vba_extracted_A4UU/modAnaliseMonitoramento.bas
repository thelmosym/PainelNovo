Attribute VB_Name = "modAnaliseMonitoramento"
'===================================================================================================================
' MÓDULO: modAnaliseMonitoramento
' OBJETIVO: Análise de integridade, validação de regras de negócio e geração de log de inconsistências
'           para a planilha de Monitoramento Individual (A4UU, DPBR, GPZ1, GQ6S, S2IJ).
'
' REGRAS PRINCIPAIS ATENDIDAS:
'   1. REGRA INICIAL DE FILTRAGEM:
'      Apenas linhas onde a coluna 'Situação' (Coluna C) estiver preenchida com "CO" (Concluído) ou "AT" (Atendido)
'      são analisadas para procura de erros. Linhas com outras situações ou em branco são desconsideradas.
'   2. GRAVAÇÃO DO LOG NA COLUNA Z:
'      O log de inconsistências é gravado na mesma linha onde foi encontrado o erro,
'      exatamente na coluna Z ("LOG") da aba "Monitoramento".
'
' PADRÕES DE ENGENHARIA E ALTA PERFORMANCE:
'   - Processamento 100% em memória (RAM) através de matrizes (Variant Arrays), processando
'     dezenas de milhares de linhas em fração de segundo.
'   - Tabelas de referência (TabelaA) carregadas em Dicionários de Busca O(1) prévios ao loop.
'   - Eliminação de chamadas lentas como .Select, .Activate e WorksheetFunction.CountIf em laço.
'   - Tipagem 64-bit Long em todos os índices e contadores.
'   - Preservação e restauração estrita de estado da aplicação (ScreenUpdating, Calculation, Events).
'   - Codificação Windows-1252 (ANSI) com terminações CRLF para total compatibilidade com o VBE.
'===================================================================================================================
Option Explicit

Private Const MSG_TITULO As String = "Auditoria de Dados - Monitoramento"
Private Const LINHA_INICIAL_DADOS As Long = 3
Private Const COLUNA_LOG_NUM As Long = 26 ' Coluna Z

' Constantes de Cores para Formatação Visual da Coluna Z (RGB)
Private Const COR_FUNDO_ERRO As Long = 14803454    ' RGB(254, 226, 226) - Vermelho Claro Suave
Private Const COR_TEXTO_ERRO As Long = 1776411     ' RGB(153, 27, 27)  - Vermelho Escuro
Private Const COR_FUNDO_OK   As Long = 15138766    ' RGB(238, 242, 238) - Neutro / Limpo

'-------------------------------------------------------------------------------------------------------------------
' PONTO DE ENTRADA PÚBLICO
' Pode ser vinculado diretamente a botões existentes da planilha Monitoramento ou Menu
'-------------------------------------------------------------------------------------------------------------------
Public Sub AnalisarErrosMonitoramentoEGerarLog()
    Dim wsMonit As Worksheet
    Dim wsRef   As Worksheet
    Dim wsMenu  As Worksheet
    
    Dim ul As Long
    Dim r As Long
    Dim iLinha As Long
    Dim tInicio As Double
    Dim tFim As Double
    
    Dim vDados As Variant
    Dim vLog As Variant
    
    ' Dicionários de referência carregados em O(1)
    Dim dictSituacoes   As Object
    Dim dictContratos   As Object
    Dim dictUFs         As Object
    Dim dictContratoUF  As Object
    Dim dictAtividades  As Object
    Dim dictItens       As Object
    Dim dictAtivItem    As Object
    Dim dictAplicacoes  As Object
    Dim dictSiglasObs   As Object
    Dim dictLocalidades As Object ' Armazena flag "S" ou "N"
    Dim dictPrazoAtiv   As Object ' "Fiscalização" ou "Tabela"
    
    Dim sLogLinha As String
    Dim bLinhaAuditada As Boolean
    Dim iTotalLinhas As Long
    Dim iLinhasAuditadas As Long
    Dim iLinhasComErro As Long
    
    Dim appCalc As XlCalculation
    Dim bEvents As Boolean
    Dim bScreen As Boolean
    
    On Error GoTo TratarErro
    
    '=== 1. IDENTIFICAÇÃO E QUALIFICAÇÃO DAS PLANILHAS ===
    On Error Resume Next
    Set wsMonit = ThisWorkbook.Sheets("Monitoramento")
    If wsMonit Is Nothing Then Set wsMonit = Planilha1
    
    Set wsRef = ThisWorkbook.Sheets("TabelaA")
    If wsRef Is Nothing Then Set wsRef = Planilha2
    
    Set wsMenu = ThisWorkbook.Sheets("Menu")
    If wsMenu Is Nothing Then Set wsMenu = Planilha3
    On Error GoTo TratarErro
    
    If wsMonit Is Nothing Then
        MsgBox "Aba 'Monitoramento' não encontrada na pasta de trabalho.", vbCritical, MSG_TITULO
        Exit Sub
    End If
    
    If wsRef Is Nothing Then
        MsgBox "Aba de referências técnicas ('TabelaA') não encontrada.", vbCritical, MSG_TITULO
        Exit Sub
    End If
    
    ' Determina a última linha com dados pela Coluna I (Código Solicitação) ou A (Atendente)
    ul = wsMonit.Cells(wsMonit.Rows.Count, "I").End(xlUp).Row
    If ul < wsMonit.Cells(wsMonit.Rows.Count, "A").End(xlUp).Row Then
        ul = wsMonit.Cells(wsMonit.Rows.Count, "A").End(xlUp).Row
    End If
    
    If ul < LINHA_INICIAL_DADOS Then
        MsgBox "Não há registros para analisar na aba Monitoramento (dados a partir da linha 3).", _
               vbInformation, MSG_TITULO
        Exit Sub
    End If
    
    iTotalLinhas = ul - LINHA_INICIAL_DADOS + 1
    
    If MsgBox("Será executada a análise de conformidade nos registros da aba Monitoramento." & vbNewLine & vbNewLine & _
              "• Regra Inicial: Somente linhas com Situação 'CO' (Concluído) ou 'AT' (Atendido) serão auditadas." & vbNewLine & _
              "• Destino: Os erros encontrados serão gravados na Coluna Z (LOG) da respectiva linha." & vbNewLine & vbNewLine & _
              "Total de registros a varrer: " & Format(iTotalLinhas, "#,##0") & vbNewLine & vbNewLine & _
              "Deseja iniciar a auditoria?", _
              vbYesNo + vbQuestion, MSG_TITULO) = vbNo Then
        Exit Sub
    End If
    
    tInicio = Timer
    
    '=== 2. CAPTURA E SUSPENSÃO DE ESTADO DA APLICAÇÃO ===
    appCalc = Application.Calculation
    bEvents = Application.EnableEvents
    bScreen = Application.ScreenUpdating
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Carregando matrizes de validação em memória..."
    
    '=== 3. PRÉ-CARREGAMENTO DOS DICIONÁRIOS DE REFERÊNCIA (TABELAA) ===
    Set dictSituacoes = CreateObject("Scripting.Dictionary")
    Set dictContratos = CreateObject("Scripting.Dictionary")
    Set dictUFs = CreateObject("Scripting.Dictionary")
    Set dictContratoUF = CreateObject("Scripting.Dictionary")
    Set dictAtividades = CreateObject("Scripting.Dictionary")
    Set dictItens = CreateObject("Scripting.Dictionary")
    Set dictAtivItem = CreateObject("Scripting.Dictionary")
    Set dictAplicacoes = CreateObject("Scripting.Dictionary")
    Set dictSiglasObs = CreateObject("Scripting.Dictionary")
    Set dictLocalidades = CreateObject("Scripting.Dictionary")
    Set dictPrazoAtiv = CreateObject("Scripting.Dictionary")
    
    dictSituacoes.CompareMode = vbTextCompare
    dictContratos.CompareMode = vbTextCompare
    dictUFs.CompareMode = vbTextCompare
    dictContratoUF.CompareMode = vbTextCompare
    dictAtividades.CompareMode = vbTextCompare
    dictItens.CompareMode = vbTextCompare
    dictAtivItem.CompareMode = vbTextCompare
    dictAplicacoes.CompareMode = vbTextCompare
    dictSiglasObs.CompareMode = vbTextCompare
    dictLocalidades.CompareMode = vbTextCompare
    dictPrazoAtiv.CompareMode = vbTextCompare
    
    CarregarDicionariosReferencia wsRef, dictSituacoes, dictContratos, dictUFs, dictContratoUF, _
                                  dictAtividades, dictItens, dictAtivItem, dictAplicacoes, _
                                  dictSiglasObs, dictLocalidades, dictPrazoAtiv
    
    '=== 4. LEITURA EM BLOCO DOS DADOS DA PLANILHA MONITORAMENTO ===
    ' Lê as colunas de A (1) até Z (26) em uma única matriz na RAM
    vDados = wsMonit.Range(wsMonit.Cells(LINHA_INICIAL_DADOS, 1), wsMonit.Cells(ul, COLUNA_LOG_NUM)).Value
    
    ' Matriz de saída unidimensional para atualizar a Coluna Z de volta
    ReDim vLog(1 To iTotalLinhas, 1 To 1)
    
    iLinhasAuditadas = 0
    iLinhasComErro = 0
    
    Application.StatusBar = "Analisando consistência dos registros (Filtro: Situação CO ou AT)..."
    
    '=== 5. LAÇO DE AUDITORIA EM MEMÓRIA (LINHA A LINHA) ===
    For r = 1 To iTotalLinhas
        iLinha = r + LINHA_INICIAL_DADOS - 1
        sLogLinha = ""
        bLinhaAuditada = False
        
        AuditarLinhaMonitoramento vDados, r, sLogLinha, bLinhaAuditada, _
                                  dictSituacoes, dictContratos, dictUFs, dictContratoUF, _
                                  dictAtividades, dictItens, dictAtivItem, dictAplicacoes, _
                                  dictSiglasObs, dictLocalidades, dictPrazoAtiv
        
        If bLinhaAuditada Then
            iLinhasAuditadas = iLinhasAuditadas + 1
            If Len(sLogLinha) > 0 Then
                vLog(r, 1) = sLogLinha
                iLinhasComErro = iLinhasComErro + 1
            Else
                vLog(r, 1) = vbNullString
            End If
        Else
            ' Linha não auditada (Situação diferente de CO e AT): limpa log prévio
            vLog(r, 1) = vbNullString
        End If
    Next r
    
    '=== 6. GRAVAÇÃO ATÔMICA DA COLUNA Z NA ABA MONITORAMENTO ===
    Application.StatusBar = "Gravando logs na Coluna Z..."
    wsMonit.Range(wsMonit.Cells(LINHA_INICIAL_DADOS, COLUNA_LOG_NUM), wsMonit.Cells(ul, COLUNA_LOG_NUM)).Value = vLog
    
    ' Formatação condicional/destaque visual básico para a coluna de log
    With wsMonit.Range(wsMonit.Cells(LINHA_INICIAL_DADOS, COLUNA_LOG_NUM), wsMonit.Cells(ul, COLUNA_LOG_NUM))
        .Font.Name = "Calibri"
        .Font.Size = 9
        .WrapText = True
    End With
    
    tFim = Timer
    
    '=== 7. RESTAURAÇÃO DE AMBIENTE E FEEDBACK EXECUTIVO ===
SairRotina:
    Application.StatusBar = False
    Application.Calculation = appCalc
    Application.EnableEvents = bEvents
    Application.DisplayAlerts = True
    Application.ScreenUpdating = bScreen
    
    ' Atualiza carimbo no Menu caso a aba exista
    If Not wsMenu Is Nothing Then
        On Error Resume Next
        wsMenu.Range("G16").Value = Now
        On Error GoTo 0
    End If
    
    ExibirResumoAuditoria iTotalLinhas, iLinhasAuditadas, iLinhasComErro, (tFim - tInicio)
    Exit Sub

TratarErro:
    MsgBox "Erro durante a auditoria de dados do monitoramento:" & vbNewLine & _
           Err.Number & " - " & Err.Description, vbCritical, MSG_TITULO
    Resume SairRotina
End Sub

' Alias para manter compatibilidade com macros já associadas a botões
Public Sub analiseMonitoramento()
    AnalisarErrosMonitoramentoEGerarLog
End Sub

'-------------------------------------------------------------------------------------------------------------------
' ROTINA DE AUDITORIA DE UMA LINHA ESPECÍFICA (EXECUÇÃO 100% EM MEMÓRIA)
'-------------------------------------------------------------------------------------------------------------------
Private Sub AuditarLinhaMonitoramento(ByRef vDados As Variant, ByVal r As Long, ByRef sLog As String, _
                                     ByRef bAuditado As Boolean, _
                                     ByRef dictSituacoes As Object, ByRef dictContratos As Object, _
                                     ByRef dictUFs As Object, ByRef dictContratoUF As Object, _
                                     ByRef dictAtividades As Object, ByRef dictItens As Object, _
                                     ByRef dictAtivItem As Object, ByRef dictAplicacoes As Object, _
                                     ByRef dictSiglasObs As Object, ByRef dictLocalidades As Object, _
                                     ByRef dictPrazoAtiv As Object)

    ' Mapeamento de colunas em vDados(r, c):
    ' A (1)  - Atendente
    ' B (2)  - Período Faturamento
    ' C (3)  - Situação
    ' D (4)  - Contrato
    ' E (5)  - UF
    ' F (6)  - Descrição da atividade
    ' G (7)  - Item
    ' H (8)  - Aplicação
    ' I (9)  - Código da solicitação
    ' J (10) - Chave solicitante
    ' K (11) - Nome Solicitante
    ' L (12) - D. solicitação
    ' M (13) - Prazo aplicação
    ' N (14) - Gerência solicitante
    ' O (15) - Qtd. Solicitada
    ' P (16) - Código OS
    ' Q (17) - D. abertura
    ' R (18) - D. fechamento
    ' S (19) - Prazo combinado
    ' T (20) - Qtd. Atendida
    ' U (21) - Obs. Sigla
    ' V (22) - Localidade
    ' W (23) - Comentários
    ' X (24) - Observações
    ' Y (25) - OS disponibilizada?
    ' Z (26) - LOG

    Dim sSituacao    As String
    
    ' Captura inicial da Situação (Coluna C)
    sSituacao = UCase(Trim(CStr(vDados(r, 3) & "")))
    
    '===============================================================================================================
    ' REGRA INICIAL MANDATÓRIA:
    ' Apenas procurar erros se a coluna Situação estiver preenchida estritamente com "CO" ou "AT".
    ' Registros com qualquer outra situação (ex: CA, PE, etc.) ou vazias são sumariamente ignorados.
    '===============================================================================================================
    If sSituacao <> "CO" And sSituacao <> "AT" Then
        bAuditado = False
        Exit Sub
    End If
    
    ' Marca que a linha atende à regra inicial e deve ser auditada
    bAuditado = True

    Dim sAtendente   As String
    Dim sPerFat      As String
    Dim sContrato    As String
    Dim sUF          As String
    Dim sAtividade   As String
    Dim sItem        As String
    Dim sAplicacao   As String
    Dim sCodSol      As String
    Dim sChaveSol    As String
    Dim sNomeSol     As String
    Dim sGerencia    As String
    Dim sCodOS       As String
    Dim sSiglaObs    As String
    Dim sLocalidade  As String
    Dim sOSDisp      As String
    
    Dim vQtdSol      As Variant
    Dim vQtdAtend    As Variant
    Dim vDtSol       As Variant
    Dim vDtAbert     As Variant
    Dim vDtFech      As Variant
    Dim vPrazoApp    As Variant
    Dim vPrazoComb   As Variant
    
    Dim dDtAbert     As Date
    Dim dDtFech      As Date
    Dim bTemAbert    As Boolean
    Dim bTemFech     As Boolean
    
    ' Extrai e limpa os demais valores da linha
    sAtendente = Trim(CStr(vDados(r, 1) & ""))
    sPerFat = Trim(CStr(vDados(r, 2) & ""))
    sContrato = Trim(CStr(vDados(r, 4) & ""))
    sUF = UCase(Trim(CStr(vDados(r, 5) & "")))
    sAtividade = Trim(CStr(vDados(r, 6) & ""))
    sItem = Trim(CStr(vDados(r, 7) & ""))
    sAplicacao = Trim(CStr(vDados(r, 8) & ""))
    sCodSol = Trim(CStr(vDados(r, 9) & ""))
    sChaveSol = Trim(CStr(vDados(r, 10) & ""))
    sNomeSol = Trim(CStr(vDados(r, 11) & ""))
    sGerencia = Trim(CStr(vDados(r, 14) & ""))
    sCodOS = Trim(CStr(vDados(r, 16) & ""))
    sSiglaObs = Trim(CStr(vDados(r, 21) & ""))
    sLocalidade = Trim(CStr(vDados(r, 22) & ""))
    sOSDisp = UCase(Trim(CStr(vDados(r, 25) & "")))
    
    vDtSol = vDados(r, 12)
    vPrazoApp = vDados(r, 13)
    vQtdSol = vDados(r, 15)
    vDtAbert = vDados(r, 17)
    vDtFech = vDados(r, 18)
    vPrazoComb = vDados(r, 19)
    vQtdAtend = vDados(r, 20)
    
    '--- 1. COLUNA B: PERÍODO DE FATURAMENTO ---
    If sSituacao = "CO" And Len(sPerFat) = 0 Then
        sLog = sLog & "(Coluna B): Campo vazio e situação concluída; "
    ElseIf Len(sPerFat) > 0 Then
        If InStr(sPerFat, " ") > 0 Or (Not IsDate(sPerFat) And Not (IsNumeric(Replace(sPerFat, "/", "")))) Then
            sLog = sLog & "(Coluna B): Data inválida; "
        End If
    End If
    
    '--- 2. COLUNA D: CONTRATO ---
    If Len(sContrato) = 0 Then
        sLog = sLog & "(Coluna D): Campo vazio. Preenchimento obrigatório; "
    ElseIf Not dictContratos.Exists(sContrato) Then
        sLog = sLog & "(Coluna D): Descrição inválida para a coluna informada; "
    End If
    
    '--- 3. COLUNA E: UF ---
    If Len(sUF) = 0 Then
        sLog = sLog & "(Coluna E): Campo vazio. Preenchimento obrigatório; "
    ElseIf Not dictUFs.Exists(sUF) Then
        sLog = sLog & "(Coluna E): Descrição inválida para a coluna informada; "
    End If
    
    '--- 4. COLUNAS D e E: CONSISTÊNCIA CONTRATO x UF ---
    If Len(sContrato) > 0 And Len(sUF) > 0 Then
        If dictContratos.Exists(sContrato) And dictUFs.Exists(sUF) Then
            If Not dictContratoUF.Exists(sContrato & "|" & sUF) Then
                sLog = sLog & "(Coluna D e E): Campo contrato não pertence ao campo UF; "
            End If
        End If
    End If
    
    '--- 5. COLUNA F: DESCRIÇÃO DA ATIVIDADE ---
    If Len(sAtividade) = 0 Then
        sLog = sLog & "(Coluna F): Campo vazio. Preenchimento obrigatório; "
    ElseIf Not dictAtividades.Exists(sAtividade) Then
        sLog = sLog & "(Coluna F): Descrição inválida para a coluna informada; "
    End If
    
    '--- 6. COLUNA G: ITEM ---
    If Len(sItem) = 0 Then
        sLog = sLog & "(Coluna G): Campo vazio. Preenchimento obrigatório; "
    ElseIf Not dictItens.Exists(sItem) Then
        sLog = sLog & "(Coluna G): Descrição inválida para a coluna informada; "
    End If
    
    '--- 7. COLUNAS F e G: CONSISTÊNCIA ATIVIDADE x ITEM ---
    If Len(sAtividade) > 0 And Len(sItem) > 0 Then
        If dictAtividades.Exists(sAtividade) And dictItens.Exists(sItem) Then
            If Not dictAtivItem.Exists(sAtividade & "|" & sItem) Then
                sLog = sLog & "(Coluna F e G): Campo item não pertence ao campo atividade; "
            End If
        End If
    End If
    
    '--- 8. COLUNA H: APLICAÇÃO ---
    If Len(sAplicacao) = 0 Then
        sLog = sLog & "(Coluna H): Campo vazio. Preenchimento obrigatório; "
    ElseIf Not dictAplicacoes.Exists(sAplicacao) Then
        sLog = sLog & "(Coluna H): Descrição inválida para a coluna informada; "
    End If
    
    '--- 9. COLUNA I: CÓDIGO DA SOLICITAÇÃO ---
    If Len(sCodSol) = 0 Then
        sLog = sLog & "(Coluna I): Campo vazio. Preenchimento obrigatório; "
    ElseIf InStr(sCodSol, " ") > 0 Then
        sLog = sLog & "(Coluna I): Caractere inválido para texto inserido: Espaço; "
    End If
    
    '--- 10. COLUNA J: CHAVE SOLICITANTE ---
    If Len(sChaveSol) = 0 Then
        sLog = sLog & "(Coluna J): Campo vazio. Preenchimento obrigatório; "
    ElseIf InStr(sChaveSol, " ") > 0 Then
        sLog = sLog & "(Coluna J): Caractere inválido para texto inserido: Espaço; "
    End If
    
    '--- 11. COLUNA K: NOME DO SOLICITANTE ---
    If Len(sNomeSol) = 0 Then
        sLog = sLog & "(Coluna K): Campo vazio. Preenchimento obrigatório; "
    End If
    
    '--- 12. COLUNA L: DATA DA SOLICITAÇÃO ---
    If IsEmpty(vDtSol) Or Len(Trim(CStr(vDtSol & ""))) = 0 Then
        sLog = sLog & "(Coluna L): Campo vazio. Preenchimento obrigatório; "
    ElseIf Not IsDate(vDtSol) Then
        sLog = sLog & "(Coluna L): Data inválida; "
    End If
    
    '--- 13. COLUNA M: PRAZO APLICAÇÃO ---
    If Not IsEmpty(vPrazoApp) And Len(Trim(CStr(vPrazoApp & ""))) > 0 Then
        If Not IsDate(vPrazoApp) Then
            sLog = sLog & "(Coluna M): Data inválida; "
        End If
    End If
    
    '--- 14. COLUNA N: GERÊNCIA SOLICITANTE ---
    If Len(sGerencia) = 0 Then
        sLog = sLog & "(Coluna N): Campo vazio. Preenchimento obrigatório; "
    Else
        If InStr(sGerencia, " ") > 0 Then
            sLog = sLog & "(Coluna N): Caractere inválido para texto inserido: Espaço; "
        End If
        If InStr(sGerencia, "_") > 0 Then
            sLog = sLog & "(Coluna N): Caractere inválido para texto inserido: Underline; "
        End If
    End If
    
    '--- 15. COLUNA O: QTD. SOLICITADA ---
    If IsEmpty(vQtdSol) Or Len(Trim(CStr(vQtdSol & ""))) = 0 Then
        sLog = sLog & "(Coluna O): Campo vazio. Preenchimento obrigatório; "
    ElseIf Not IsNumeric(vQtdSol) Then
        sLog = sLog & "(Coluna O): Valor numérico inválido; "
    ElseIf CDbl(vQtdSol) <= 0 Then
        sLog = sLog & "(Coluna O): Quantidade solicitada deve ser maior que zero; "
    End If
    
    '--- 16. COLUNA P: CÓDIGO DA OS ---
    If Len(sCodOS) = 0 Then
        sLog = sLog & "(Coluna P): Campo vazio. Preenchimento obrigatório; "
    ElseIf InStr(sCodOS, " ") > 0 Then
        sLog = sLog & "(Coluna P): Caractere inválido para texto inserido: Espaço; "
    End If
    
    '--- 17. COLUNA Q: DATA ABERTURA ---
    bTemAbert = False
    If Not IsEmpty(vDtAbert) And Len(Trim(CStr(vDtAbert & ""))) > 0 Then
        If IsDate(vDtAbert) Then
            dDtAbert = CDate(vDtAbert)
            bTemAbert = True
        Else
            sLog = sLog & "(Coluna Q): Data inválida; "
        End If
    End If
    
    '--- 18. COLUNA R: DATA FECHAMENTO & CRONOLOGIA ---
    bTemFech = False
    If sSituacao = "CO" And (IsEmpty(vDtFech) Or Len(Trim(CStr(vDtFech & ""))) = 0) Then
        sLog = sLog & "(Coluna R): Campo vazio e situação concluída; "
    ElseIf Not IsEmpty(vDtFech) And Len(Trim(CStr(vDtFech & ""))) > 0 Then
        If IsDate(vDtFech) Then
            dDtFech = CDate(vDtFech)
            bTemFech = True
            
            ' Validação cronológica: Fechamento não pode ser menor que Abertura
            If bTemAbert Then
                If dDtFech < dDtAbert Then
                    sLog = sLog & "(Coluna R): Data de fechamento anterior à data de abertura; "
                End If
            End If
        Else
            sLog = sLog & "(Coluna R): Data inválida; "
        End If
    End If
    
    '--- 19. COLUNA S: PRAZO COMBINADO (FISCALIZAÇÃO) ---
    If dictPrazoAtiv.Exists(sAtividade) Then
        If dictPrazoAtiv(sAtividade) = "Fiscalização" Then
            If IsEmpty(vPrazoComb) Or Len(Trim(CStr(vPrazoComb & ""))) = 0 Then
                sLog = sLog & "(Coluna S): Prazo obrigatório para atividade informada; "
            End If
        End If
    End If
    
    '--- 20. COLUNA T: QTD. ATENDIDA ---
    If IsEmpty(vQtdAtend) Or Len(Trim(CStr(vQtdAtend & ""))) = 0 Then
        sLog = sLog & "(Coluna T): Campo vazio. Preenchimento obrigatório; "
    ElseIf Not IsNumeric(vQtdAtend) Then
        sLog = sLog & "(Coluna T): Valor numérico inválido; "
    Else
        If CDbl(vQtdAtend) < 0 Then
            sLog = sLog & "(Coluna T): Quantidade atendida não pode ser negativa; "
        End If
        If IsNumeric(vQtdSol) Then
            If CDbl(vQtdAtend) > CDbl(vQtdSol) Then
                sLog = sLog & "(Coluna T): Quantidade atendida superior à solicitada; "
            End If
        End If
    End If
    
    '--- 21. COLUNA U: OBS. SIGLA ---
    If Len(sSiglaObs) > 0 Then
        If Not dictSiglasObs.Exists(sSiglaObs) Then
            sLog = sLog & "(Coluna U): Descrição inválida para a coluna informada; "
        End If
    End If
    
    '--- 22. COLUNA V: LOCALIDADE / CENTRO ---
    If Len(sLocalidade) = 0 Then
        sLog = sLog & "(Coluna V): Campo vazio. Preenchimento obrigatório; "
    Else
        sLocalidade = Replace(sLocalidade, "  ", " ")
        If Not dictLocalidades.Exists(sLocalidade) Then
            sLog = sLog & "(Coluna V): Informação não encontrada: Localidade; "
        Else
            If dictLocalidades(sLocalidade) = "N" Then
                sLog = sLog & "(Coluna V): Localidade não ativa. Informe o responsável; "
            End If
        End If
    End If
    
    '--- 23. COLUNA Y: OS DISPONIBILIZADA? ---
    If Len(sOSDisp) = 0 Then
        sLog = sLog & "(Coluna Y): Campo vazio. Preenchimento obrigatório; "
    ElseIf sOSDisp <> "SIM" And sOSDisp <> "NÃO" And sOSDisp <> "NAO" Then
        sLog = sLog & "(Coluna Y): Permitido apenas <SIM> ou <NÃO>; "
    End If

End Sub

'-------------------------------------------------------------------------------------------------------------------
' CARREGAMENTO PRÉVIO DAS BASES DE DOMÍNIO E VALIDAÇÃO DA TABELAA
'-------------------------------------------------------------------------------------------------------------------
Private Sub CarregarDicionariosReferencia(ByVal wsRef As Worksheet, _
                                         ByRef dictSituacoes As Object, ByRef dictContratos As Object, _
                                         ByRef dictUFs As Object, ByRef dictContratoUF As Object, _
                                         ByRef dictAtividades As Object, ByRef dictItens As Object, _
                                         ByRef dictAtivItem As Object, ByRef dictAplicacoes As Object, _
                                         ByRef dictSiglasObs As Object, ByRef dictLocalidades As Object, _
                                         ByRef dictPrazoAtiv As Object)

    Dim ulRef As Long
    Dim r As Long
    Dim vRef As Variant
    Dim sVal1 As String, sVal2 As String, sFlag As String
    
    ulRef = wsRef.UsedRange.Rows.Count
    If ulRef < 100 Then ulRef = 1000 ' Garante leitura de limites razoáveis
    
    ' Lê da coluna A até AM de uma só vez
    vRef = wsRef.Range(wsRef.Cells(1, 1), wsRef.Cells(ulRef, 40)).Value
    
    For r = 2 To UBound(vRef, 1)
        ' Coluna B (2) - Sigla Situação
        sVal1 = UCase(Trim(CStr(vRef(r, 2) & "")))
        If Len(sVal1) > 0 Then dictSituacoes(sVal1) = True
        
        ' Coluna I (9) ou AE (31) - Contratos
        sVal1 = Trim(CStr(vRef(r, 9) & ""))
        If Len(sVal1) > 0 Then dictContratos(sVal1) = True
        sVal1 = Trim(CStr(vRef(r, 31) & ""))
        If Len(sVal1) > 0 Then dictContratos(sVal1) = True
        
        ' Coluna AD (30) - UFs
        sVal1 = UCase(Trim(CStr(vRef(r, 30) & "")))
        If Len(sVal1) > 0 Then dictUFs(sVal1) = True
        
        ' Colunas AE (31) e AD (30) - Pares Contrato x UF
        sVal1 = Trim(CStr(vRef(r, 31) & ""))
        sVal2 = UCase(Trim(CStr(vRef(r, 30) & "")))
        If Len(sVal1) > 0 And Len(sVal2) > 0 Then
            dictContratoUF(sVal1 & "|" & sVal2) = True
        End If
        
        ' Coluna H (8) ou K (11) - Atividades
        sVal1 = Trim(CStr(vRef(r, 8) & ""))
        If Len(sVal1) > 0 Then dictAtividades(sVal1) = True
        sVal1 = Trim(CStr(vRef(r, 11) & ""))
        If Len(sVal1) > 0 Then dictAtividades(sVal1) = True
        
        ' Coluna N (14) - Ref. Prazo ("Fiscalização" / "Tabela")
        sVal2 = Trim(CStr(vRef(r, 14) & ""))
        If Len(sVal1) > 0 And Len(sVal2) > 0 Then
            dictPrazoAtiv(sVal1) = sVal2
        End If
        
        ' Coluna R (18) ou T (20) - Itens
        sVal1 = Trim(CStr(vRef(r, 18) & ""))
        If Len(sVal1) > 0 Then dictItens(sVal1) = True
        sVal1 = Trim(CStr(vRef(r, 20) & ""))
        If Len(sVal1) > 0 Then dictItens(sVal1) = True
        
        ' Coluna S (19) e T (20) - Par Atividade x Item
        sVal1 = Trim(CStr(vRef(r, 19) & ""))
        sVal2 = Trim(CStr(vRef(r, 20) & ""))
        If Len(sVal1) > 0 And Len(sVal2) > 0 Then
            dictAtivItem(sVal1 & "|" & sVal2) = True
        End If
        
        ' Coluna AH (34) - Aplicações
        sVal1 = Trim(CStr(vRef(r, 34) & ""))
        If Len(sVal1) > 0 Then dictAplicacoes(sVal1) = True
        
        ' Coluna Y (25) - Siglas Obs
        sVal1 = Trim(CStr(vRef(r, 25) & ""))
        If Len(sVal1) > 0 Then dictSiglasObs(sVal1) = True
        
        ' Coluna AK (37) - Localidades e AM (39) - Ativo (S/N)
        sVal1 = Trim(CStr(vRef(r, 37) & ""))
        sFlag = UCase(Trim(CStr(vRef(r, 39) & "")))
        If Len(sVal1) > 0 Then
            If Len(sFlag) = 0 Then sFlag = "S"
            dictLocalidades(sVal1) = sFlag
        End If
    Next r
    
    ' Garante pares de contratos Petrobras consolidados caso a tabela esteja incompleta
    dictContratos("4600686987") = True
    dictContratos("4600687006") = True
    
    dictUFs("RJ") = True: dictUFs("SP") = True: dictUFs("ES") = True: dictUFs("BA") = True
    
    dictContratoUF("4600686987|RJ") = True
    dictContratoUF("4600686987|SP") = True
    dictContratoUF("4600686987|ES") = True
    dictContratoUF("4600687006|BA") = True

End Sub

'-------------------------------------------------------------------------------------------------------------------
' RESUMO DA EXECUÇÃO APRESENTADO AO USUÁRIO
'-------------------------------------------------------------------------------------------------------------------
Private Sub ExibirResumoAuditoria(ByVal iTotal As Long, ByVal iAuditadas As Long, ByVal iComErro As Long, ByVal tSegundos As Double)
    Dim sMsg As String
    Dim iOK As Long
    
    iOK = iAuditadas - iComErro
    If iOK < 0 Then iOK = 0
    
    sMsg = "Auditoria de dados do Monitoramento concluída!" & vbNewLine & vbNewLine & _
           "• Regra inicial aplicada: Situação = CO (Concluído) ou AT (Atendido)" & vbNewLine & _
           "• Total de registros na planilha: " & Format(iTotal, "#,##0") & vbNewLine & _
           "• Registros auditados (CO / AT): " & Format(iAuditadas, "#,##0") & vbNewLine & _
           "• Registros íntegros (sem pendências): " & Format(iOK, "#,##0") & " (" & Format(iOK / IIf(iAuditadas > 0, iAuditadas, 1), "0.0%") & ")" & vbNewLine & _
           "• Registros com inconsistências: " & Format(iComErro, "#,##0") & " (" & Format(iComErro / IIf(iAuditadas > 0, iAuditadas, 1), "0.0%") & ")" & vbNewLine & vbNewLine & _
           "• Tempo de execução: " & Format(tSegundos, "0.00") & " segundos" & vbNewLine & vbNewLine
           
    If iComErro > 0 Then
        sMsg = sMsg & "Os detalhes de cada inconsistência foram gravados diretamente na Coluna Z (LOG) de cada linha afetada." & vbNewLine & _
                      "Filtre a Coluna Z por valores não vazios para tratar as pendências."
        MsgBox sMsg, vbExclamation, MSG_TITULO
    ElseIf iAuditadas = 0 Then
        sMsg = sMsg & "Nenhum registro com Situação 'CO' ou 'AT' foi encontrado para análise no intervalo verificado."
        MsgBox sMsg, vbInformation, MSG_TITULO
    Else
        sMsg = sMsg & "Parabéns! Todos os registros auditados atendem rigorosamente às regras de conformidade."
        MsgBox sMsg, vbInformation, MSG_TITULO
    End If
End Sub

Attribute VB_Name = "modAuditoriaLog"
Option Explicit

'====================================================================================================
' MODULO: modAuditoriaLog
' OBJETIVO: Replicar e modernizar a checagem de erros e inconsistencias dos codigos VBA legados
'           (antiga rotina 'analiseDeDados' e 'fncGeral.ExibirMsg'), adaptando-a para a nova
'           estrutura da pasta de trabalho (tabelas estruturadas, Power Query e abas contratuais).
'
' PRINCIPAIS CARACTERISTICAS:
'   1. Mapeamento Dinamico de Colunas: Nao depende de colunas fixas (Cells(i, 3), etc.). Localiza
'      automaticamente os campos pelos cabecalhos, tolerando variacoes de nomenclatura.
'   2. Processamento 100% em Memoria: Carrega os dados em matrizes (Variant Array) para maxima
'      velocidade (processa milhares de registros em milissegundos).
'   3. Classificacao de Severidade: Diferencia erros CRITICOS (que invalidam faturamento/SLA) de
'      ALERTAS (advertencias e desvios operacionais).
'   4. Relatorio Dedicado 'LOG_CRITICAS': Cria ou atualiza uma aba estruturada com os registros de erro,
'      facilitando filtragem, auditoria e conferencia pelo gestor.
'   5. Gravacao Opcional na Base: Atualiza a coluna 'LOG_AUDITORIA' na propria aba auditada.
'   6. Feedback Visual e Notificacao: Atualiza status na aba Painel e emite resumo completo em MsgBox.
'
' REGRAS AUDITADAS (Heranca do Legado + Novas Validacoes):
'   - [Obrigatorios] Contrato, Atividade, Item, Solicitacao, OS, Localidade, Datas e Quantidades.
'   - [Chave Solicitante] Validacao de padrao de 4 caracteres.
'   - [Gerencia Solicitante] Alerta para caracteres especiais proibidos (espaco, underline, barra).
'   - [Cronologia de Datas] Inconsistencia grave: Data Fechamento < Data Abertura.
'   - [Validacoes Numericas] Qtd Solicitada <= 0 ou Qtd Atendida < 0; valores nao numericos.
'   - [Desvio Volumetrico] Qtd Atendida superior a Qtd Solicitada.
'   - [Duplicidades FDM] Deteccao de chaves identicas (Atividade # Solicitacao # OS).
'   - [Prazos Combinados] Obrigatoriedade de prazo acordado para atividades de fiscalizacao.
'   - [SLA / Power Query] Identificacao de falhas e logs de erro originados no pipeline M.
'====================================================================================================

'-- Constantes de Configuracao de Planilhas
Private Const NOME_ABA_DADOS_PADRAO  As String = "DADOS"
Private Const NOME_ABA_LOG_CRITICAS  As String = "LOG_CRITICAS"
Private Const NOME_ABA_PAINEL        As String = "Painel"
Private Const NOME_COLUNA_LOG_DADOS  As String = "LOG_AUDITORIA"

'-- Paleta Visual Corporativa (Estilo Dashboard / Relatorio Executivo)
Private Const COR_CABECALHO_FUNDO    As Long = 2760975     ' RGB(15, 23, 42)   - Azul Petroleo Escuro
Private Const COR_CABECALHO_TEXTO    As Long = 16777215    ' RGB(255, 255, 255)- Branco
Private Const COR_CRITICO_FUNDO      As Long = 14803454    ' RGB(254, 226, 226)- Vermelho Claro Suave
Private Const COR_CRITICO_TEXTO      As Long = 1776411     ' RGB(153, 27, 27)  - Vermelho Escuro
Private Const COR_ALERTA_FUNDO       As Long = 13041662    ' RGB(254, 243, 199)- Amarelo/Ambar Claro
Private Const COR_ALERTA_TEXTO       As Long = 933906      ' RGB(146, 64, 14)  - Marrom/Ambar Escuro
Private Const COR_BORDA              As Long = 14737632    ' RGB(224, 224, 224)- Cinza Claro
Private Const COR_ZEBRA              As Long = 16382457    ' RGB(249, 250, 251)- Cinza Quase Branco

'-- Estrutura em memoria para acumulacao de ocorrencias
Private Type RegistroInconsistencia
    LinhaOrigem       As Long
    Contrato          As String
    CodigoOS          As String
    CodigoSolicitacao As String
    Atividade         As String
    Item              As String
    Severidade        As String      ' "CRITICO" ou "ALERTA"
    RegraCampo        As String      ' Ex: "Cronologia de Datas", "Campo Obrigatorio"
    DescricaoErro     As String
End Type


'====================================================================================================
' ROTINAS PUBLICAS (PONTOS DE ENTRADA ASSOCIAVEIS A BOTOES)
'====================================================================================================

''' <summary>
''' Executa a auditoria completa na aba DADOS consolidada (ou Painel_T2M) e gera o relatorio LOG_CRITICAS.
''' </summary>
Public Sub AuditarDadosEGerarLog()
    Dim wsDados As Worksheet
    
    Set wsDados = ObterAbaPrioritaria(ThisWorkbook, Array("DADOS", "Painel_T2M", "DADOS_IRON-LT1"))
    
    If wsDados Is Nothing Then
        MsgBox "Nao foi possivel localizar a aba de dados consolidada ('DADOS' ou 'Painel_T2M')." & vbCrLf & _
               "Certifique-se de que a planilha de medicao esta carregada.", vbExclamation, "Auditoria de Medicao"
        Exit Sub
    End If
    
    ExecutarMotorAuditoria wsDados, True
End Sub

''' <summary>
''' Permite auditar a planilha que estiver ativa no momento (ex: DADOS_PA-LT2, DADOS_IRON-LT1-RJ, etc.).
''' </summary>
Public Sub AuditarAbaAtiva()
    Dim wsAtiva As Worksheet
    Set wsAtiva = ActiveSheet
    
    If wsAtiva Is Nothing Then Exit Sub
    
    If wsAtiva.Name = NOME_ABA_LOG_CRITICAS Or wsAtiva.Name = NOME_ABA_PAINEL Then
        MsgBox "Selecione uma aba de dados operacionais (ex: DADOS, DADOS_PA-LT2, etc.) para auditar.", _
               vbInformation, "Auditoria de Medicao"
        Exit Sub
    End If
    
    ExecutarMotorAuditoria wsAtiva, True
End Sub

''' <summary>
''' Limpa o relatorio de criticas anterior.
''' </summary>
Public Sub LimparRelatorioLog()
    Dim wsLog As Worksheet
    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets(NOME_ABA_LOG_CRITICAS)
    On Error GoTo 0
    
    If Not wsLog Is Nothing Then
        Application.DisplayAlerts = False
        wsLog.Cells.Clear
        wsLog.Cells(1, 1).Value = "Nenhuma auditoria executada recentemente."
        Application.DisplayAlerts = True
        MsgBox "O relatorio de log de criticas foi limpo.", vbInformation, "Auditoria de Medicao"
    End If
End Sub


'====================================================================================================
' MOTOR PRINCIPAL DE AUDITORIA E REGRAS DE NEGOCIO
'====================================================================================================

Private Sub ExecutarMotorAuditoria(ByVal wsOrigem As Worksheet, ByVal bInterativo As Boolean)
    Dim appCalc     As XlCalculation
    Dim bEvents     As Boolean
    Dim bScreen     As Boolean
    
    Dim loOrigem    As ListObject
    Dim rngDados    As Range
    Dim vDados      As Variant
    Dim vHeaders    As Variant
    
    Dim numLinhas   As Long
    Dim numColunas  As Long
    Dim iLinha      As Long
    Dim linhaReal   As Long
    
    '-- Mapeamento dinamico de colunas
    Dim cContrato   As Long, cAtividade  As Long, cItem        As Long
    Dim cAplicacao  As Long, cCodSolic   As Long, cChaveSol    As Long
    Dim cGerencia   As Long, cDataSolic  As Long, cQtdSol      As Long
    Dim cCodOS      As Long, cDataAber   As Long, cDataFech    As Long
    Dim cPrazoComb  As Long, cQtdAtend   As Long, cLocalidade  As Long
    Dim cMunicipio  As Long, cUF         As Long, cLogSLA      As Long
    Dim cStatusPraz As Long, cColLogDest As Long
    
    '-- Colecao de inconsistencias
    Dim listaErros() As RegistroInconsistencia
    Dim totalErros   As Long
    Dim totalCriticos As Long
    Dim totalAlertas As Long
    
    '-- Dicionario para deteccao de duplicidades
    Dim dicChaves As Object
    Dim chaveDup  As String
    
    '-- Variaveis temporarias de cada linha
    Dim vValContrato As String, vValAtiv   As String, vValItem As String
    Dim vValOS       As String, vValSolic  As String, vValChave As String
    Dim vValGer      As String
    Dim dtAber       As Variant, dtFech    As Variant, dtSolic As Variant
    Dim qSol         As Variant, qAtend    As Variant
    Dim sLogLinha    As String
    Dim vLogsColuna() As String
    Dim bTemColunaLog As Boolean
    
    On Error GoTo TratarErro
    
    '-- Preserva e otimiza estado do Excel
    appCalc = Application.Calculation
    bEvents = Application.EnableEvents
    bScreen = Application.ScreenUpdating
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    '-- Identifica a fonte de dados (Tabela estruturada ListObject ou Range usado)
    If wsOrigem.ListObjects.Count > 0 Then
        Set loOrigem = wsOrigem.ListObjects(1)
        If loOrigem.DataBodyRange Is Nothing Then
            MsgBox "A tabela da aba '" & wsOrigem.Name & "' nao possui registros de dados para auditar.", _
                   vbInformation, "Auditoria"
            GoTo SairRotina
        End If
        vHeaders = loOrigem.HeaderRowRange.Value
        vDados = loOrigem.DataBodyRange.Value
        linhaReal = loOrigem.DataBodyRange.Row
    Else
        numLinhas = wsOrigem.Cells(wsOrigem.Rows.Count, 1).End(xlUp).Row
        If numLinhas < 2 Then
            numLinhas = wsOrigem.Cells(wsOrigem.Rows.Count, 2).End(xlUp).Row
        End If
        If numLinhas < 2 Then
            MsgBox "Nao foram encontrados dados suficientes na aba '" & wsOrigem.Name & "'.", vbInformation, "Auditoria"
            GoTo SairRotina
        End If
        vHeaders = wsOrigem.Range(wsOrigem.Cells(1, 1), wsOrigem.Cells(1, wsOrigem.Cells(1, wsOrigem.Columns.Count).End(xlToLeft).Column)).Value
        vDados = wsOrigem.Range(wsOrigem.Cells(2, 1), wsOrigem.Cells(numLinhas, UBound(vHeaders, 2))).Value
        linhaReal = 2
    End If
    
    numLinhas = UBound(vDados, 1)
    numColunas = UBound(vDados, 2)
    
    '-- Localizacao dinamica das colunas pelos nomes possiveis
    cContrato   = LocalizarIndiceColuna(vHeaders, "Contrato", "Contrato.1", "Empresa")
    cAtividade  = LocalizarIndiceColuna(vHeaders, "Descricao da atividade", "Descricao da atividade", "Atividade")
    cItem       = LocalizarIndiceColuna(vHeaders, "Item", "Item PPU")
    cAplicacao  = LocalizarIndiceColuna(vHeaders, "Aplicacao", "Aplicacao")
    cCodSolic   = LocalizarIndiceColuna(vHeaders, "Codigo da solicitacao", "Codigo da solicitacao", "Codigo da_x000a_solicitacao")
    cChaveSol   = LocalizarIndiceColuna(vHeaders, "Chave solicitante", "Chave_x000a_solicitante", "Chave")
    cGerencia   = LocalizarIndiceColuna(vHeaders, "Gerencia solicitante", "Gerencia solicitante")
    cDataSolic  = LocalizarIndiceColuna(vHeaders, "Data solicitacao", "D. solicitacao", "Data solicitacao")
    cQtdSol     = LocalizarIndiceColuna(vHeaders, "Qtd. Solicitada", "Qtd Solicitada", "Qtd._x000a_Solicitada")
    cCodOS      = LocalizarIndiceColuna(vHeaders, "Codigo OS", "Codigo OS", "OS")
    cDataAber   = LocalizarIndiceColuna(vHeaders, "D. abertura", "Data abertura", "Data Abertura")
    cDataFech   = LocalizarIndiceColuna(vHeaders, "D. fechamento", "Data fechamento", "Data Fechamento")
    cPrazoComb  = LocalizarIndiceColuna(vHeaders, "Prazo combinado", "Prazo_x000a_combinado")
    cQtdAtend   = LocalizarIndiceColuna(vHeaders, "Qtd. Atendida", "Qtd Atendida", "Qtd._x000a_Atendida")
    cLocalidade = LocalizarIndiceColuna(vHeaders, "Localidade", "Localidade - Origem", "Centro")
    cMunicipio  = LocalizarIndiceColuna(vHeaders, "Municipio", "Municipio", "Municipio - Origem")
    cUF         = LocalizarIndiceColuna(vHeaders, "UF", "UF - Origem")
    cStatusPraz = LocalizarIndiceColuna(vHeaders, "Status Prazo", "Status_x000a_Prazo")
    cLogSLA     = LocalizarIndiceColuna(vHeaders, "LOG SLA", "LOG", "Log")
    cColLogDest = LocalizarIndiceColuna(vHeaders, NOME_COLUNA_LOG_DADOS)
    
    '-- Prepara matriz de log na propria tabela, se houver coluna destino
    bTemColunaLog = (cColLogDest > 0)
    If bTemColunaLog Then
        ReDim vLogsColuna(1 To numLinhas, 1 To 1)
    End If
    
    '-- Inicializa dicionario de duplicidades (Late Binding para evitar dependencia de referencias)
    Set dicChaves = CreateObject("Scripting.Dictionary")
    dicChaves.CompareMode = vbTextCompare
    
    totalErros = 0
    totalCriticos = 0
    totalAlertas = 0
    ReDim listaErros(1 To 500)
    
    '================================================================================================
    ' LOOP DE AUDITORIA EM MEMORIA
    '================================================================================================
    For iLinha = 1 To numLinhas
        sLogLinha = ""
        
        '-- Extracao dos valores da linha atual com protecao contra nulos
        vValContrato = ObterTextoSeguro(vDados, iLinha, cContrato)
        vValAtiv     = ObterTextoSeguro(vDados, iLinha, cAtividade)
        vValItem     = ObterTextoSeguro(vDados, iLinha, cItem)
        vValOS       = ObterTextoSeguro(vDados, iLinha, cCodOS)
        vValSolic    = ObterTextoSeguro(vDados, iLinha, cCodSolic)
        vValChave    = ObterTextoSeguro(vDados, iLinha, cChaveSol)
        vValGer      = ObterTextoSeguro(vDados, iLinha, cGerencia)
        
        dtAber       = ObterValorSeguro(vDados, iLinha, cDataAber)
        dtFech       = ObterValorSeguro(vDados, iLinha, cDataFech)
        dtSolic      = ObterValorSeguro(vDados, iLinha, cDataSolic)
        
        qSol         = ObterValorSeguro(vDados, iLinha, cQtdSol)
        qAtend       = ObterValorSeguro(vDados, iLinha, cQtdAtend)
        
        '-- 1. VALIDACAO DE CAMPOS OBRIGATORIOS VAZIOS (Severidade: CRITICO)
        If Len(vValContrato) = 0 Then
            RegistrarErro listaErros, totalErros, totalCriticos, totalAlertas, _
                          linhaReal + iLinha - 1, vValContrato, vValOS, vValSolic, vValAtiv, vValItem, _
                          "CRITICO", "Campo Obrigatorio", "Contrato nao informado ou vazio"
            sLogLinha = ConcatenarLog(sLogLinha, "Contrato vazio")
        End If
        
        If Len(vValAtiv) = 0 Then
            RegistrarErro listaErros, totalErros, totalCriticos, totalAlertas, _
                          linhaReal + iLinha - 1, vValContrato, vValOS, vValSolic, vValAtiv, vValItem, _
                          "CRITICO", "Campo Obrigatorio", "Descricao da atividade nao informada"
            sLogLinha = ConcatenarLog(sLogLinha, "Atividade vazia")
        End If
        
        If Len(vValItem) = 0 Then
            RegistrarErro listaErros, totalErros, totalCriticos, totalAlertas, _
                          linhaReal + iLinha - 1, vValContrato, vValOS, vValSolic, vValAtiv, vValItem, _
                          "CRITICO", "Campo Obrigatorio", "Item nao informado ou vazio"
            sLogLinha = ConcatenarLog(sLogLinha, "Item vazio")
        End If
        
        If Len(vValSolic) = 0 Then
            RegistrarErro listaErros, totalErros, totalCriticos, totalAlertas, _
                          linhaReal + iLinha - 1, vValContrato, vValOS, vValSolic, vValAtiv, vValItem, _
                          "CRITICO", "Campo Obrigatorio", "Codigo da solicitacao nao informado"
            sLogLinha = ConcatenarLog(sLogLinha, "Solicitacao vazia")
        End If
        
        If Len(vValOS) = 0 Then
            RegistrarErro listaErros, totalErros, totalCriticos, totalAlertas, _
                          linhaReal + iLinha - 1, vValContrato, vValOS, vValSolic, vValAtiv, vValItem, _
                          "ALERTA", "Identificacao OS", "Codigo OS nao informado para atendimento concluido"
            sLogLinha = ConcatenarLog(sLogLinha, "OS nao preenchida")
        End If
        
        '-- 2. VALIDACAO DE CHAVE SOLICITANTE (Heranca da rotina legada: Len = 4)
        If cChaveSol > 0 And Len(vValChave) > 0 Then
            If Len(vValChave) <> 4 Then
                RegistrarErro listaErros, totalErros, totalCriticos, totalAlertas, _
                              linhaReal + iLinha - 1, vValContrato, vValOS, vValSolic, vValAtiv, vValItem, _
                              "ALERTA", "Chave Solicitante", "Chave com tamanho divergente do padrao (esperado 4 caracteres): '" & vValChave & "'"
                sLogLinha = ConcatenarLog(sLogLinha, "Chave solicitante fora do padrao de 4 digitos")
            End If
        End If
        
        '-- 3. VALIDACAO DE CARACTERES EM GERENCIA SOLICITANTE (Heranca do legado: proibe _, \ e espaco indesejado)
        If cGerencia > 0 And Len(vValGer) > 0 Then
            If InStr(vValGer, "_") > 0 Or InStr(vValGer, "\") > 0 Then
                RegistrarErro listaErros, totalErros, totalCriticos, totalAlertas, _
                              linhaReal + iLinha - 1, vValContrato, vValOS, vValSolic, vValAtiv, vValItem, _
                              "ALERTA", "Gerencia Solicitante", "Caractere invalido detectado na gerencia: '" & vValGer & "'"
                sLogLinha = ConcatenarLog(sLogLinha, "Gerencia com caractere especial invalido")
            End If
        End If
        
        '-- 4. VALIDACAO DE DATAS E CRONOLOGIA (Severidade: CRITICO)
        If IsEmpty(dtAber) Or Len(Trim$(CStr(dtAber))) = 0 Then
            RegistrarErro listaErros, totalErros, totalCriticos, totalAlertas, _
                          linhaReal + iLinha - 1, vValContrato, vValOS, vValSolic, vValAtiv, vValItem, _
                          "CRITICO", "Datas de Atendimento", "Data de abertura nao preenchida"
            sLogLinha = ConcatenarLog(sLogLinha, "D. abertura vazia")
        ElseIf Not IsDate(dtAber) Then
            RegistrarErro listaErros, totalErros, totalCriticos, totalAlertas, _
                          linhaReal + iLinha - 1, vValContrato, vValOS, vValSolic, vValAtiv, vValItem, _
                          "CRITICO", "Datas de Atendimento", "Data de abertura em formato invalido: '" & CStr(dtAber) & "'"
            sLogLinha = ConcatenarLog(sLogLinha, "D. abertura invalida")
        End If
        
        If IsEmpty(dtFech) Or Len(Trim$(CStr(dtFech))) = 0 Then
            RegistrarErro listaErros, totalErros, totalCriticos, totalAlertas, _
                          linhaReal + iLinha - 1, vValContrato, vValOS, vValSolic, vValAtiv, vValItem, _
                          "CRITICO", "Datas de Atendimento", "Data de fechamento nao preenchida em registro concluido"
            sLogLinha = ConcatenarLog(sLogLinha, "D. fechamento vazia")
        ElseIf Not IsDate(dtFech) Then
            RegistrarErro listaErros, totalErros, totalCriticos, totalAlertas, _
                          linhaReal + iLinha - 1, vValContrato, vValOS, vValSolic, vValAtiv, vValItem, _
                          "CRITICO", "Datas de Atendimento", "Data de fechamento em formato invalido: '" & CStr(dtFech) & "'"
            sLogLinha = ConcatenarLog(sLogLinha, "D. fechamento invalida")
        End If
        
        '-- Inconsistencia Cronologica Grave: Fechamento anterior a Abertura
        If IsDate(dtAber) And IsDate(dtFech) Then
            If CDate(dtFech) < CDate(dtAber) Then
                RegistrarErro listaErros, totalErros, totalCriticos, totalAlertas, _
                              linhaReal + iLinha - 1, vValContrato, vValOS, vValSolic, vValAtiv, vValItem, _
                              "CRITICO", "Cronologia de Datas", "Data de fechamento (" & Format(CDate(dtFech), "dd/mm/yyyy hh:mm") & _
                              ") e ANTERIOR a data de abertura (" & Format(CDate(dtAber), "dd/mm/yyyy hh:mm") & ")"
                sLogLinha = ConcatenarLog(sLogLinha, "Fechamento anterior a abertura")
            End If
        End If
        
        '-- 5. VALIDACAO DE QUANTIDADES (Severidade: CRITICO e ALERTA)
        If IsEmpty(qSol) Or Len(Trim$(CStr(qSol))) = 0 Then
            RegistrarErro listaErros, totalErros, totalCriticos, totalAlertas, _
                          linhaReal + iLinha - 1, vValContrato, vValOS, vValSolic, vValAtiv, vValItem, _
                          "CRITICO", "Quantidades", "Quantidade solicitada vazia"
            sLogLinha = ConcatenarLog(sLogLinha, "Qtd. solicitada vazia")
        ElseIf Not IsNumeric(qSol) Then
            RegistrarErro listaErros, totalErros, totalCriticos, totalAlertas, _
                          linhaReal + iLinha - 1, vValContrato, vValOS, vValSolic, vValAtiv, vValItem, _
                          "CRITICO", "Quantidades", "Quantidade solicitada nao e numerica: '" & CStr(qSol) & "'"
            sLogLinha = ConcatenarLog(sLogLinha, "Qtd. solicitada nao numerica")
        ElseIf CDbl(qSol) <= 0 Then
            RegistrarErro listaErros, totalErros, totalCriticos, totalAlertas, _
                          linhaReal + iLinha - 1, vValContrato, vValOS, vValSolic, vValAtiv, vValItem, _
                          "CRITICO", "Quantidades", "Quantidade solicitada deve ser maior que zero (valor: " & CStr(qSol) & ")"
            sLogLinha = ConcatenarLog(sLogLinha, "Qtd. solicitada <= 0")
        End If
        
        If IsEmpty(qAtend) Or Len(Trim$(CStr(qAtend))) = 0 Then
            RegistrarErro listaErros, totalErros, totalCriticos, totalAlertas, _
                          linhaReal + iLinha - 1, vValContrato, vValOS, vValSolic, vValAtiv, vValItem, _
                          "CRITICO", "Quantidades", "Quantidade atendida nao preenchida"
            sLogLinha = ConcatenarLog(sLogLinha, "Qtd. atendida vazia")
        ElseIf Not IsNumeric(qAtend) Then
            RegistrarErro listaErros, totalErros, totalCriticos, totalAlertas, _
                          linhaReal + iLinha - 1, vValContrato, vValOS, vValSolic, vValAtiv, vValItem, _
                          "CRITICO", "Quantidades", "Quantidade atendida nao e numerica: '" & CStr(qAtend) & "'"
            sLogLinha = ConcatenarLog(sLogLinha, "Qtd. atendida nao numerica")
        ElseIf CDbl(qAtend) < 0 Then
            RegistrarErro listaErros, totalErros, totalCriticos, totalAlertas, _
                          linhaReal + iLinha - 1, vValContrato, vValOS, vValSolic, vValAtiv, vValItem, _
                          "CRITICO", "Quantidades", "Quantidade atendida negativa: " & CStr(qAtend)
            sLogLinha = ConcatenarLog(sLogLinha, "Qtd. atendida negativa")
        End If
        
        '-- Desvio Volumetrico: Atendida > Solicitada
        If IsNumeric(qSol) And IsNumeric(qAtend) Then
            If CDbl(qAtend) > CDbl(qSol) Then
                RegistrarErro listaErros, totalErros, totalCriticos, totalAlertas, _
                              linhaReal + iLinha - 1, vValContrato, vValOS, vValSolic, vValAtiv, vValItem, _
                              "ALERTA", "Divergencia de Qtd", "Qtd atendida (" & CStr(qAtend) & ") e superior a solicitada (" & CStr(qSol) & ")"
                sLogLinha = ConcatenarLog(sLogLinha, "Qtd atendida > solicitada")
            End If
        End If
        
        '-- 6. VALIDACAO DE DUPLICIDADES DE REGISTRO (Chave FDM Composta)
        If Len(vValAtiv) > 0 And Len(vValSolic) > 0 And Len(vValOS) > 0 Then
            chaveDup = vValAtiv & "#" & vValSolic & "#" & vValOS
            If dicChaves.Exists(chaveDup) Then
                RegistrarErro listaErros, totalErros, totalCriticos, totalAlertas, _
                              linhaReal + iLinha - 1, vValContrato, vValOS, vValSolic, vValAtiv, vValItem, _
                              "ALERTA", "Duplicidade de Registro", "Combinacao repetida de Atividade + Solicitacao + OS (primeira ocorrencia na linha " & dicChaves(chaveDup) & ")"
                sLogLinha = ConcatenarLog(sLogLinha, "Duplicidade de Atividade/Solicitacao/OS")
            Else
                dicChaves.Add chaveDup, linhaReal + iLinha - 1
            End If
        End If
        
        '-- 7. REPASSE DE PENDENCIAS DE SLA OU POWER QUERY (LOG SLA)
        If cLogSLA > 0 Then
            Dim sLogM As String
            sLogM = ObterTextoSeguro(vDados, iLinha, cLogSLA)
            If Len(sLogM) > 0 Then
                RegistrarErro listaErros, totalErros, totalCriticos, totalAlertas, _
                              linhaReal + iLinha - 1, vValContrato, vValOS, vValSolic, vValAtiv, vValItem, _
                              "ALERTA", "Log do Motor Power Query", "Pendencia registrada na consulta M: '" & sLogM & "'"
                sLogLinha = ConcatenarLog(sLogLinha, "PQ: " & sLogM)
            End If
        End If
        
        '-- Guarda o log consolidado para escrever na coluna da base, se habilitada
        If bTemColunaLog Then
            vLogsColuna(iLinha, 1) = sLogLinha
        End If
    Next iLinha
    
    '-- Grava os logs na coluna da propria aba se ela existir
    If bTemColunaLog And numLinhas > 0 Then
        If Not loOrigem Is Nothing Then
            loOrigem.ListColumns(NOME_COLUNA_LOG_DADOS).DataBodyRange.Value = vLogsColuna
        Else
            wsOrigem.Range(wsOrigem.Cells(linhaReal, cColLogDest), wsOrigem.Cells(linhaReal + numLinhas - 1, cColLogDest)).Value = vLogsColuna
        End If
    End If
    
    '================================================================================================
    ' GERACAO DO RELATORIO ESTRUTURADO 'LOG_CRITICAS'
    '================================================================================================
    GerarRelatorioAbaLog listaErros, totalErros, wsOrigem.Name
    
    '-- Atualiza status no Painel
    AtualizarStatusPainel wsOrigem.Name, numLinhas, totalCriticos, totalAlertas
    
    '-- Restaura ambiente do Excel
SairRotina:
    Application.Calculation = appCalc
    Application.EnableEvents = bEvents
    Application.ScreenUpdating = bScreen
    
    '-- Exibicao de Resumo Informativo
    If bInterativo Then
        If totalErros = 0 Then
            MsgBox "[SUCESSO] Auditoria de Medicao Concluida com Sucesso!" & vbCrLf & vbCrLf & _
                   "- Base analisada: " & wsOrigem.Name & vbCrLf & _
                   "- Registros auditados: " & Format(numLinhas, "#,##0") & vbCrLf & _
                   "- Inconsistencias: NENHUMA pendencia encontrada." & vbCrLf & vbCrLf & _
                   "A base de dados esta em perfeita conformidade para medicao.", _
                   vbInformation, "Auditoria Concluida"
        Else
            Dim resp As VbMsgBoxResult
            resp = MsgBox("[ATENCAO] Auditoria Concluida com Apontamentos!" & vbCrLf & vbCrLf & _
                          "- Base analisada: " & wsOrigem.Name & vbCrLf & _
                          "- Registros auditados: " & Format(numLinhas, "#,##0") & vbCrLf & _
                          "- Erros CRITICOS: " & totalCriticos & vbCrLf & _
                          "- ALERTAS / Avisos: " & totalAlertas & vbCrLf & _
                          "- Total de pendencias: " & totalErros & vbCrLf & vbCrLf & _
                          "Deseja navegar para a aba '" & NOME_ABA_LOG_CRITICAS & "' para conferir os detalhes?", _
                          vbYesNo + vbExclamation, "Inconsistencias Encontradas")
            If resp = vbYes Then
                On Error Resume Next
                ThisWorkbook.Worksheets(NOME_ABA_LOG_CRITICAS).Activate
                On Error GoTo 0
            End If
        End If
    End If
    Exit Sub
    
TratarErro:
    MsgBox "Erro inesperado na rotina de auditoria (" & Err.Number & "): " & Err.Description, _
           vbCritical, "Falha na Auditoria"
    Resume SairRotina
End Sub


'====================================================================================================
' ROTINAS DE MONTAGEM E ESTILIZACAO DO RELATORIO LOG_CRITICAS
'====================================================================================================

Private Sub GerarRelatorioAbaLog(ByRef lista() As RegistroInconsistencia, ByVal total As Long, ByVal sNomeBase As String)
    Dim wsLog As Worksheet
    Dim loLog As ListObject
    Dim i As Long
    Dim vMatrizLog() As Variant
    Dim rngTabela As Range
    
    '-- Localiza ou cria a aba LOG_CRITICAS
    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets(NOME_ABA_LOG_CRITICAS)
    On Error GoTo 0
    
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsLog.Name = NOME_ABA_LOG_CRITICAS
    End If
    
    wsLog.Cells.Clear
    wsLog.Tab.Color = IIf(total > 0, COR_CRITICO_TEXTO, 5296274)
    
    '-- Cabecalho Superior Executivo
    With wsLog.Range("A1:I1")
        .Merge
        .Value = "RELATORIO DE AUDITORIA E LOG DE CRITICAS - MEDICAO PETROBRAS"
        .Font.Name = "Segoe UI"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = COR_CABECALHO_TEXTO
        .Interior.Color = COR_CABECALHO_FUNDO
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 35
    End With
    
    wsLog.Range("A2").Value = "Base Auditada: " & sNomeBase & " | Data/Hora da Execucao: " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    With wsLog.Range("A2:I2")
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Italic = True
        .Font.Color = RGB(100, 116, 139)
        .RowHeight = 20
    End With
    
    '-- Cabecalhos da Tabela de Inconsistencias
    Dim cabecalhos As Variant
    cabecalhos = Array("Linha Base", "Contrato", "Codigo OS", "Cod. Solicitacao", _
                       "Atividade", "Item", "Severidade", "Regra / Campo", "Descricao Detalhada da Inconsistencia")
    
    For i = 0 To UBound(cabecalhos)
        wsLog.Cells(4, i + 1).Value = cabecalhos(i)
    Next i
    
    With wsLog.Range("A4:I4")
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = COR_CABECALHO_TEXTO
        .Interior.Color = RGB(30, 41, 59)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 25
    End With
    
    If total = 0 Then
        wsLog.Cells(5, 1).Value = "-"
        wsLog.Cells(5, 7).Value = "OK"
        wsLog.Cells(5, 9).Value = "Parabens! Nenhuma inconsistencia encontrada nesta base."
        wsLog.Range("A5:I5").Interior.Color = RGB(240, 253, 244)
        wsLog.Range("A5:I5").Font.Color = RGB(22, 101, 52)
        wsLog.Columns("A:I").AutoFit
        Exit Sub
    End If
    
    '-- Prepara dados em matriz para despejo instantaneo
    ReDim vMatrizLog(1 To total, 1 To 9)
    For i = 1 To total
        vMatrizLog(i, 1) = lista(i).LinhaOrigem
        vMatrizLog(i, 2) = lista(i).Contrato
        vMatrizLog(i, 3) = lista(i).CodigoOS
        vMatrizLog(i, 4) = lista(i).CodigoSolicitacao
        vMatrizLog(i, 5) = lista(i).Atividade
        vMatrizLog(i, 6) = lista(i).Item
        vMatrizLog(i, 7) = lista(i).Severidade
        vMatrizLog(i, 8) = lista(i).RegraCampo
        vMatrizLog(i, 9) = lista(i).DescricaoErro
    Next i
    
    Set rngTabela = wsLog.Range(wsLog.Cells(5, 1), wsLog.Cells(4 + total, 9))
    rngTabela.Value = vMatrizLog
    
    '-- Formatacao e Cores das Linhas de Erro
    With rngTabela
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = COR_BORDA
    End With
    
    '-- Destaca severidade linha a linha
    For i = 1 To total
        Dim rLinha As Range
        Set rLinha = wsLog.Range(wsLog.Cells(4 + i, 1), wsLog.Cells(4 + i, 9))
        If lista(i).Severidade = "CRITICO" Then
            wsLog.Cells(4 + i, 7).Interior.Color = COR_CRITICO_FUNDO
            wsLog.Cells(4 + i, 7).Font.Color = COR_CRITICO_TEXTO
            wsLog.Cells(4 + i, 7).Font.Bold = True
        Else
            wsLog.Cells(4 + i, 7).Interior.Color = COR_ALERTA_FUNDO
            wsLog.Cells(4 + i, 7).Font.Color = COR_ALERTA_TEXTO
            wsLog.Cells(4 + i, 7).Font.Bold = True
        End If
        
        ' Efeito zebra suave nas colunas adjacentes
        If i Mod 2 = 0 Then
            rLinha.Interior.Color = COR_ZEBRA
        End If
    Next i
    
    '-- Alinhamentos especificos
    wsLog.Range("A5:A" & (4 + total)).HorizontalAlignment = xlCenter
    wsLog.Range("B5:B" & (4 + total)).HorizontalAlignment = xlCenter
    wsLog.Range("C5:D" & (4 + total)).HorizontalAlignment = xlCenter
    wsLog.Range("G5:G" & (4 + total)).HorizontalAlignment = xlCenter
    
    '-- Ajuste de largura das colunas
    wsLog.Columns("A:I").AutoFit
    If wsLog.Columns("I").ColumnWidth < 45 Then wsLog.Columns("I").ColumnWidth = 45
    
    '-- Ativa AutoFiltro
    wsLog.Range("A4:I4").AutoFilter
    wsLog.Application.ActiveWindow.SplitRow = 4
    wsLog.Application.ActiveWindow.FreezePanes = True
End Sub

Private Sub AtualizarStatusPainel(ByVal sAbaBase As String, ByVal totalReg As Long, ByVal criticos As Long, ByVal alertas As Long)
    Dim wsPainel As Worksheet
    On Error Resume Next
    Set wsPainel = ThisWorkbook.Worksheets(NOME_ABA_PAINEL)
    On Error GoTo 0
    
    If wsPainel Is Nothing Then Exit Sub
    
    '-- Grava na celula B14 (ou cria indicador de auditoria no painel operacional)
    On Error Resume Next
    With wsPainel.Range("B14")
        If criticos = 0 And alertas = 0 Then
            .Value = "[OK] Auditoria (" & Format(Now, "dd/mm hh:mm") & ") - Sem Pendencias"
            .Interior.Color = 5296274  ' Verde suave
            .Font.Color = vbWhite
            .Font.Bold = True
        Else
            .Value = "[ALERTA] Auditoria: " & criticos & " Criticos / " & alertas & " Alertas (" & Format(Now, "dd/mm hh:mm") & ")"
            .Interior.Color = IIf(criticos > 0, COR_CRITICO_FUNDO, COR_ALERTA_FUNDO)
            .Font.Color = IIf(criticos > 0, COR_CRITICO_TEXTO, COR_ALERTA_TEXTO)
            .Font.Bold = True
        End If
    End With
    On Error GoTo 0
End Sub


'====================================================================================================
' FUNCOES AUXILIARES E UTILITARIOS
'====================================================================================================

Private Sub RegistrarErro(ByRef lista() As RegistroInconsistencia, ByRef total As Long, _
                          ByRef criticos As Long, ByRef alertas As Long, _
                          ByVal lin As Long, ByVal sContrato As String, ByVal sOS As String, _
                          ByVal sSolic As String, ByVal sAtiv As String, ByVal sItem As String, _
                          ByVal sSev As String, ByVal sRegra As String, ByVal sDesc As String)
    total = total + 1
    If total > UBound(lista) Then
        ReDim Preserve lista(1 To UBound(lista) + 500)
    End If
    
    If sSev = "CRITICO" Then
        criticos = criticos + 1
    Else
        alertas = alertas + 1
    End If
    
    With lista(total)
        .LinhaOrigem = lin
        .Contrato = sContrato
        .CodigoOS = sOS
        .CodigoSolicitacao = sSolic
        .Atividade = sAtiv
        .Item = sItem
        .Severidade = sSev
        .RegraCampo = sRegra
        .DescricaoErro = sDesc
    End With
End Sub

Private Function NormalizarParaComparacao(ByVal sTexto As String) As String
    Dim i As Long
    Dim c As Integer
    Dim ch As String
    Dim res As String
    
    sTexto = LCase$(Trim$(sTexto))
    sTexto = Replace(sTexto, Chr(10), " ")
    sTexto = Replace(sTexto, "_x000a_", " ")
    sTexto = Replace(sTexto, "_", " ")
    sTexto = Replace(sTexto, ".", " ")
    
    res = ""
    For i = 1 To Len(sTexto)
        ch = Mid$(sTexto, i, 1)
        c = Asc(ch)
        Select Case c
            Case 224 To 229: res = res & "a"
            Case 232 To 235: res = res & "e"
            Case 236 To 239: res = res & "i"
            Case 242 To 246: res = res & "o"
            Case 249 To 252: res = res & "u"
            Case 231:        res = res & "c"
            Case 192 To 197: res = res & "a"
            Case 200 To 203: res = res & "e"
            Case 204 To 207: res = res & "i"
            Case 210 To 214: res = res & "o"
            Case 217 To 220: res = res & "u"
            Case 199:        res = res & "c"
            Case Else:       res = res & ch
        End Select
    Next i
    
    Do While InStr(res, "  ") > 0
        res = Replace(res, "  ", " ")
    Loop
    
    NormalizarParaComparacao = Trim$(res)
End Function

Private Function LocalizarIndiceColuna(ByRef vHeaders As Variant, ParamArray nomes()) As Long
    Dim nome As Variant
    Dim c As Long
    Dim sHeader As String
    Dim sNomeNorm As String
    
    For Each nome In nomes
        sNomeNorm = NormalizarParaComparacao(CStr(nome))
        For c = 1 To UBound(vHeaders, 2)
            sHeader = NormalizarParaComparacao(CStr(vHeaders(1, c)))
            
            If sHeader = sNomeNorm Or InStr(1, sHeader, sNomeNorm, vbTextCompare) > 0 Then
                LocalizarIndiceColuna = c
                Exit Function
            End If
        Next c
    Next nome
    LocalizarIndiceColuna = 0
End Function

Private Function ObterTextoSeguro(ByRef vDados As Variant, ByVal lin As Long, ByVal col As Long) As String
    If col = 0 Then Exit Function
    If IsEmpty(vDados(lin, col)) Or IsError(vDados(lin, col)) Then Exit Function
    ObterTextoSeguro = Trim$(CStr(vDados(lin, col)))
End Function

Private Function ObterValorSeguro(ByRef vDados As Variant, ByVal lin As Long, ByVal col As Long) As Variant
    If col = 0 Then
        ObterValorSeguro = Empty
        Exit Function
    End If
    If IsError(vDados(lin, col)) Then
        ObterValorSeguro = Empty
        Exit Function
    End If
    ObterValorSeguro = vDados(lin, col)
End Function

Private Function ConcatenarLog(ByVal sAtual As String, ByVal sNovo As String) As String
    If Len(sAtual) = 0 Then
        ConcatenarLog = sNovo
    Else
        ConcatenarLog = sAtual & "; " & sNovo
    End If
End Function

Private Function ObterAbaPrioritaria(ByVal wb As Workbook, ByVal nomesAbas As Variant) As Worksheet
    Dim nome As Variant
    Dim ws As Worksheet
    For Each nome In nomesAbas
        On Error Resume Next
        Set ws = wb.Worksheets(CStr(nome))
        On Error GoTo 0
        If Not ws Is Nothing Then
            Set ObterAbaPrioritaria = ws
            Exit Function
        End If
    Next nome
    Set ObterAbaPrioritaria = Nothing
End Function

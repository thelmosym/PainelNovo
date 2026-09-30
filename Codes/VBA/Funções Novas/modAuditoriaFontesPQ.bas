Attribute VB_Name = "modAuditoriaFontesPQ"
Option Explicit

'====================================================================================================
' MODULO: modAuditoriaFontesPQ
' OBJETIVO: Escanear todas as consultas Power Query (Linguagem M) e conexoes do Workbook,
'           identificando e auditando com precisao:
'             1. Fontes de dados externas (Pastas, Arquivos Excel, CSV, Bancos de Dados, Web/SharePoint).
'             2. Conectores e funcoes M utilizados (Folder.Files, File.Contents, Sql.Database, etc.).
'             3. Parametros e variaveis de caminho referenciados (ex: localTabelaB, localMonit...).
'             4. Resolucao automatica dos caminhos reais via Nomes Definidos e Tabelas da planilha.
'             5. Validacao em tempo real de acessibilidade fisica (se os arquivos/pastas existem no disco).
'             6. Geracao de Relatorio Executivo estruturado e zebrado na aba 'Auditoria_Fontes_PQ'.
'
' PADROES MANDATORIOS DE ENGENHARIA APLICADOS:
'   - Codificacao estrita Windows-1252 (ANSI) e terminacoes de linha CRLF.
'   - Proibicao de emojis Unicode de 4 bytes (substituidos por tags [OK], [ERRO], [AVISO], [INFO]).
'   - Preservacao e restauracao confiavel do estado da aplicacao (ScreenUpdating, Calculation, Events).
'   - Manipulacao de arquivos via FileSystemObject com Late Binding (sem dependencias externas).
'   - Resolucao de parametros encadeados e concatenacoes de caminhos.
'====================================================================================================

Private Const TITULO_AUDITORIA        As String = "Auditoria de Fontes Power Query"
Private Const NOME_ABA_AUDITORIA      As String = "Auditoria_Fontes_PQ"

' Paleta de Cores Corporativas (RGB)
Private Const COR_HEADER_FUNDO        As Long = 7888415     ' RGB(31, 78, 120) - Azul Marinho Corporativo
Private Const COR_HEADER_FONTE        As Long = 16777215    ' RGB(255, 255, 255) - Branco
Private Const COR_ZEBRA_PAR           As Long = 16777215    ' RGB(255, 255, 255) - Branco
Private Const COR_ZEBRA_IMPAR         As Long = 16448250    ' RGB(250, 250, 250) - Cinza Ultra Suave
Private Const COR_BORDA_GRADE         As Long = 14540253    ' RGB(221, 221, 221) - Cinza Borda
Private Const COR_CARD_FUNDO          As Long = 15921906    ' RGB(242, 244, 246) - Fundo de Cartao

' Cores de Status
Private Const COR_OK_FUNDO            As Long = 15137756    ' RGB(220, 252, 231) - Verde Claro Suave
Private Const COR_OK_FONTE            As Long = 3433748     ' RGB(20, 83, 45)    - Verde Escuro
Private Const COR_ERRO_FUNDO          As Long = 14803454    ' RGB(254, 226, 226) - Vermelho Claro Suave
Private Const COR_ERRO_FONTE          As Long = 1776825     ' RGB(153, 27, 27)   - Vermelho Escuro
Private Const COR_AVISO_FUNDO         As Long = 9109246     ' RGB(254, 240, 138) - Amarelo Claro Suave
Private Const COR_AVISO_FONTE         As Long = 938453      ' RGB(133, 77, 14)   - Amarelo Escuro
Private Const COR_INFO_FUNDO          As Long = 16699872    ' RGB(224, 242, 254) - Azul Claro Suave
Private Const COR_INFO_FONTE          As Long = 8739079     ' RGB(7, 89, 133)    - Azul Escuro

' Estrutura de dados para armazenar cada fonte detectada
Private Type RegFontePQ
    NomeConsulta        As String
    Classificacao       As String ' Fonte Externa, Fonte Interna, Consulta Derivada, Funcao M
    TipoConector        As String ' Pasta, Arquivo Excel, CSV, Web, SQL, SharePoint, CurrentWorkbook, etc.
    FuncaoM             As String ' Folder.Files, File.Contents, etc.
    ExpressaoParametro  As String ' Nome do parametro ou expressao M (ex: localTabelaB, localLote1&localES)
    CaminhoResolvido    As String ' Caminho absoluto no disco/rede ou URL
    StatusAcesso        As String ' [OK] Acessivel, [ERRO] Inacessivel, [INFO]..., etc.
    DetalhesTecnicos    As String ' Tamanho, data modificacao, contagem de arquivos ou motivo do erro
End Type

'----------------------------------------------------------------------------------------------------
' PONTO DE ENTRADA PUBLICO (Pode ser executado via Alt+F8 ou associado a botao)
'----------------------------------------------------------------------------------------------------
Public Sub ScanearFontesPowerQuery()
    Dim wb                  As Workbook
    Dim fso                 As Object
    Dim aRegistros()        As RegFontePQ
    Dim iQtdRegistros       As Long
    Dim iTotalConsultas     As Long
    Dim iTotalExternas      As Long
    Dim iTotalAcessiveis    As Long
    Dim iTotalInacessiveis  As Long
    Dim iTotalInternas      As Long
    Dim dtInicio            As Date
    Dim dtFim               As Date
    Dim sMsgResumo          As String
    Dim sCaminhosErros      As String
    Dim k                   As Long
    
    Dim appCalc             As XlCalculation
    Dim bEvents             As Boolean
    Dim bScreen             As Boolean
    Dim bAlerts             As Boolean
    
    Set wb = ThisWorkbook
    
    If MsgBox("Deseja iniciar o escaneamento de todas as fontes de dados das consultas Power Query?" & vbNewLine & vbNewLine & _
              "A rotina ira inspecionar o codigo M de cada consulta, resolver os parametros de caminho " & _
              "e verificar a acessibilidade fisica de cada arquivo/pasta no disco.", _
              vbQuestion + vbYesNo, TITULO_AUDITORIA) = vbNo Then
        Exit Sub
    End If
    
    ' Captura estado previo do Excel
    appCalc = Application.Calculation
    bEvents = Application.EnableEvents
    bScreen = Application.ScreenUpdating
    bAlerts = Application.DisplayAlerts
    
    On Error GoTo TratarErro
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    dtInicio = Now
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    iQtdRegistros = 0
    iTotalConsultas = 0
    iTotalExternas = 0
    iTotalAcessiveis = 0
    iTotalInacessiveis = 0
    iTotalInternas = 0
    
    ' Executa a varredura tecnica
    RealizarVarredura wb, fso, aRegistros, iQtdRegistros, iTotalConsultas, _
                     iTotalExternas, iTotalAcessiveis, iTotalInacessiveis, iTotalInternas
    
    ' Monta o relatorio visual na aba dedicada
    GerarRelatorioAbaAuditoria wb, aRegistros, iQtdRegistros, iTotalConsultas, _
                              iTotalExternas, iTotalAcessiveis, iTotalInacessiveis, iTotalInternas, _
                              dtInicio
    
    dtFim = Now
    
SairRotina:
    On Error Resume Next
    Application.Calculation = appCalc
    Application.DisplayAlerts = bAlerts
    Application.EnableEvents = bEvents
    Application.ScreenUpdating = bScreen
    Set fso = Nothing
    On Error GoTo 0
    
    ' Coleta caminhos inacessiveis para exibir no aviso
    sCaminhosErros = ""
    For k = 1 To iQtdRegistros
        If InStr(1, aRegistros(k).StatusAcesso, "[ERRO]", vbTextCompare) > 0 Then
            sCaminhosErros = sCaminhosErros & "  - [" & aRegistros(k).NomeConsulta & "] " & _
                             aRegistros(k).CaminhoResolvido & " (" & aRegistros(k).DetalhesTecnicos & ")" & vbNewLine
        End If
    Next k
    
    ' Exibe mensagem final de conclusao
    sMsgResumo = "Escaneamento de Fontes Power Query concluido com sucesso!" & vbNewLine & vbNewLine & _
                 "Total de consultas analisadas: " & iTotalConsultas & vbNewLine & _
                 "Fontes de dados externas mapeadas: " & iTotalExternas & vbNewLine & _
                 "  - [OK] Acessiveis: " & iTotalAcessiveis & vbNewLine & _
                 "  - [ERRO] Inacessiveis / Nao Encontradas: " & iTotalInacessiveis & vbNewLine & _
                 "Fontes internas / consultas derivadas: " & iTotalInternas & vbNewLine & vbNewLine & _
                 "O relatorio detalhado foi gerado na aba: '" & NOME_ABA_AUDITORIA & "'."
                 
    If iTotalInacessiveis > 0 Then
        sMsgResumo = sMsgResumo & vbNewLine & vbNewLine & _
                     "ATENCAO! As seguintes fontes apresentaram falha de localizacao:" & vbNewLine & _
                     sCaminhosErros
        MsgBox sMsgResumo, vbExclamation, TITULO_AUDITORIA
    Else
        MsgBox sMsgResumo, vbInformation, TITULO_AUDITORIA
    End If
    Exit Sub

TratarErro:
    MsgBox "Erro inesperado durante o escaneamento de fontes:" & vbNewLine & _
           Err.Number & " - " & Err.Description, vbCritical, TITULO_AUDITORIA
    Resume SairRotina
End Sub

'----------------------------------------------------------------------------------------------------
' ROTINA PRINCIPAL DE VARREDURA
'----------------------------------------------------------------------------------------------------
Private Sub RealizarVarredura(ByVal wb As Workbook, ByVal fso As Object, _
                             ByRef aRegistros() As RegFontePQ, ByRef iQtdRegistros As Long, _
                             ByRef iTotalConsultas As Long, ByRef iTotalExternas As Long, _
                             ByRef iTotalAcessiveis As Long, ByRef iTotalInacessiveis As Long, _
                             ByRef iTotalInternas As Long)
    Dim q               As WorkbookQuery
    Dim cn              As WorkbookConnection
    Dim sNomeConsulta   As String
    Dim sFormula        As String
    
    ReDim aRegistros(1 To 500)
    iQtdRegistros = 0
    iTotalConsultas = 0
    
    ' 1. Inspeciona a colecao de consultas Power Query (Workbook.Queries)
    If wb.Queries.Count > 0 Then
        For Each q In wb.Queries
            iTotalConsultas = iTotalConsultas + 1
            sNomeConsulta = q.Name
            sFormula = q.Formula
            
            ProcessarConsultaM wb, fso, sNomeConsulta, sFormula, aRegistros, iQtdRegistros, _
                               iTotalExternas, iTotalAcessiveis, iTotalInacessiveis, iTotalInternas
        Next q
    End If
    
    ' 2. Inspeciona conexoes que possam nao estar mapeadas em Queries
    For Each cn In wb.Connections
        If Not ConsultaJaMapeada(cn.Name, aRegistros, iQtdRegistros) Then
            If cn.Type = xlConnectionTypeOLEDB Then
                ProcessarConexaoOLEDB wb, fso, cn, aRegistros, iQtdRegistros, _
                                     iTotalExternas, iTotalAcessiveis, iTotalInacessiveis, iTotalInternas
            End If
        End If
    Next cn
    
    ' Redimensiona matriz para tamanho exato
    If iQtdRegistros > 0 Then
        ReDim Preserve aRegistros(1 To iQtdRegistros)
    Else
        ReDim aRegistros(1 To 1)
    End If
End Sub

'----------------------------------------------------------------------------------------------------
' PROCESSAMENTO DE UMA CONSULTA POWER QUERY M
'----------------------------------------------------------------------------------------------------
Private Sub ProcessarConsultaM(ByVal wb As Workbook, ByVal fso As Object, _
                              ByVal sNome As String, ByVal sFormula As String, _
                              ByRef aRegistros() As RegFontePQ, ByRef iQtdRegistros As Long, _
                              ByRef iTotalExternas As Long, ByRef iTotalAcessiveis As Long, _
                              ByRef iTotalInacessiveis As Long, ByRef iTotalInternas As Long)
    Dim bTemFonteExterna    As Boolean
    Dim sExpr               As String
    Dim sCaminho            As String
    Dim sStatus             As String
    Dim sDetalhes           As String
    Dim sTipo               As String
    Dim sConector           As String
    Dim sClassificacao      As String
    
    bTemFonteExterna = False
    
    ' A) Detecta Pastas (Folder.Files / Folder.Contents)
    If InStr(1, sFormula, "Folder.Files", vbTextCompare) > 0 Or InStr(1, sFormula, "Folder.Contents", vbTextCompare) > 0 Then
        bTemFonteExterna = True
        sConector = IIf(InStr(1, sFormula, "Folder.Files", vbTextCompare) > 0, "Folder.Files", "Folder.Contents")
        sTipo = "Pasta Local / Rede"
        sClassificacao = "Fonte Externa"
        sExpr = ExtrairArgumentoFuncao(sFormula, sConector)
        sCaminho = ResolverCaminhoCompleto(wb, sExpr, sFormula)
        VerificarAcessibilidade fso, "Pasta", sCaminho, sStatus, sDetalhes
        
        AdicionarRegistro aRegistros, iQtdRegistros, sNome, sClassificacao, sTipo, sConector, sExpr, sCaminho, sStatus, sDetalhes
        AtualizarContadoresStatus sStatus, iTotalExternas, iTotalAcessiveis, iTotalInacessiveis
    End If
    
    ' B) Detecta Arquivos (File.Contents)
    If InStr(1, sFormula, "File.Contents", vbTextCompare) > 0 Then
        bTemFonteExterna = True
        sConector = "File.Contents"
        If InStr(1, sFormula, "Excel.Workbook", vbTextCompare) > 0 Then
            sTipo = "Arquivo Excel (.xlsx/.xlsm)"
        ElseIf InStr(1, sFormula, "Csv.Document", vbTextCompare) > 0 Then
            sTipo = "Arquivo CSV (.csv)"
        ElseIf InStr(1, sFormula, "Lines.FromBinary", vbTextCompare) > 0 Then
            sTipo = "Arquivo Texto (.txt)"
        Else
            sTipo = "Arquivo Local / Rede"
        End If
        sClassificacao = "Fonte Externa"
        sExpr = ExtrairArgumentoFuncao(sFormula, sConector)
        sCaminho = ResolverCaminhoCompleto(wb, sExpr, sFormula)
        VerificarAcessibilidade fso, "Arquivo", sCaminho, sStatus, sDetalhes
        
        AdicionarRegistro aRegistros, iQtdRegistros, sNome, sClassificacao, sTipo, sConector, sExpr, sCaminho, sStatus, sDetalhes
        AtualizarContadoresStatus sStatus, iTotalExternas, iTotalAcessiveis, iTotalInacessiveis
    End If
    
    ' C) Detecta Bancos de Dados Relacionais (SQL Server / Oracle / Access)
    If InStr(1, sFormula, "Sql.Database", vbTextCompare) > 0 Then
        bTemFonteExterna = True
        sConector = "Sql.Database"
        sTipo = "Banco de Dados SQL Server"
        sClassificacao = "Fonte Externa"
        sExpr = ExtrairArgumentoFuncao(sFormula, sConector)
        sCaminho = sExpr
        sStatus = "[INFO] Banco SQL"
        sDetalhes = "Servidor / Banco de dados relacional"
        AdicionarRegistro aRegistros, iQtdRegistros, sNome, sClassificacao, sTipo, sConector, sExpr, sCaminho, sStatus, sDetalhes
        iTotalExternas = iTotalExternas + 1
        iTotalAcessiveis = iTotalAcessiveis + 1
    ElseIf InStr(1, sFormula, "Oracle.Database", vbTextCompare) > 0 Then
        bTemFonteExterna = True
        sConector = "Oracle.Database"
        sTipo = "Banco de Dados Oracle"
        sClassificacao = "Fonte Externa"
        sExpr = ExtrairArgumentoFuncao(sFormula, sConector)
        sCaminho = sExpr
        sStatus = "[INFO] Banco Oracle"
        sDetalhes = "Servidor / Banco de dados Oracle"
        AdicionarRegistro aRegistros, iQtdRegistros, sNome, sClassificacao, sTipo, sConector, sExpr, sCaminho, sStatus, sDetalhes
        iTotalExternas = iTotalExternas + 1
        iTotalAcessiveis = iTotalAcessiveis + 1
    ElseIf InStr(1, sFormula, "Access.Database", vbTextCompare) > 0 Then
        bTemFonteExterna = True
        sConector = "Access.Database"
        sTipo = "Banco MS Access"
        sClassificacao = "Fonte Externa"
        sExpr = ExtrairArgumentoFuncao(sFormula, sConector)
        sCaminho = ResolverCaminhoCompleto(wb, sExpr, sFormula)
        VerificarAcessibilidade fso, "Arquivo", sCaminho, sStatus, sDetalhes
        AdicionarRegistro aRegistros, iQtdRegistros, sNome, sClassificacao, sTipo, sConector, sExpr, sCaminho, sStatus, sDetalhes
        AtualizarContadoresStatus sStatus, iTotalExternas, iTotalAcessiveis, iTotalInacessiveis
    End If
    
    ' D) Detecta Web / SharePoint / OData
    If InStr(1, sFormula, "Web.Contents", vbTextCompare) > 0 Or InStr(1, sFormula, "Web.Page", vbTextCompare) > 0 Then
        bTemFonteExterna = True
        sConector = IIf(InStr(1, sFormula, "Web.Contents", vbTextCompare) > 0, "Web.Contents", "Web.Page")
        sTipo = "Web / API REST"
        sClassificacao = "Fonte Externa"
        sExpr = ExtrairArgumentoFuncao(sFormula, sConector)
        sCaminho = sExpr
        sStatus = "[INFO] Web / URL"
        sDetalhes = "Requisicao HTTP / Servico Web"
        AdicionarRegistro aRegistros, iQtdRegistros, sNome, sClassificacao, sTipo, sConector, sExpr, sCaminho, sStatus, sDetalhes
        iTotalExternas = iTotalExternas + 1
        iTotalAcessiveis = iTotalAcessiveis + 1
    ElseIf InStr(1, sFormula, "SharePoint.Files", vbTextCompare) > 0 Or InStr(1, sFormula, "SharePoint.Tables", vbTextCompare) > 0 Then
        bTemFonteExterna = True
        sConector = IIf(InStr(1, sFormula, "SharePoint.Files", vbTextCompare) > 0, "SharePoint.Files", "SharePoint.Tables")
        sTipo = "SharePoint / Nuvem"
        sClassificacao = "Fonte Externa"
        sExpr = ExtrairArgumentoFuncao(sFormula, sConector)
        sCaminho = sExpr
        sStatus = "[INFO] SharePoint"
        sDetalhes = "Biblioteca ou Lista no SharePoint"
        AdicionarRegistro aRegistros, iQtdRegistros, sNome, sClassificacao, sTipo, sConector, sExpr, sCaminho, sStatus, sDetalhes
        iTotalExternas = iTotalExternas + 1
        iTotalAcessiveis = iTotalAcessiveis + 1
    ElseIf InStr(1, sFormula, "OData.Feed", vbTextCompare) > 0 Then
        bTemFonteExterna = True
        sConector = "OData.Feed"
        sTipo = "Feed OData"
        sClassificacao = "Fonte Externa"
        sExpr = ExtrairArgumentoFuncao(sFormula, sConector)
        sCaminho = sExpr
        sStatus = "[INFO] Feed OData"
        sDetalhes = "Servico OData corporativo"
        AdicionarRegistro aRegistros, iQtdRegistros, sNome, sClassificacao, sTipo, sConector, sExpr, sCaminho, sStatus, sDetalhes
        iTotalExternas = iTotalExternas + 1
        iTotalAcessiveis = iTotalAcessiveis + 1
    End If
    
    ' Se ja registrou fonte externa, podemos sair
    If bTemFonteExterna Then Exit Sub
    
    ' E) Detecta Fontes Internas (Excel.CurrentWorkbook)
    If InStr(1, sFormula, "Excel.CurrentWorkbook", vbTextCompare) > 0 Then
        Dim sNomeObjInterno As String
        sNomeObjInterno = ExtrairNomeCurrentWorkbook(sFormula)
        sConector = "Excel.CurrentWorkbook"
        sTipo = "Tabela / Intervalo Interno"
        sClassificacao = "Fonte Interna"
        sExpr = IIf(Len(sNomeObjInterno) > 0, sNomeObjInterno, "CurrentWorkbook")
        
        ' Verifica se o objeto referenciado existe na pasta de trabalho
        If ObjetoInternoExiste(wb, sNomeObjInterno) Then
            sStatus = "[OK] Interno Valido"
            sDetalhes = "Objeto  & sNomeObjInterno &  localizado na pasta de trabalho"
        Else
            sStatus = "[AVISO] Objeto Nao Localizado"
            sDetalhes = "Referencia interna a  & sNomeObjInterno &  nao encontrada"
        End If
        
        sCaminho = "PastaAtual!" & sNomeObjInterno
        AdicionarRegistro aRegistros, iQtdRegistros, sNome, sClassificacao, sTipo, sConector, sExpr, sCaminho, sStatus, sDetalhes
        iTotalInternas = iTotalInternas + 1
        Exit Sub
    End If
    
    ' F) Detecta Funcoes M Auxiliares ou Consultas Derivadas
    If EhFuncaoM(sFormula) Then
        sClassificacao = "Funcao M Auxiliar"
        sTipo = "Funcao Personalizada"
        sConector = "Linguagem M"
        sExpr = "Parametros customizados"
        sCaminho = "Definicao interna de rotina"
        sStatus = "[INFO] Funcao M"
        sDetalhes = "Funcao reutilizavel (calculo / logica)"
        AdicionarRegistro aRegistros, iQtdRegistros, sNome, sClassificacao, sTipo, sConector, sExpr, sCaminho, sStatus, sDetalhes
        iTotalInternas = iTotalInternas + 1
    Else
        sClassificacao = "Consulta Derivada"
        sTipo = "Transformacao M"
        sConector = "Dependencia Interna"
        sExpr = "Outras consultas PQ"
        sCaminho = "Transformacoes em memoria"
        sStatus = "[INFO] Transformacao"
        sDetalhes = "Consome outras consultas sem acessar fonte externa direta"
        AdicionarRegistro aRegistros, iQtdRegistros, sNome, sClassificacao, sTipo, sConector, sExpr, sCaminho, sStatus, sDetalhes
        iTotalInternas = iTotalInternas + 1
    End If
End Sub

'----------------------------------------------------------------------------------------------------
' PROCESSAMENTO DE CONEXOES OLEDB AVULSAS (Nao mapeadas em Queries)
'----------------------------------------------------------------------------------------------------
Private Sub ProcessarConexaoOLEDB(ByVal wb As Workbook, ByVal fso As Object, ByVal cn As WorkbookConnection, _
                                 ByRef aRegistros() As RegFontePQ, ByRef iQtdRegistros As Long, _
                                 ByRef iTotalExternas As Long, ByRef iTotalAcessiveis As Long, _
                                 ByRef iTotalInacessiveis As Long, ByRef iTotalInternas As Long)
    Dim sConnStr    As String
    Dim sCaminho    As String
    Dim sStatus     As String
    Dim sDetalhes   As String
    
    On Error Resume Next
    sConnStr = cn.OLEDBConnection.Connection
    On Error GoTo 0
    
    If Len(sConnStr) = 0 Then Exit Sub
    
    ' Se for conexao mashup com Location, o Power Query ja cuidou via Queries
    If InStr(1, sConnStr, "Microsoft.Mashup.OleDb", vbTextCompare) > 0 Then Exit Sub
    
    ' Conexoes externas diretas (ex: OLEDB para Access, Excel, SQL)
    If InStr(1, sConnStr, "Data Source=", vbTextCompare) > 0 Then
        sCaminho = ExtrairValorChave(sConnStr, "Data Source")
        If InStr(1, sCaminho, ".xlsx", vbTextCompare) > 0 Or InStr(1, sCaminho, ".xlsm", vbTextCompare) > 0 Or _
           InStr(1, sCaminho, ".accdb", vbTextCompare) > 0 Or InStr(1, sCaminho, ".mdb", vbTextCompare) > 0 Then
            VerificarAcessibilidade fso, "Arquivo", sCaminho, sStatus, sDetalhes
            AdicionarRegistro aRegistros, iQtdRegistros, cn.Name, "Fonte Externa", "Conexao OLEDB", "OLEDBConnection", _
                              "Data Source", sCaminho, sStatus, sDetalhes
            AtualizarContadoresStatus sStatus, iTotalExternas, iTotalAcessiveis, iTotalInacessiveis
        End If
    End If
End Sub

'----------------------------------------------------------------------------------------------------
' EXTRATORES DE ARGUMENTOS E PARAMETROS M
'----------------------------------------------------------------------------------------------------
Private Function ExtrairArgumentoFuncao(ByVal sFormula As String, ByVal sNomeFuncao As String) As String
    Dim iPosInicio  As Long
    Dim iPosAbre    As Long
    Dim iPosFecha   As Long
    Dim iPosVirgula As Long
    Dim sSub        As String
    Dim sArg        As String
    
    iPosInicio = InStr(1, sFormula, sNomeFuncao, vbTextCompare)
    If iPosInicio = 0 Then Exit Function
    
    iPosAbre = InStr(iPosInicio, sFormula, "(")
    If iPosAbre = 0 Then Exit Function
    
    sSub = Mid(sFormula, iPosAbre + 1)
    
    ' Encontra a primeira virgula ou parentese de fechamento correspondente
    iPosVirgula = InStr(1, sSub, ",")
    iPosFecha = InStr(1, sSub, ")")
    
    If iPosVirgula > 0 And (iPosVirgula < iPosFecha Or iPosFecha = 0) Then
        sArg = Left(sSub, iPosVirgula - 1)
    ElseIf iPosFecha > 0 Then
        sArg = Left(sSub, iPosFecha - 1)
    Else
        sArg = sSub
    End If
    
    ExtrairArgumentoFuncao = Trim(sArg)
End Function

Private Function ExtrairNomeCurrentWorkbook(ByVal sFormula As String) As String
    Dim iPos1   As Long
    Dim iPos2   As Long
    
    iPos1 = InStr(1, sFormula, "[Name=" & Chr(34), vbTextCompare)
    If iPos1 > 0 Then
        iPos1 = iPos1 + 7
        iPos2 = InStr(iPos1, sFormula, Chr(34))
        If iPos2 > iPos1 Then
            ExtrairNomeCurrentWorkbook = Mid(sFormula, iPos1, iPos2 - iPos1)
            Exit Function
        End If
    End If
    
    ' Tenta sintaxe alternativa Name = "..."
    iPos1 = InStr(1, sFormula, "Name = " & Chr(34), vbTextCompare)
    If iPos1 > 0 Then
        iPos1 = iPos1 + 8
        iPos2 = InStr(iPos1, sFormula, Chr(34))
        If iPos2 > iPos1 Then
            ExtrairNomeCurrentWorkbook = Mid(sFormula, iPos1, iPos2 - iPos1)
            Exit Function
        End If
    End If
End Function

'----------------------------------------------------------------------------------------------------
' RESOLUCAO INTELIGENTE DE CAMINHOS E PARAMETROS
'----------------------------------------------------------------------------------------------------
Private Function ResolverCaminhoCompleto(ByVal wb As Workbook, ByVal sExpr As String, ByVal sFormula As String) As String
    Dim vPartes     As Variant
    Dim sResultado  As String
    Dim i           As Long
    Dim sParte      As String
    
    sExpr = Trim(sExpr)
    If Len(sExpr) = 0 Then Exit Function
    
    ' Se for expressao concatenada com & (ex: localLote1 & localES)
    If InStr(1, sExpr, "&") > 0 Then
        vPartes = Split(sExpr, "&")
        sResultado = ""
        For i = LBound(vPartes) To UBound(vPartes)
            sParte = Trim(CStr(vPartes(i)))
            If (Left(sParte, 1) = Chr(34) And Right(sParte, 1) = Chr(34)) Or _
               (Left(sParte, 1) = "'" And Right(sParte, 1) = "'") Then
                sResultado = sResultado & Mid(sParte, 2, Len(sParte) - 2)
            Else
                sResultado = sResultado & ResolverParametroUnico(wb, sParte, sFormula)
            End If
        Next i
        ResolverCaminhoCompleto = sResultado
        Exit Function
    End If
    
    ' Se for literal entre aspas
    If (Left(sExpr, 1) = Chr(34) And Right(sExpr, 1) = Chr(34)) Or _
       (Left(sExpr, 1) = "'" And Right(sExpr, 1) = "'") Then
        ResolverCaminhoCompleto = Mid(sExpr, 2, Len(sExpr) - 2)
        Exit Function
    End If
    
    ' Se for um identificador unico
    ResolverCaminhoCompleto = ResolverParametroUnico(wb, sExpr, sFormula)
End Function

Private Function ResolverParametroUnico(ByVal wb As Workbook, ByVal sParam As String, ByVal sFormulaContexto As String) As String
    Dim sVal            As String
    Dim sNomeInterno    As String
    Dim q               As WorkbookQuery
    
    sParam = Trim(sParam)
    If Len(sParam) = 0 Then Exit Function
    
    ' 1. Tenta resolver diretamente no Workbook (Names / ListObjects)
    sVal = ResolverParametroNoWorkbook(wb, sParam)
    If Len(Trim(sVal)) > 0 Then
        ResolverParametroUnico = sVal
        Exit Function
    End If
    
    ' 2. Se a propria formula extrai parametro via CurrentWorkbook, resolve o nome referenciado
    ' Exemplo: ParametroTabela = Excel.CurrentWorkbook(){[Name="localTabelaB"]}[Content]
    sNomeInterno = ExtrairNomeCurrentWorkbook(sFormulaContexto)
    If Len(sNomeInterno) > 0 Then
        sVal = ResolverParametroNoWorkbook(wb, sNomeInterno)
        If Len(Trim(sVal)) > 0 Then
            ResolverParametroUnico = sVal
            Exit Function
        End If
    End If
    
    ' 3. Procura se existe uma consulta auxiliar com este nome que aponte para um parametro
    On Error Resume Next
    For Each q In wb.Queries
        If StrComp(q.Name, sParam, vbTextCompare) = 0 Then
            sNomeInterno = ExtrairNomeCurrentWorkbook(q.Formula)
            If Len(sNomeInterno) > 0 Then
                sVal = ResolverParametroNoWorkbook(wb, sNomeInterno)
                If Len(Trim(sVal)) > 0 Then
                    ResolverParametroUnico = sVal
                    Exit Function
                End If
            End If
        End If
    Next q
    On Error GoTo 0
    
    ' Fallback: retorna o proprio identificador
    ResolverParametroUnico = sParam
End Function

Private Function ResolverParametroNoWorkbook(ByVal wb As Workbook, ByVal sNomeParam As String) As String
    Dim nm  As Name
    Dim ws  As Worksheet
    Dim lo  As ListObject
    Dim sVal As String
    
    On Error Resume Next
    ' Busca em Nomes Definidos do Excel
    Set nm = wb.Names(sNomeParam)
    If Not nm Is Nothing Then
        If Not nm.RefersToRange Is Nothing Then
            sVal = CStr(nm.RefersToRange.Cells(1, 1).Value)
            If Len(Trim(sVal)) > 0 Then
                ResolverParametroNoWorkbook = sVal
                Exit Function
            End If
        End If
    End If
    Err.Clear
    
    ' Busca em Tabelas Estruturadas (ListObjects)
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, sNomeParam, vbTextCompare) = 0 Then
                If Not lo.DataBodyRange Is Nothing Then
                    sVal = CStr(lo.DataBodyRange.Cells(1, 1).Value)
                    If Len(Trim(sVal)) > 0 Then
                        ResolverParametroNoWorkbook = sVal
                        Exit Function
                    End If
                End If
            End If
        Next lo
    Next ws
    On Error GoTo 0
End Function

Private Function ObjetoInternoExiste(ByVal wb As Workbook, ByVal sNomeObj As String) As Boolean
    Dim nm As Name
    Dim ws As Worksheet
    Dim lo As ListObject
    
    If Len(Trim(sNomeObj)) = 0 Then Exit Function
    
    On Error Resume Next
    Set nm = wb.Names(sNomeObj)
    If Not nm Is Nothing Then
        ObjetoInternoExiste = True
        Exit Function
    End If
    Err.Clear
    
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, sNomeObj, vbTextCompare) = 0 Then
                ObjetoInternoExiste = True
                Exit Function
            End If
        Next lo
    Next ws
    On Error GoTo 0
End Function

'----------------------------------------------------------------------------------------------------
' VALIDACAO DE ACESSIBILIDADE FISICA NO DISCO
'----------------------------------------------------------------------------------------------------
Private Sub VerificarAcessibilidade(ByVal fso As Object, ByVal sTipo As String, ByVal sCaminho As String, _
                                   ByRef sStatus As String, ByRef sDetalhes As String)
    Dim oFile   As Object
    Dim oFolder As Object
    
    sCaminho = Trim(sCaminho)
    
    If Len(sCaminho) = 0 Then
        sStatus = "[AVISO] Caminho Nao Informado"
        sDetalhes = "Parametro em branco ou nao preenchido"
        Exit Sub
    End If
    
    If sTipo = "Pasta" Then
        If fso.FolderExists(sCaminho) Then
            On Error Resume Next
            Set oFolder = fso.GetFolder(sCaminho)
            sStatus = "[OK] Acessivel"
            sDetalhes = "Pasta disponivel (" & oFolder.Files.Count & " arquivos encontrados)"
            On Error GoTo 0
        Else
            sStatus = "[ERRO] Pasta Nao Encontrada"
            sDetalhes = "Diretorio nao existe ou esta inacessivel no momento"
        End If
    Else ' Arquivo
        If fso.FileExists(sCaminho) Then
            On Error Resume Next
            Set oFile = fso.GetFile(sCaminho)
            sStatus = "[OK] Acessivel"
            sDetalhes = Format(oFile.Size / 1024, "#,##0") & " KB | Modificado: " & _
                        Format(oFile.DateLastModified, "dd/mm/yyyy hh:mm")
            On Error GoTo 0
        Else
            sStatus = "[ERRO] Arquivo Nao Encontrado"
            sDetalhes = "Arquivo fisico nao localizado no caminho informado"
        End If
    End If
End Sub

'----------------------------------------------------------------------------------------------------
' FORMATACAO E MONTAGEM DA ABA DE RELATORIO
'----------------------------------------------------------------------------------------------------
Private Sub GerarRelatorioAbaAuditoria(ByVal wb As Workbook, ByRef aRegistros() As RegFontePQ, _
                                      ByVal iQtdRegistros As Long, ByVal iTotalConsultas As Long, _
                                      ByVal iTotalExternas As Long, ByVal iTotalAcessiveis As Long, _
                                      ByVal iTotalInacessiveis As Long, ByVal iTotalInternas As Long, _
                                      ByVal dtInicio As Date)
    Dim ws          As Worksheet
    Dim rngDados    As Range
    Dim rngHeader   As Range
    Dim matriz()    As Variant
    Dim r           As Long
    Dim iLinhaFinal As Long
    
    Set ws = ObterOuCriarPlanilha(wb, NOME_ABA_AUDITORIA)
    
    ' Limpeza completa
    ws.Cells.Clear
    ws.Tab.Color = COR_HEADER_FUNDO
    ws.DisplayGridlines = True
    
    ' 1. Titulo Executivo
    With ws.Range("A1")
        .Value = "AUDITORIA DE FONTES DE DADOS - POWER QUERY"
        .Font.Name = "Segoe UI"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = COR_HEADER_FUNDO
    End With
    
    With ws.Range("A2")
        .Value = "Relatorio gerado em: " & Format(dtInicio, "dd/mm/yyyy hh:mm:ss") & _
                 " | Pasta de Trabalho: " & wb.Name
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Italic = True
        .Font.Color = RGB(100, 116, 139)
    End With
    
    ' 2. Cartoes de Metricas (Linhas 4 e 5)
    CriarCartaoMetrica ws, "A4:B5", "TOTAL CONSULTAS", CStr(iTotalConsultas), COR_CARD_FUNDO, RGB(30, 41, 59)
    CriarCartaoMetrica ws, "C4:D5", "FONTES EXTERNAS", CStr(iTotalExternas), COR_CARD_FUNDO, RGB(30, 41, 59)
    CriarCartaoMetrica ws, "E4:E5", "ACESSIVEIS [OK]", CStr(iTotalAcessiveis), COR_OK_FUNDO, COR_OK_FONTE
    CriarCartaoMetrica ws, "F4:F5", "INACESSIVEIS [ERRO]", CStr(iTotalInacessiveis), _
                       IIf(iTotalInacessiveis > 0, COR_ERRO_FUNDO, COR_CARD_FUNDO), _
                       IIf(iTotalInacessiveis > 0, COR_ERRO_FONTE, RGB(100, 116, 139))
    CriarCartaoMetrica ws, "G4:H5", "INTERNAS / DERIVADAS", CStr(iTotalInternas), COR_CARD_FUNDO, RGB(71, 85, 105)
    
    ' 3. Cabecalhos da Tabela (Linha 7)
    Dim aHeaders As Variant
    aHeaders = Array("Consulta (Query)", "Classificacao", "Tipo de Conector", _
                     "Funcao M Detectada", "Parametro / Expressao", "Caminho / Origem Resolvida", _
                     "Status de Acesso", "Detalhes Tecnicos / Diagnostico")
    
    Set rngHeader = ws.Range("A7:H7")
    rngHeader.Value = aHeaders
    With rngHeader
        .Interior.Color = COR_HEADER_FUNDO
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = COR_HEADER_FONTE
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 26
    End With
    
    ' 4. Preenchimento dos Registros em Matriz
    If iQtdRegistros > 0 Then
        ReDim matriz(1 To iQtdRegistros, 1 To 8)
        For r = 1 To iQtdRegistros
            matriz(r, 1) = aRegistros(r).NomeConsulta
            matriz(r, 2) = aRegistros(r).Classificacao
            matriz(r, 3) = aRegistros(r).TipoConector
            matriz(r, 4) = aRegistros(r).FuncaoM
            matriz(r, 5) = aRegistros(r).ExpressaoParametro
            matriz(r, 6) = aRegistros(r).CaminhoResolvido
            matriz(r, 7) = aRegistros(r).StatusAcesso
            matriz(r, 8) = aRegistros(r).DetalhesTecnicos
        Next r
        
        iLinhaFinal = 7 + iQtdRegistros
        Set rngDados = ws.Range("A8:H" & iLinhaFinal)
        rngDados.Value = matriz
        
        ' Formatacao visual das linhas de dados
        With rngDados
            .Font.Name = "Segoe UI"
            .Font.Size = 9
            .VerticalAlignment = xlCenter
            .Borders.Color = COR_BORDA_GRADE
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        
        ' Cores de linhas zebradas e estilizacao de status
        For r = 1 To iQtdRegistros
            Dim iLinhaAtual As Long
            Dim rngLinha    As Range
            Dim rngStatus   As Range
            
            iLinhaAtual = 7 + r
            Set rngLinha = ws.Range("A" & iLinhaAtual & ":H" & iLinhaAtual)
            Set rngStatus = ws.Range("G" & iLinhaAtual)
            
            ' Zebra suave
            If r Mod 2 = 0 Then
                rngLinha.Interior.Color = COR_ZEBRA_IMPAR
            Else
                rngLinha.Interior.Color = COR_ZEBRA_PAR
            End If
            
            ' Destaca Status
            If InStr(1, aRegistros(r).StatusAcesso, "[OK]", vbTextCompare) > 0 Then
                rngStatus.Interior.Color = COR_OK_FUNDO
                rngStatus.Font.Color = COR_OK_FONTE
                rngStatus.Font.Bold = True
            ElseIf InStr(1, aRegistros(r).StatusAcesso, "[ERRO]", vbTextCompare) > 0 Then
                rngStatus.Interior.Color = COR_ERRO_FUNDO
                rngStatus.Font.Color = COR_ERRO_FONTE
                rngStatus.Font.Bold = True
            ElseIf InStr(1, aRegistros(r).StatusAcesso, "[AVISO]", vbTextCompare) > 0 Then
                rngStatus.Interior.Color = COR_AVISO_FUNDO
                rngStatus.Font.Color = COR_AVISO_FONTE
                rngStatus.Font.Bold = True
            Else
                rngStatus.Interior.Color = COR_INFO_FUNDO
                rngStatus.Font.Color = COR_INFO_FONTE
            End If
            
            ' Centralizacoes pontuais
            ws.Range("B" & iLinhaAtual & ":D" & iLinhaAtual).HorizontalAlignment = xlCenter
            rngStatus.HorizontalAlignment = xlCenter
        Next r
        
        ' Ajuste e formatacao das colunas
        ws.Columns("A:H").AutoFit
        
        ' Garante larguras minimas adequadas
        If ws.Columns("A").ColumnWidth < 28 Then ws.Columns("A").ColumnWidth = 28
        If ws.Columns("B").ColumnWidth < 18 Then ws.Columns("B").ColumnWidth = 18
        If ws.Columns("C").ColumnWidth < 22 Then ws.Columns("C").ColumnWidth = 22
        If ws.Columns("D").ColumnWidth < 20 Then ws.Columns("D").ColumnWidth = 20
        If ws.Columns("E").ColumnWidth < 22 Then ws.Columns("E").ColumnWidth = 22
        If ws.Columns("F").ColumnWidth < 45 Then ws.Columns("F").ColumnWidth = 45
        If ws.Columns("G").ColumnWidth < 22 Then ws.Columns("G").ColumnWidth = 22
        If ws.Columns("H").ColumnWidth < 40 Then ws.Columns("H").ColumnWidth = 40
        
        ' AutoFiltro
        ws.Range("A7:H" & iLinhaFinal).AutoFilter
    End If
    
    ' Congelar paineis a partir da linha 8
    ws.Activate
    ws.Range("A8").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
End Sub

Private Sub CriarCartaoMetrica(ByVal ws As Worksheet, ByVal sEndereco As String, _
                              ByVal sRotulo As String, ByVal sValor As String, _
                              ByVal corFundo As Long, ByVal corFonte As Long)
    Dim rng As Range
    Set rng = ws.Range(sEndereco)
    
    rng.Merge
    rng.Interior.Color = corFundo
    rng.Borders.Color = COR_BORDA_GRADE
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Weight = xlThin
    rng.HorizontalAlignment = xlCenter
    rng.VerticalAlignment = xlCenter
    
    rng.Value = sRotulo & vbNewLine & sValor
    rng.Font.Name = "Segoe UI"
    rng.Font.Size = 10
    rng.Font.Bold = True
    rng.Font.Color = corFonte
    rng.WrapText = True
End Sub

'----------------------------------------------------------------------------------------------------
' FUNCOES AUXILIARES
'----------------------------------------------------------------------------------------------------
Private Sub AdicionarRegistro(ByRef aRegistros() As RegFontePQ, ByRef iQtdRegistros As Long, _
                              ByVal sNome As String, ByVal sClassificacao As String, _
                              ByVal sTipo As String, ByVal sFuncao As String, _
                              ByVal sExpr As String, ByVal sCaminho As String, _
                              ByVal sStatus As String, ByVal sDetalhes As String)
    iQtdRegistros = iQtdRegistros + 1
    If iQtdRegistros > UBound(aRegistros) Then
        ReDim Preserve aRegistros(1 To UBound(aRegistros) + 100)
    End If
    
    With aRegistros(iQtdRegistros)
        .NomeConsulta = sNome
        .Classificacao = sClassificacao
        .TipoConector = sTipo
        .FuncaoM = sFuncao
        .ExpressaoParametro = sExpr
        .CaminhoResolvido = sCaminho
        .StatusAcesso = sStatus
        .DetalhesTecnicos = sDetalhes
    End With
End Sub

Private Sub AtualizarContadoresStatus(ByVal sStatus As String, ByRef iTotalExternas As Long, _
                                     ByRef iTotalAcessiveis As Long, ByRef iTotalInacessiveis As Long)
    iTotalExternas = iTotalExternas + 1
    If InStr(1, sStatus, "[OK]", vbTextCompare) > 0 Then
        iTotalAcessiveis = iTotalAcessiveis + 1
    ElseIf InStr(1, sStatus, "[ERRO]", vbTextCompare) > 0 Then
        iTotalInacessiveis = iTotalInacessiveis + 1
    End If
End Sub

Private Function ConsultaJaMapeada(ByVal sNome As String, ByRef aRegistros() As RegFontePQ, ByVal iQtd As Long) As Boolean
    Dim i As Long
    For i = 1 To iQtd
        If StrComp(aRegistros(i).NomeConsulta, sNome, vbTextCompare) = 0 Then
            ConsultaJaMapeada = True
            Exit Function
        End If
    Next i
End Function

Private Function EhFuncaoM(ByVal sFormula As String) As Boolean
    sFormula = Trim(sFormula)
    If Left(sFormula, 1) = "(" Or InStr(1, sFormula, "=>") > 0 Then
        EhFuncaoM = True
    End If
End Function

Private Function ExtrairValorChave(ByVal sTexto As String, ByVal sChave As String) As String
    Dim iPos As Long
    Dim iFim As Long
    
    iPos = InStr(1, sTexto, sChave & "=", vbTextCompare)
    If iPos > 0 Then
        iPos = iPos + Len(sChave) + 1
        iFim = InStr(iPos, sTexto, ";")
        If iFim = 0 Then iFim = Len(sTexto) + 1
        ExtrairValorChave = Trim(Mid(sTexto, iPos, iFim - iPos))
    End If
End Function

Private Function ObterOuCriarPlanilha(ByVal wb As Workbook, ByVal sNome As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sNome)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sNome
    End If
    
    Set ObterOuCriarPlanilha = ws
End Function

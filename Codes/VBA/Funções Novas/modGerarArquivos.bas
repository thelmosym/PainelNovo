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
'   - Cópia de fórmulas ativas na aba MC, redirecionando referências para as abas locais (ARM, DADOS, etc.)
'     e quebrando apenas vínculos externos de tabelas de apoio não exportadas (PPU / FDM).
'   - Gravação de Hyperlinks na aba Painel (F14:F17) com sinalização de cores (Amarelo, Verde Claro, Vermelho).
'====================================================================================================

Private Const MSG_TITULO As String = "Geração de Arquivos por Contrato"
Private Const NOME_PLANILHA_LINKS As String = "Painel"
Private Const COLUNA_CELULA_LINK As String = "F"
Private Const LINHA_INICIAL_LINK As Long = 14
Private Const NOME_SUBPASTA_DESTINO As String = "MEDIÇÃO"

' Paleta de cores corporativa para status dos links na aba Painel (RGB)
Private Const COR_STATUS_AMARELO     As Long = 65535        ' RGB(255, 255, 0)   - Amarelo (processando)
Private Const COR_STATUS_VERDE_CLARO As Long = 5296274      ' RGB(146, 208, 80)  - Verde suave (mesmo padrão das demais rotinas)
Private Const COR_STATUS_VERMELHO    As Long = 255          ' RGB(255, 0, 0)     - Vermelho (erro / falha)

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
    
    '-- Inicializa células F14:F17 em amarelo (indica geração em andamento)
    Dim iInicial As Long
    For iInicial = LBound(vContratos) To UBound(vContratos)
        Dim rngIni As Range
        Set rngIni = wsLinks.Range(COLUNA_CELULA_LINK & ObterLinhaCelulaLink(iInicial))
        rngIni.Hyperlinks.Delete
        rngIni.Value = "Gerando: " & CStr(vContratos(iInicial)) & "..."
        If rngIni.MergeCells Then
            rngIni.MergeArea.Interior.Color = COR_STATUS_AMARELO
            rngIni.MergeArea.Font.Color = vbBlack
            rngIni.MergeArea.Font.Bold = False
        Else
            rngIni.Interior.Color = COR_STATUS_AMARELO
            rngIni.Font.Color = vbBlack
            rngIni.Font.Bold = False
        End If
    Next iInicial
    DoEvents
    
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
    
    '-- Garante que qualquer contrato que não tenha sido gerado com sucesso termine em vermelho
    On Error Resume Next
    If Not wsLinks Is Nothing And iArquivosOK < (UBound(vContratos) + 1) Then
        Dim iFim As Long
        For iFim = LBound(vContratos) To UBound(vContratos)
            Dim rngFim As Range
            Set rngFim = wsLinks.Range(COLUNA_CELULA_LINK & ObterLinhaCelulaLink(iFim))
            If rngFim.Hyperlinks.Count = 0 Then
                If rngFim.MergeCells Then
                    rngFim.MergeArea.Interior.Color = COR_STATUS_VERMELHO
                    rngFim.MergeArea.Font.Color = vbWhite
                    rngFim.MergeArea.Font.Bold = True
                Else
                    rngFim.Interior.Color = COR_STATUS_VERMELHO
                    rngFim.Font.Color = vbWhite
                    rngFim.Font.Bold = True
                End If
            End If
        Next iFim
    End If
    On Error GoTo 0
    
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
    On Error Resume Next
    If Not wsLinks Is Nothing Then
        Dim iErrIdx As Long
        For iErrIdx = LBound(vContratos) To UBound(vContratos)
            Dim rngErr As Range
            Set rngErr = wsLinks.Range(COLUNA_CELULA_LINK & ObterLinhaCelulaLink(iErrIdx))
            If rngErr.Hyperlinks.Count = 0 Then
                rngErr.Hyperlinks.Delete
                If rngErr.MergeCells Then
                    rngErr.MergeArea.Interior.Color = COR_STATUS_VERMELHO
                    rngErr.MergeArea.Font.Color = vbWhite
                    rngErr.MergeArea.Font.Bold = True
                Else
                    rngErr.Interior.Color = COR_STATUS_VERMELHO
                    rngErr.Font.Color = vbWhite
                    rngErr.Font.Bold = True
                End If
                If Len(rngErr.Value) = 0 Or InStr(rngErr.Value, "Gerando") > 0 Then
                    rngErr.Value = "Erro ao gerar: " & CStr(vContratos(iErrIdx))
                End If
            End If
        Next iErrIdx
    End If
    On Error GoTo 0
    
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
    Dim wbDestino        As Workbook
    Dim wsOrigem         As Worksheet
    Dim wsDestino        As Worksheet
    Dim iTipo            As Long
    Dim sTipo            As String
    Dim sNomeAbaOrig     As String
    Dim iCopiadas        As Long
    Dim sFaltantes       As String
    Dim sNomeArquivo     As String
    Dim sCaminhoCompleto As String
    Dim sSufixoComHifen  As String
    Dim rngArea          As Range
    
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
    
    ' Ajusta as fórmulas da aba MC para referenciar as abas locais (ARM, DADOS, etc.) e quebra vínculos externos
    AjustarFormulasAbaMC wbDestino, wbOrigem, sContrato
    
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
    
    ' Se a planilha de origem possuir ListObject (Tabela estruturada), recria a tabela no destino
    On Error Resume Next
    If rngOrigem.Worksheet.ListObjects.Count > 0 Then
        Dim tblOrig As ListObject
        Dim tblDest As ListObject
        Dim rngTabelaDest As Range
        
        Set tblOrig = rngOrigem.Worksheet.ListObjects(1)
        If rngOrigem.Rows.Count >= 2 Then
            Set rngTabelaDest = wsDestino.Range("A1").Resize(rngOrigem.Rows.Count, rngOrigem.Columns.Count)
            Set tblDest = wsDestino.ListObjects.Add(xlSrcRange, rngTabelaDest, , xlYes)
            If Not tblDest Is Nothing Then
                tblDest.Name = tblOrig.Name
                tblDest.TableStyle = tblOrig.TableStyle
            End If
        End If
    End If
    On Error GoTo 0
    Exit Sub

TratarErroLocal:
    Application.CutCopyMode = False
End Sub

Private Sub CopiarComFormatacaoCompleta(ByVal rngOrigem As Range, ByVal wsDestino As Worksheet)
    If rngOrigem Is Nothing Then Exit Sub
    On Error GoTo TratarErroLocal
    
    ' Copia conteúdo preservando integralmente formatos e fórmulas (sem congelar em valores)
    rngOrigem.Copy
    wsDestino.Range("A1").PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False
    
    ReplicarLargurasColunas rngOrigem, wsDestino
    ReplicarAlturasLinhas rngOrigem, wsDestino
    Exit Sub

TratarErroLocal:
    Application.CutCopyMode = False
    On Error Resume Next
    rngOrigem.Copy
    wsDestino.Range("A1").PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False
    On Error GoTo 0
End Sub

'----------------------------------------------------------------------------------------------------
' AJUSTE DE FÓRMULAS DA ABA MC PARA REFERENCIAR AS ABAS DO NOVO ARQUIVO
'----------------------------------------------------------------------------------------------------
Private Sub AjustarFormulasAbaMC(ByVal wbDestino As Workbook, ByVal wbOrigem As Workbook, ByVal sContrato As String)
    Dim wsMC                As Worksheet
    Dim rngFormulas         As Range
    Dim cel                 As Range
    Dim sFormula            As String
    Dim sFormulaOriginal    As String
    Dim sNomeWbOrigem       As String
    Dim sCaminhoWbOrigem    As String
    Dim sAbaOrigARM         As String
    Dim sAbaOrigDADOS       As String
    Dim sAbaOrigFRETE       As String
    Dim sAbaOrigMC          As String
    Dim vLinks              As Variant
    Dim iLink               As Long
    
    On Error Resume Next
    Set wsMC = wbDestino.Worksheets("MC")
    On Error GoTo 0
    If wsMC Is Nothing Then Exit Sub
    
    ' Força o cálculo antes de ajustar para garantir que todos os valores estejam computados
    On Error Resume Next
    wbDestino.Calculate
    On Error GoTo 0
    
    sNomeWbOrigem = wbOrigem.Name
    sCaminhoWbOrigem = wbOrigem.FullName
    sAbaOrigARM = "ARM_" & sContrato
    sAbaOrigDADOS = "DADOS_" & sContrato
    sAbaOrigFRETE = "FRETE_" & sContrato
    sAbaOrigMC = "MC_" & sContrato
    
    On Error Resume Next
    Set rngFormulas = wsMC.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    
    If Not rngFormulas Is Nothing Then
        For Each cel In rngFormulas
            sFormula = cel.Formula
            sFormulaOriginal = sFormula
            
            ' 1. Remove referências com caminho completo da pasta de trabalho de origem
            If InStr(1, sFormula, sCaminhoWbOrigem, vbTextCompare) > 0 Then
                sFormula = Replace(sFormula, "'[" & sCaminhoWbOrigem & "]'", "")
                sFormula = Replace(sFormula, "[" & sCaminhoWbOrigem & "]", "")
                sFormula = Replace(sFormula, "'" & sCaminhoWbOrigem & "'!", "")
                sFormula = Replace(sFormula, sCaminhoWbOrigem & "!", "")
            End If
            
            ' 2. Remove referências pelo nome do arquivo de origem
            If InStr(1, sFormula, sNomeWbOrigem, vbTextCompare) > 0 Then
                sFormula = Replace(sFormula, "'[" & sNomeWbOrigem & "]'", "")
                sFormula = Replace(sFormula, "[" & sNomeWbOrigem & "]", "")
                sFormula = Replace(sFormula, "'" & sNomeWbOrigem & "'!", "")
                sFormula = Replace(sFormula, sNomeWbOrigem & "!", "")
            End If
            
            ' 3. Redireciona referências de abas específicas do contrato para as abas locais no novo arquivo
            If InStr(1, sFormula, sAbaOrigARM, vbTextCompare) > 0 Then
                sFormula = Replace(sFormula, "'" & sAbaOrigARM & "'!", "ARM!")
                sFormula = Replace(sFormula, sAbaOrigARM & "!", "ARM!")
            End If
            
            If InStr(1, sFormula, sAbaOrigDADOS, vbTextCompare) > 0 Then
                sFormula = Replace(sFormula, "'" & sAbaOrigDADOS & "'!", "DADOS!")
                sFormula = Replace(sFormula, sAbaOrigDADOS & "!", "DADOS!")
            End If
            
            If InStr(1, sFormula, sAbaOrigFRETE, vbTextCompare) > 0 Then
                sFormula = Replace(sFormula, "'" & sAbaOrigFRETE & "'!", "FRETE!")
                sFormula = Replace(sFormula, sAbaOrigFRETE & "!", "FRETE!")
            End If
            
            If InStr(1, sFormula, sAbaOrigMC, vbTextCompare) > 0 Then
                sFormula = Replace(sFormula, "'" & sAbaOrigMC & "'!", "MC!")
                sFormula = Replace(sFormula, sAbaOrigMC & "!", "MC!")
            End If
            
            If sFormula <> sFormulaOriginal Then
                On Error Resume Next
                cel.Formula = sFormula
                On Error GoTo 0
            End If
        Next cel
    End If
    
    ' 4. Quebra vínculos externos remanescentes (tabelas de apoio como PPU e FDM que não existem no destino),
    ' convertendo-os em seus valores estáticos calculados para garantir um caderno 100% independente
    On Error Resume Next
    vLinks = wbDestino.LinkSources(xlExcelLinks)
    If Not IsEmpty(vLinks) Then
        For iLink = LBound(vLinks) To UBound(vLinks)
            wbDestino.BreakLink Name:=vLinks(iLink), Type:=xlLinkTypeExcelLinks
        Next iLink
    End If
    
    ' Recalcula a pasta destino para assegurar coerência completa de todas as fórmulas
    wbDestino.Calculate
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
    
    ' Link gerado sem erros: muda para verde claro
    If rngCelula.MergeCells Then
        rngCelula.MergeArea.Interior.Color = COR_STATUS_VERDE_CLARO
        rngCelula.MergeArea.Font.Bold = False
    Else
        rngCelula.Interior.Color = COR_STATUS_VERDE_CLARO
        rngCelula.Font.Bold = False
    End If
    
    ws.Hyperlinks.Add Anchor:=rngCelula, Address:=sCaminhoCompleto, TextToDisplay:=sNomeArq
    Exit Sub

TratarErroLocal:
    On Error Resume Next
    rngCelula.Value = sCaminhoCompleto
    If rngCelula.MergeCells Then
        rngCelula.MergeArea.Interior.Color = COR_STATUS_VERMELHO
        rngCelula.MergeArea.Font.Color = vbWhite
        rngCelula.MergeArea.Font.Bold = True
    Else
        rngCelula.Interior.Color = COR_STATUS_VERMELHO
        rngCelula.Font.Color = vbWhite
        rngCelula.Font.Bold = True
    End If
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
    
    ' Caso de erro: deixa na cor vermelha
    If rngCelula.MergeCells Then
        rngCelula.MergeArea.Interior.Color = COR_STATUS_VERMELHO
        rngCelula.MergeArea.Font.Color = vbWhite
        rngCelula.MergeArea.Font.Bold = True
    Else
        rngCelula.Interior.Color = COR_STATUS_VERMELHO
        rngCelula.Font.Color = vbWhite
        rngCelula.Font.Bold = True
    End If
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

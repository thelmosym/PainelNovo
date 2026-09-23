Option Explicit

'-----------------------------------------------------------------------
' MÓDULO: modGerarArquivosPorContrato
'---------------------------------------------------------------------
' Gera 4 arquivos finais (um por contrato: PA-LT2, IRON-LT1-RJ,
' IRON-LT1-SP, IRON-LT1-ES), cada um com 4 abas fixas (MC, ARM, DADOS,
' FRETE), copiando os dados das respectivas ABAS de origem da
' planilha principal, nomeadas como "{TIPO}_{CONTRATO}".
'
' Nome do arquivo final: "{NumeroContrato}-PLA-MemoriaPetrobras-
' {Contrato}[-{Sufixo}]_{DDMMAAAA}_{HHMMSS}.xlsx"
' Ex. PA-LT2: 4600687006-PLA-MemoriaPetrobras-PA-LT2-BA_18092026_133421.xlsx
' Ex. IRON-LT1-RJ: 4600686987-PLA-MemoriaPetrobras-IRON-LT1-RJ_18092026_133425.xlsx
'
' Local de salvamento: subpasta "MEDIÇÃO" dentro da mesma pasta onde
' este arquivo (Painel) está salvo. A subpasta é criada
' automaticamente caso não exista.
'
' Ao final, grava um HYPERLINK para cada arquivo gerado nas células
' K7 (PA-LT2), K8 (IRON-LT1-RJ), K9 (IRON-LT1-SP), K10 (IRON-LT1-ES)
' da aba "Painel".
'
' Associe esta macro (GerarArquivosPorContrato) a um botão via:
' Inserir > Formas/Botão > Botão direito > Atribuir Macro.
'-----------------------------------------------------------------------

Private Const MSG_TITULO As String = "Geração de Arquivos por Contrato"

' Aba onde os links dos arquivos gerados serão gravados
Private Const NOME_PLANILHA_LINKS As String = "Painel"

' Coluna e linha inicial onde os links serão gravados (K7 a K10)
Private Const COLUNA_CELULA_LINK As String = "K"
Private Const LINHA_INICIAL_LINK As Integer = 7

' Nome da subpasta onde os arquivos finais serão salvos (dentro da
' mesma pasta em que este arquivo Painel está salvo)
Private Const NOME_SUBPASTA_DESTINO As String = "MEDIÇÃO"

' GUID do rótulo de confidencialidade "Pública" do tenant da Petrobras.
' Enquanto esta constante estiver VAZIA (""), a aplicação do rótulo é
' pulada (ver AplicarPerfilPublico).
Private Const PUBLICO_LABEL_ID As String = "140b9f7d-8e3a-482f-9702-4b7ffc40985a"

Private Const MSO_ASSIGNMENT_METHOD_STANDARD As Long = 1
Private Const MSO_ASSIGNMENT_METHOD_PRIVILEGED As Long = 2

' Tipos de aba de origem — também serão os nomes das abas no arquivo
' final, NESTA ORDEM
Private Function ListaTipos() As Variant
    ListaTipos = Array("MC", "ARM", "DADOS", "FRETE")
End Function

' Contratos — cada um gera um arquivo final separado (nome das ABAS
' de origem, SEM sufixo de estado).
' A ORDEM DESTE ARRAY DEFINE O MAPEAMENTO DE LINHAS (K7 a K10) e
' TAMBÉM o mapeamento dos arrays ListaNumerosContrato e
' ListaSufixosEstado abaixo — todos DEVEM manter a MESMA ORDEM.
Private Function ListaContratos() As Variant
    ListaContratos = Array("PA-LT2", "IRON-LT1-RJ", "IRON-LT1-SP", "IRON-LT1-ES")
End Function

' Número do contrato SAP de cada contrato, NA MESMA ORDEM de
' ListaContratos() acima.
' Os três lotes IRON (RJ/SP/ES) compartilham o mesmo número de
' contrato-mãe (4600686987).
' Ordem: PA-LT2 | IRON-LT1-RJ | IRON-LT1-SP | IRON-LT1-ES
Private Function ListaNumerosContrato() As Variant
    ListaNumerosContrato = Array("4600687006", "4600686987", "4600686987", "4600686987")
End Function

' Sufixo de estado usado apenas na composição do NOME DO ARQUIVO final
' (NÃO altera o nome das abas de origem, que continuam sem sufixo).
' NA MESMA ORDEM de ListaContratos() acima.
' Os contratos IRON-LT1-RJ/SP/ES já têm o estado embutido no próprio
' nome do contrato — por isso recebem sufixo VAZIO aqui, evitando
' duplicação (ex.: "IRON-LT1-RJ-RJ"). Apenas PA-LT2 (que não tem
' estado no nome) recebe sufixo real ("BA").
Private Function ListaSufixosEstado() As Variant
    ListaSufixosEstado = Array("BA", "", "", "")
End Function


'-----------------------------------------------------------------------
' PROCEDIMENTO PRINCIPAL — associar este ao botão
'-----------------------------------------------------------------------
Public Sub GerarArquivosPorContrato()

    Dim wbOrigem As Workbook
    Dim wsLinks As Worksheet
    Dim sPastaDestino As String
    Dim vContratos As Variant
    Dim vTipos As Variant
    Dim vNumeros As Variant
    Dim vSufixos As Variant
    Dim iContrato As Integer
    Dim iArquivosOK As Integer
    Dim sResumoFinal As String

    On Error GoTo TratarErro

    Set wbOrigem = ThisWorkbook
    vContratos = ListaContratos()
    vTipos = ListaTipos()
    vNumeros = ListaNumerosContrato()
    vSufixos = ListaSufixosEstado()

    '-- 1) PASTA DE DESTINO = SUBPASTA "MEDIÇÃO" NA RAIZ DO ARQUIVO ------
    If wbOrigem.Path = "" Then
        MsgBox "Procedimento interrompido." & vbNewLine & vbNewLine & _
            "Este arquivo ainda não foi salvo, portanto não é possível " & _
            "determinar a pasta de destino automaticamente." & vbNewLine & _
            "Salve este arquivo (Ctrl+S) e tente novamente.", _
            vbExclamation, MSG_TITULO
        Exit Sub
    End If

    sPastaDestino = ObterOuCriarSubpasta(wbOrigem.Path, NOME_SUBPASTA_DESTINO)
    If sPastaDestino = "" Then
        MsgBox "Procedimento interrompido. Não foi possível criar/acessar " & _
            "a pasta '" & NOME_SUBPASTA_DESTINO & "'.", vbCritical, MSG_TITULO
        Exit Sub
    End If

    '-- 2) LOCALIZA (OU CRIA) A ABA ONDE OS LINKS SERÃO GRAVADOS ---------
    Set wsLinks = ObterOuCriarPlanilha(wbOrigem, NOME_PLANILHA_LINKS)

    '-- 3) CONFIRMAÇÃO ANTES DE INICIAR -----------------------------------
    If MsgBox("Serão gerados " & (UBound(vContratos) + 1) & " arquivos (um por contrato):" & _
        vbNewLine & vbNewLine & _
        Join(vContratos, ", ") & vbNewLine & vbNewLine & _
        "Cada arquivo terá as abas: " & Join(vTipos, ", ") & vbNewLine & _
        "Destino: " & sPastaDestino & vbNewLine & _
        "Links serão gravados em: " & wsLinks.Name & "!" & COLUNA_CELULA_LINK & _
        LINHA_INICIAL_LINK & ":" & COLUNA_CELULA_LINK & _
        (LINHA_INICIAL_LINK + UBound(vContratos)) & vbNewLine & vbNewLine & _
        "INICIAR?", vbYesNo + vbQuestion, MSG_TITULO) = vbNo Then
        MsgBox "Procedimento cancelado.", vbInformation, MSG_TITULO
        Exit Sub
    End If

    Application.ScreenUpdating = False
    iArquivosOK = 0

    '-- 4) GERA UM ARQUIVO PARA CADA CONTRATO -----------------------------
    For iContrato = LBound(vContratos) To UBound(vContratos)
        Dim sContrato As String
        Dim sNumero As String
        Dim sSufixo As String

        sContrato = CStr(vContratos(iContrato))
        sNumero = CStr(vNumeros(iContrato))
        sSufixo = CStr(vSufixos(iContrato))

        Application.StatusBar = "Gerando arquivo do contrato " & sContrato & "..."

        Dim sCaminhoGerado As String
        sCaminhoGerado = GerarArquivoDoContrato(wbOrigem, sContrato, sNumero, _
            sSufixo, vTipos, sPastaDestino)

        If sCaminhoGerado <> "" Then
            iArquivosOK = iArquivosOK + 1
            sResumoFinal = sResumoFinal & "OK - " & sContrato & " -> " & sCaminhoGerado & vbNewLine

            '-- GRAVA O LINK DO ARQUIVO NA CÉLULA CORRESPONDENTE ------
            EscreverLinkArquivo wsLinks, iContrato, sContrato, sCaminhoGerado
        Else
            sResumoFinal = sResumoFinal & "FALHA - " & sContrato & " (ver mensagens anteriores)" & vbNewLine

            '-- Marca a célula com erro (sem link) -----------------------
            EscreverErroLink wsLinks, iContrato, sContrato
        End If
    Next iContrato

    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Processamento concluído." & vbNewLine & vbNewLine & _
        "Arquivos gerados com sucesso: " & iArquivosOK & " de " & (UBound(vContratos) + 1) & _
        vbNewLine & vbNewLine & sResumoFinal & vbNewLine & _
        "Links gravados em: " & wsLinks.Name & "!" & COLUNA_CELULA_LINK & LINHA_INICIAL_LINK & _
        ":" & COLUNA_CELULA_LINK & (LINHA_INICIAL_LINK + UBound(vContratos)), _
        vbInformation, MSG_TITULO

    Exit Sub

TratarErro:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Erro inesperado:" & vbNewLine & Err.Number & " - " & Err.Description, _
        vbCritical, MSG_TITULO
End Sub


'-----------------------------------------------------------------------
' Verifica se a subpasta informada existe dentro da pasta base; se não
' existir, tenta CRIAR. Retorna o caminho completo da subpasta (sem
' barra final), ou "" em caso de falha.
'-----------------------------------------------------------------------
Private Function ObterOuCriarSubpasta(sPastaBase As String, sNomeSubpasta As String) As String

    Dim sCaminhoCompleto As String

    On Error GoTo TratarErroLocal

    sCaminhoCompleto = sPastaBase & Application.PathSeparator & sNomeSubpasta

    If Dir(sCaminhoCompleto, vbDirectory) = "" Then
        MkDir sCaminhoCompleto
    End If

    ObterOuCriarSubpasta = sCaminhoCompleto
    Exit Function

TratarErroLocal:
    ObterOuCriarSubpasta = ""
End Function


'-----------------------------------------------------------------------
' Retorna o NÚMERO DA LINHA (7 a 10) correspondente ao índice do
' contrato no array vContratos.
'-----------------------------------------------------------------------
Private Function ObterLinhaCelulaLink(iIndiceContrato As Integer) As Integer
    ObterLinhaCelulaLink = LINHA_INICIAL_LINK + iIndiceContrato
End Function


'-----------------------------------------------------------------------
' Grava um HYPERLINK para o arquivo gerado na célula {COLUNA},{linha}
' correspondente ao contrato (coluna K, linhas 7 a 10).
'-----------------------------------------------------------------------
Private Sub EscreverLinkArquivo(ws As Worksheet, iIndiceContrato As Integer, _
    sContrato As String, sCaminhoCompleto As String)

    Dim iLinha As Integer
    Dim rngCelula As Range
    Dim sNomeArq As String

    On Error GoTo TratarErroLocal

    iLinha = ObterLinhaCelulaLink(iIndiceContrato)
    Set rngCelula = ws.Range(COLUNA_CELULA_LINK & iLinha)

    sNomeArq = Mid(sCaminhoCompleto, InStrRev(sCaminhoCompleto, Application.PathSeparator) + 1)

    rngCelula.Hyperlinks.Delete
    rngCelula.ClearContents
    rngCelula.Interior.ColorIndex = xlColorIndexNone
    rngCelula.Font.ColorIndex = xlColorIndexAutomatic

    ws.Hyperlinks.Add Anchor:=rngCelula, Address:=sCaminhoCompleto, _
        TextToDisplay:=sNomeArq

    Exit Sub

TratarErroLocal:
    On Error Resume Next
    rngCelula.Value = sCaminhoCompleto
    On Error GoTo 0
End Sub


'-----------------------------------------------------------------------
' Marca a célula correspondente ao contrato com uma mensagem de erro
' (sem hyperlink), quando a geração daquele arquivo falhou.
'-----------------------------------------------------------------------
Private Sub EscreverErroLink(ws As Worksheet, iIndiceContrato As Integer, sContrato As String)

    Dim iLinha As Integer
    Dim rngCelula As Range

    On Error Resume Next

    iLinha = ObterLinhaCelulaLink(iIndiceContrato)
    Set rngCelula = ws.Range(COLUNA_CELULA_LINK & iLinha)

    rngCelula.Hyperlinks.Delete
    rngCelula.Value = "Falha ao gerar arquivo de " & sContrato
    rngCelula.Font.Color = RGB(255, 0, 0)

    On Error GoTo 0
End Sub


'-----------------------------------------------------------------------
' Retorna a Worksheet informada; se não existir, CRIA uma nova aba com
' esse nome.
'-----------------------------------------------------------------------
Private Function ObterOuCriarPlanilha(wb As Workbook, nomeAba As String) As Worksheet
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


'-----------------------------------------------------------------------
' Gera o arquivo final de UM contrato específico, com as 4 abas
' (MC, ARM, DADOS, FRETE) copiadas a partir das respectivas abas de
' origem "{TIPO}_{CONTRATO}" da planilha principal.
'
' Nome do arquivo: "{sNumero}-PLA-MemoriaPetrobras-{sContrato}
' [-{sSufixo}]_{DDMMAAAA}_{HHMMSS}.xlsx"
' O sufixo só é anexado (com hífen) quando NÃO estiver vazio — evita
' duplicar o estado no nome dos contratos IRON, que já o possuem
' embutido (ex.: "IRON-LT1-RJ").
'
' Retorna o caminho completo do arquivo gerado, ou "" em caso de falha
' total (nenhuma das 4 abas de origem encontrada).
'-----------------------------------------------------------------------
Private Function GerarArquivoDoContrato(wbOrigem As Workbook, sContrato As String, _
    sNumero As String, sSufixo As String, vTipos As Variant, _
    sPastaDestino As String) As String

    Dim wbDestino As Workbook
    Dim iTipo As Integer
    Dim sTipo As String
    Dim sNomeAbaOrig As String
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim rngArea As Range
    Dim iCopiadas As Integer
    Dim sFaltantes As String
    Dim sNomeArquivo As String
    Dim sCaminhoCompleto As String
    Dim sSufixoComHifen As String

    On Error GoTo TratarErroLocal

    GerarArquivoDoContrato = ""

    Set wbDestino = Workbooks.Add(xlWBATWorksheet)
    Do While wbDestino.Worksheets.Count > 1
        Application.DisplayAlerts = False
        wbDestino.Worksheets(wbDestino.Worksheets.Count).Delete
        Application.DisplayAlerts = True
    Loop

    iCopiadas = 0

    For iTipo = LBound(vTipos) To UBound(vTipos)
        sTipo = CStr(vTipos(iTipo))
        sNomeAbaOrig = sTipo & "_" & sContrato ' nomes das ABAS DE ORIGEM permanecem sem sufixo

        Set wsOrigem = ObterPlanilha(wbOrigem, sNomeAbaOrig)

        If iTipo = LBound(vTipos) Then
            Set wsDestino = wbDestino.Worksheets(1)
        Else
            Set wsDestino = wbDestino.Worksheets.Add( _
                After:=wbDestino.Worksheets(wbDestino.Worksheets.Count))
        End If

        On Error Resume Next
        wsDestino.Name = sTipo
        On Error GoTo TratarErroLocal

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
        MsgBox "Nenhuma aba de origem encontrada para o contrato " & sContrato & _
            ". Arquivo não será gerado." & vbNewLine & vbNewLine & sFaltantes, _
            vbExclamation, MSG_TITULO
        wbDestino.Close SaveChanges:=False
        Exit Function
    End If

    If sFaltantes <> "" Then
        MsgBox "Atenção — contrato " & sContrato & ": algumas abas de origem " & _
            "não foram encontradas e ficaram em branco no arquivo final:" & _
            vbNewLine & vbNewLine & sFaltantes, vbExclamation, MSG_TITULO
    End If

    AplicarPerfilPublico wbDestino

    '-- Monta o NOME DO ARQUIVO, evitando duplicar o sufixo de estado ---
    If sSufixo <> "" Then
        sSufixoComHifen = "-" & sSufixo
    Else
        sSufixoComHifen = ""
    End If

    ' Ex. PA-LT2: 4600687006-PLA-MemoriaPetrobras-PA-LT2-BA_18092026_133421.xlsx
    ' Ex. IRON-LT1-RJ: 4600686987-PLA-MemoriaPetrobras-IRON-LT1-RJ_18092026_133425.xlsx
    sNomeArquivo = sNumero & "-PLA-MemoriaPetrobras-" & sContrato & _
        sSufixoComHifen & "_" & Format(Now, "ddmmyyyy") & "_" & _
        Format(Now, "hhmmss") & ".xlsx"

    sCaminhoCompleto = sPastaDestino & Application.PathSeparator & sNomeArquivo

    wbDestino.SaveAs Filename:=sCaminhoCompleto, FileFormat:=xlOpenXMLWorkbook
    wbDestino.Close SaveChanges:=False

    GerarArquivoDoContrato = sCaminhoCompleto
    Exit Function

TratarErroLocal:
    MsgBox "Erro ao gerar arquivo do contrato " & sContrato & ":" & vbNewLine & _
        Err.Number & " - " & Err.Description, vbCritical, MSG_TITULO
    On Error Resume Next
    If Not wbDestino Is Nothing Then wbDestino.Close SaveChanges:=False
    On Error GoTo 0
    GerarArquivoDoContrato = ""
End Function


'-----------------------------------------------------------------------
' Retorna a Worksheet com o nome exato informado, ou Nothing se não
' existir no workbook (sem lançar erro).
'-----------------------------------------------------------------------
Private Function ObterPlanilha(wb As Workbook, nomeAba As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(nomeAba)
    On Error GoTo 0

    Set ObterPlanilha = ws
End Function


'-----------------------------------------------------------------------
' Detecta e RETORNA o Range correspondente à área real de dados de uma
' aba TABULAR (ignorando formatação solta sem conteúdo). Usado para
' ARM, DADOS e FRETE.
'-----------------------------------------------------------------------
Private Function ObterAreaRealDados(ws As Worksheet) As Range

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


'-----------------------------------------------------------------------
' Copia um range como VALORES + FORMATOS DE CÉLULA — usado em ARM,
' DADOS e FRETE.
'-----------------------------------------------------------------------
Private Sub CopiarValoresEFormatos(rngOrigem As Range, wsDestino As Worksheet)

    On Error GoTo TratarErroLocal

    rngOrigem.Copy
    wsDestino.Range("A1").PasteSpecial Paste:=xlPasteValues
    wsDestino.Range("A1").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ReplicarLargurasColunas rngOrigem, wsDestino
    Exit Sub

TratarErroLocal:
    Application.CutCopyMode = False
End Sub


'-----------------------------------------------------------------------
' Copia um range PRESERVANDO TODA A FORMATAÇÃO — usado exclusivamente
' na aba MC (via wsOrigem.UsedRange).
'-----------------------------------------------------------------------
Private Sub CopiarComFormatacaoCompleta(rngOrigem As Range, wsDestino As Worksheet)

    Dim rngDestinoFinal As Range

    On Error GoTo TratarErroLocal

    rngOrigem.Copy
    wsDestino.Range("A1").PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False

    Set rngDestinoFinal = wsDestino.Range("A1").Resize( _
        rngOrigem.Rows.Count, rngOrigem.Columns.Count)
    rngDestinoFinal.Value = rngDestinoFinal.Value

    ReplicarLargurasColunas rngOrigem, wsDestino
    ReplicarAlturasLinhas rngOrigem, wsDestino
    Exit Sub

TratarErroLocal:
    Application.CutCopyMode = False
    On Error Resume Next
    rngOrigem.Copy
    wsDestino.Range("A1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    On Error GoTo 0
End Sub


'-----------------------------------------------------------------------
' Replica a largura das colunas do range de origem para a planilha de
' destino.
'-----------------------------------------------------------------------
Private Sub ReplicarLargurasColunas(rngOrigem As Range, wsDestino As Worksheet)
    Dim iCol As Long
    For iCol = 1 To rngOrigem.Columns.Count
        wsDestino.Columns(iCol).ColumnWidth = rngOrigem.Worksheet.Columns( _
            rngOrigem.Column + iCol - 1).ColumnWidth
    Next iCol
End Sub


'-----------------------------------------------------------------------
' Replica a altura das linhas do range de origem para a planilha de
' destino.
'-----------------------------------------------------------------------
Private Sub ReplicarAlturasLinhas(rngOrigem As Range, wsDestino As Worksheet)
    Dim iLin As Long
    For iLin = 1 To rngOrigem.Rows.Count
        wsDestino.Rows(iLin).RowHeight = rngOrigem.Worksheet.Rows( _
            rngOrigem.Row + iLin - 1).RowHeight
    Next iLin
End Sub


'-----------------------------------------------------------------------
' Aplica o rótulo de confidencialidade "Pública" ao workbook. Requer
' que a constante PUBLICO_LABEL_ID esteja preenchida com o GUID
' correto do rótulo no tenant da Petrobras.
'
' CORREÇÃO: a guarda anterior comparava a constante PUBLICO_LABEL_ID
' contra ela mesma (o próprio GUID já preenchido), o que fazia a
' condição ser SEMPRE verdadeira e a função sempre saía sem aplicar
' o rótulo. Agora a guarda compara corretamente contra string vazia,
' que é o valor esperado apenas enquanto o GUID ainda não tiver sido
' configurado.
'-----------------------------------------------------------------------
Private Sub AplicarPerfilPublico(wb As Workbook)

    If PUBLICO_LABEL_ID = "" Then Exit Sub

    On Error GoTo TratarErroLocal

    Dim lbl As Object ' Office.LabelInfo (vinculação tardia)
    Set lbl = wb.SensitivityLabel.CreateLabelInfo
    lbl.LabelId = PUBLICO_LABEL_ID
    lbl.AssignmentMethod = MSO_ASSIGNMENT_METHOD_STANDARD
    lbl.Justification = "Rótulo aplicado automaticamente na geração do arquivo."

    wb.SensitivityLabel.SetLabel lbl, lbl
    Exit Sub

TratarErroLocal:
    Debug.Print "Não foi possível aplicar o rótulo 'Pública': " & Err.Description
End Sub



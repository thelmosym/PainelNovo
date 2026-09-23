Attribute VB_Name = "A_Calc_Medicao"

' Modulo de processamento da medicao: importa os atendimentos, enriquece
' os dados, calcula os valores e atualiza as guias Fretes e MC.
Sub Inicial_Importar_e_calcular_dados_Atendimentos()




    MsgBox "Selecione o arquivo com os dados dos atendimentos a serem tratados.", vbExclamation


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
        ' Passo 4: Inserir o valor na c�lula J7 da guia "MC"
        wsMC.Range("J7").Value = valorPeriodo
    Else
        MsgBox "Per�odo n�o encontrado na guia Tabelas.", vbExclamation
    End If




    Dim wsDadosAtual As Worksheet
    Dim wsDadosFonte As Worksheet
    Dim ultimaLinhaAtual As Long
    Dim ultimaLinhaFonte As Long
    Dim caminhoArquivo As String
    Dim wbFonte As Workbook
    Dim rngFonte As Range
    
    ' Definir a planilha atual (onde os dados ser�o atualizados)
    Set wsDadosAtual = ThisWorkbook.Sheets("Dados")
    
    ThisWorkbook.Sheets("Dados").Columns("AF").Hidden = False
    
    ' Apagar dados de B3 at� a �ltima linha preenchida
    ultimaLinhaAtual = wsDadosAtual.Cells(wsDadosAtual.Rows.Count, "B").End(xlUp).Row
    If ultimaLinhaAtual >= 3 Then
        wsDadosAtual.Range("B3:AF" & ultimaLinhaAtual).ClearContents
    End If
    
    ' Abrir caixa de di�logo para selecionar o arquivo Excel
    caminhoArquivo = Application.GetOpenFilename("Arquivos Excel (*.xls; *.xlsx), *.xls; *.xlsx", , "Selecionar Arquivo Excel")
    
    ' Verificar se o usu�rio selecionou um arquivo
    If caminhoArquivo <> "False" Then
        ' Abrir o arquivo Excel selecionado
        Set wbFonte = Workbooks.Open(caminhoArquivo)
        
        ' Definir a planilha "Dados" do arquivo selecionado
        Set wsDadosFonte = wbFonte.Sheets("Dados")
        
        ' Encontrar a �ltima linha com dados na coluna B da planilha de origem
        ultimaLinhaFonte = wsDadosFonte.Cells(wsDadosFonte.Rows.Count, "B").End(xlUp).Row
        
        ' Definir o intervalo de dados a ser copiado (da c�lula B3 at� a �ltima linha na coluna X)
        Set rngFonte = wsDadosFonte.Range("B3:X" & ultimaLinhaFonte)
        
        ' Copiar os dados do arquivo de origem
        rngFonte.Copy
        
        ' Colar os dados na planilha atual a partir da c�lula B3
        wsDadosAtual.Range("B3").PasteSpecial Paste:=xlPasteValues
        
        ' Limpar a �rea de transfer�ncia para evitar a mensagem de "muita informa��o"
        Application.CutCopyMode = False
        
        ' Fechar o arquivo fonte
        wbFonte.Close SaveChanges:=False
        
        'MsgBox "Dados atualizados com sucesso!", vbInformation
    Else
        MsgBox "Nenhum arquivo foi selecionado.", vbExclamation
    End If
    
        ActiveWorkbook.Save
    
    Call Organizar_dados
End Sub




Sub Organizar_dados()

    ' Ordena os atendimentos, limpa resultados antigos e inicia o enriquecimento.
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim ultimaLinha As Long
    

    ' Define a aba "dados"
    Sheets("DADOS").Select

    
    
    ' define a �ltima linha
    ultimaLinha = Cells(Rows.Count, "B").End(xlUp).Row
    
    ' seleciona e organiza a planilha pelas colunas K, S, V e B
    Range("b2:X" & ultimaLinha).Select

    ActiveWorkbook.Worksheets("DADOS").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("DADOS").Sort.SortFields.Add2 Key:=Range("S3:S" & ultimaLinha), Order:=xlAscending
    ActiveWorkbook.Worksheets("DADOS").Sort.SortFields.Add2 Key:=Range("V3:V" & ultimaLinha), Order:=xlAscending
    ActiveWorkbook.Worksheets("DADOS").Sort.SortFields.Add2 Key:=Range("K3:K" & ultimaLinha), Order:=xlAscending
    ActiveWorkbook.Worksheets("DADOS").Sort.SortFields.Add2 Key:=Range("B3:B" & ultimaLinha), Order:=xlAscending
    With ActiveWorkbook.Worksheets("DADOS").Sort
        .SetRange Range("b2:X" & ultimaLinha)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    Range("f3:f10000").Activate
    Selection.NumberFormat = "m/d/yyyy h:mm"
    
    Range("J3:l10000").Activate
    Selection.NumberFormat = "m/d/yyyy h:mm"
    
        Range("M3:M10000").Select
    Selection.NumberFormat = "General"
            Range("h3:h10000").Select
    Selection.NumberFormat = "General"
    
    
     'Retira os filtros da planilha "Tabelas"
    Sheets("Tabelas").Select
    Range("C2:V2").Select
    Selection.AutoFilter
    Selection.AutoFilter
   ' Range("Tabela3[#Headers]").Select
    'Selection.AutoFilter
    'Selection.AutoFilter
    
    
     ' Apaga os dados das colunas referente aos dados da medi��o
    Sheets("Dados").Select
    Range("Y3:AF10000").Select
    Selection.ClearContents
    
        ActiveWorkbook.Save
        
        
   Call AtualizarLocalidades
   
    
    
End Sub
Sub AtualizarLocalidades()

    ' Converte a localidade de origem em centro usando a guia Tabelas.

    Dim wsDados As Worksheet
    Dim wsTabelas As Worksheet
    Dim ultimaLinhaDados As Long
    Dim ultimaLinhaTabelas As Long
    Dim i As Long
    Dim localidadeOrigem As String
    Dim valorEncontrado As Variant
    
    

    
    ' Definir as planilhas
    Set wsDados = ThisWorkbook.Sheets("Dados")
    Set wsTabelas = ThisWorkbook.Sheets("Tabelas")
    
    
    ' Obter a �ltima linha com dados nas planilhas
    ultimaLinhaDados = wsDados.Cells(wsDados.Rows.Count, "S").End(xlUp).Row
    ultimaLinhaTabelas = wsTabelas.Cells(wsTabelas.Rows.Count, "M").End(xlUp).Row
    
    ' Loop pelas linhas da planilha "Dados"
    For i = 2 To ultimaLinhaDados ' Pressupondo que a linha 1 cont�m cabe�alhos
        ' Obter a localidade de origem da coluna "S"
        localidadeOrigem = wsDados.Cells(i, "S").Value
        
        ' Verificar se a localidade de origem n�o est� vazia
        If Trim(localidadeOrigem) <> "" Then
            ' Procurar a localidade na planilha "Tabelas", coluna "M"
            valorEncontrado = Application.VLookup(localidadeOrigem, wsTabelas.Range("M2:N" & ultimaLinhaTabelas), 2, False)
            
            ' Inserir o valor na coluna "AD" da guia "Dados", se encontrado
            If Not IsError(valorEncontrado) Then
                wsDados.Cells(i, "AC").Value = valorEncontrado
            Else
                ' Deixar a c�lula em branco se n�o encontrar
                wsDados.Cells(i, "AC").Value = ""
            End If
        End If
    Next i

    Range("AC2").Select
    ActiveCell.FormulaR1C1 = "Centro"
    Range("AC3").Select
    
    ' Mensagem de conclus�o
    'MsgBox "Atualiza��o de localidades conclu�da!", vbInformation
        ActiveWorkbook.Save
        
        
    Call ConcatenarEInserirResultado
    

End Sub

Sub ConcatenarEInserirResultado()

    ' Cria a chave de rota Origem X Destino para as buscas posteriores.

    Dim wsDados As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim localOrigem As String
    Dim localDestino As String
    Dim chaveConcatenada As String
    
    ' Definir a planilha
    Set wsDados = ThisWorkbook.Sheets("Dados")
    

    
    ' Encontrar a �ltima linha da coluna "S" (Localidade - Origem)
    ultimaLinha = wsDados.Cells(wsDados.Rows.Count, "S").End(xlUp).Row
    
    ' Loop pelas linhas da guia Dados para concatenar as localidades
    For i = 2 To ultimaLinha
        ' Ler os valores de Localidade - Origem e Localidade - Destino
        localOrigem = wsDados.Cells(i, "AC").Value
        localDestino = wsDados.Cells(i, "V").Value
        
        ' Concatenar as localidades no formato "Localidade - Origem X Localidade - Destino"
        chaveConcatenada = localOrigem & " X " & localDestino
        
        ' Colocar o valor concatenado na coluna Y (na linha respectiva)
        wsDados.Cells(i, "AF").Value = chaveConcatenada
    Next i
    
    'MsgBox "Concata��o conclu�da e inserida na coluna AF!", vbInformation
    
        ActiveWorkbook.Save
        
        
    Call InserirLinhaServico
    

End Sub

Sub InserirLinhaServico()

    ' Relaciona a descricao da atividade com sua linha de servico PPU.

    Dim wsDados As Worksheet
    Dim wsTabelas As Worksheet
    Dim ultimaLinhaDados As Long
    Dim ultimaLinhaTabelas As Long
    Dim i As Long
    Dim descricaoAtividade As String
    Dim linhaServico As String
    Dim celulaDescricao As Range
    
    

    
    ' Definir as planilhas
    Set wsDados = ThisWorkbook.Sheets("Dados")
    Set wsTabelas = ThisWorkbook.Sheets("Tabelas")
    
    ' Encontrar a �ltima linha da coluna B (Descri��o da Atividade) na planilha "Dados"
    ultimaLinhaDados = wsDados.Cells(wsDados.Rows.Count, "B").End(xlUp).Row
    
    ' Encontrar a �ltima linha da coluna H (Descri��o da Atividade) na planilha "Tabelas"
    ultimaLinhaTabelas = wsTabelas.Cells(wsTabelas.Rows.Count, "H").End(xlUp).Row
    
    ' Loop pelas linhas da planilha "Dados"
    For i = 2 To ultimaLinhaDados
        descricaoAtividade = wsDados.Cells(i, "B").Value
        
        ' Buscar a descri��o da atividade na coluna H da planilha "Tabelas"
        Set celulaDescricao = wsTabelas.Range("H2:H" & ultimaLinhaTabelas).Find(descricaoAtividade, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' Se a descri��o for encontrada na planilha "Tabelas"
        If Not celulaDescricao Is Nothing Then
            ' Pegar a informa��o da linha de servi�o da coluna K na planilha "Tabelas"
            linhaServico = wsTabelas.Cells(celulaDescricao.Row, "K").Value
            
            ' Inserir a informa��o da linha de servi�o na coluna AE da planilha "Dados"
            wsDados.Cells(i, "AE").Value = linhaServico
        Else
            ' Caso a descri��o n�o seja encontrada, deixar a c�lula em branco ou inserir mensagem
            wsDados.Cells(i, "AE").ClearContents
        End If
    Next i
    
    Range("AE2").Select
    ActiveCell.FormulaR1C1 = "Linha de servi�o PPU 1"
    
    'MsgBox "Linha de Servi�o inserida com sucesso!", vbInformation
    
        ActiveWorkbook.Save
        
        
    Call CalcularQuilometragemAdicional
    

End Sub
Sub CalcularQuilometragemAdicional()

    ' Calcula a quilometragem adicional para fretes, evitando duplicidade
    ' por rota, data e tipo de linha de servico.

    Dim wsDados As Worksheet
    Dim wsDistancias As Worksheet
    Dim ultimaLinhaDados As Long
    Dim ultimaLinhaDistancias As Long
    Dim i As Long
    Dim chaveBusca As String
    Dim distanciaAjustada As Variant
    Dim rngDistancias As Range
    Dim celulaDistancia As Range
    Dim localOrigem As String
    Dim localDestino As String
    Dim dataFechamentoLinha As Date
    Dim dataSomente As Date
    Dim chaveConcatenadaProcessada As Collection
    Dim chaveProcessada As String
    Dim linhaServico As String
    
    ' Definir as planilhas
    Set wsDados = ThisWorkbook.Sheets("Dados")
    Set wsDistancias = ThisWorkbook.Sheets("Tabelas")
    
    ' Definir as �ltimas linhas de cada guia
    ultimaLinhaDados = wsDados.Cells(wsDados.Rows.Count, "S").End(xlUp).Row
    ultimaLinhaDistancias = wsDistancias.Cells(wsDistancias.Rows.Count, "C").End(xlUp).Row
    
    ' Definir o range de dist�ncias
    Set rngDistancias = wsDistancias.Range("C2:C" & ultimaLinhaDistancias)
    
    ' Criar uma cole��o para armazenar as chaves processadas
    Set chaveConcatenadaProcessada = New Collection
    
    ' Loop pelas linhas da guia Dados
    For i = 2 To ultimaLinhaDados
        ' Ler os dados da linha atual
        localOrigem = wsDados.Cells(i, "S").Value
        localDestino = wsDados.Cells(i, "V").Value
        
        ' Verificar se o valor da data de fechamento � uma data v�lida
        If IsDate(wsDados.Cells(i, "K").Value) Then
            dataFechamentoLinha = CDate(wsDados.Cells(i, "K").Value) ' Convertendo para data e hora
            dataSomente = Int(dataFechamentoLinha) ' Apenas a data, ignorando o hor�rio
        Else
            dataSomente = 0 ' Atribuir um valor padr�o caso n�o seja uma data
        End If
        
        ' Verificar o tipo de linha de servi�o na coluna AE
        linhaServico = wsDados.Cells(i, "AE").Value
        
        ' Verificar se a linha de servi�o � FRE-NRM ou FRE-EXP
        If linhaServico = "FRE-NRM" Or linhaServico = "FRE-EXP" Then
            ' Obter a chave concatenada da coluna AF (Localidade Origem X Destino)
            chaveBusca = wsDados.Cells(i, "AF").Value
            
            ' Verificar se essa chave j� foi processada para a mesma data e tipo de linha de servi�o
            chaveProcessada = chaveBusca & "_" & dataSomente & "_" & linhaServico
            On Error Resume Next
            chaveConcatenadaProcessada.Add chaveProcessada, chaveProcessada
            If Err.Number = 0 Then
                ' Se a chave n�o foi processada antes, buscar a dist�ncia ajustada
                Set celulaDistancia = rngDistancias.Find(What:=chaveBusca, LookIn:=xlValues, LookAt:=xlWhole)
                
                ' Se encontrado, obter a dist�ncia ajustada
                If Not celulaDistancia Is Nothing Then
                    distanciaAjustada = wsDistancias.Cells(celulaDistancia.Row, "E").Value ' Coluna de Dist�ncia Ajustada
                    
                    ' Verificar se a dist�ncia ajustada � maior que 0 e registrar
                    If distanciaAjustada > 0 Then
                        wsDados.Cells(i, "Y").Value = distanciaAjustada
                        ' Formatar o valor na coluna Y (se for inteiro, sem casas decimais, se n�o, com 1 casa decimal)
                        If distanciaAjustada = Int(distanciaAjustada) Then
                            wsDados.Cells(i, "Y").Value = Int(distanciaAjustada)
                        Else
                            wsDados.Cells(i, "Y").Value = Round(distanciaAjustada, 1)
                        End If
                        Debug.Print "Linha: " & i & " - Quilometragem registrada: " & distanciaAjustada
                    Else
                        wsDados.Cells(i, "Y").ClearContents ' Deixar em branco se n�o for aplic�vel
                        Debug.Print "Linha: " & i & " - Quilometragem n�o aplic�vel."
                    End If
                Else
                    ' Se n�o encontrado, deixar a coluna Y em branco
                    wsDados.Cells(i, "Y").ClearContents
                    Debug.Print "Linha: " & i & " - Chave n�o encontrada na guia Tabelas."
                End If
            Else
                ' Se a chave j� foi processada para a mesma data e tipo de linha de servi�o, deixar a coluna Y em branco
                wsDados.Cells(i, "Y").ClearContents
                Debug.Print "Linha: " & i & " - Chave j� processada para " & linhaServico & "."
            End If
            On Error GoTo 0
        Else
            Debug.Print "Linha: " & i & " - Linha de servi�o n�o corresponde."
        End If
    Next i
    
    'MsgBox "C�lculo de quilometragem adicional conclu�do! Verifique os logs no depurador (CTRL+G).", vbInformation

    ActiveWorkbook.Save
    
    
Call CalcularFrete

End Sub


Sub CalcularFrete()

    ' Agrupa quantidades por data, tipo de frete e rota e calcula o frete.
    Dim wsDados As Worksheet
    Dim ultimaLinhaDados As Long
    Dim dictFretes As Object
    Dim dataAtual As Variant, tipoFrete As String, localConcatenado As String
    Dim somaQuantidade As Double, valorFrete As Double
    Dim i As Long, j As Long
    Dim colData As Long, colTipoFrete As Long, colLocalConcatenado As Long, colQuantidade As Long, colResultado As Long

    ' Inicializar a planilha
    Set wsDados = ThisWorkbook.Sheets("Dados")

    ' Determinar o n�mero de linhas
    ultimaLinhaDados = wsDados.Cells(wsDados.Rows.Count, "K").End(xlUp).Row

    ' Definir �ndices de colunas
    colData = 11 ' Coluna K (Data de Fechamento)
    colTipoFrete = 31 ' Coluna AE (Tipo de Frete)
    colLocalConcatenado = 32 ' Coluna AF (Local Concatenado)
    colQuantidade = 13 ' Coluna M (Quantidade)
    colResultado = 27 ' Coluna AA (Resultado de Frete)

    ' Inicializar dicion�rio para armazenar resultados tempor�rios
    Set dictFretes = CreateObject("Scripting.Dictionary")

    ' Percorrer cada linha da aba Dados
    For i = 2 To ultimaLinhaDados
        ' Garantir que a c�lula cont�m uma data v�lida
        If IsDate(wsDados.Cells(i, colData).Value) Then
            dataAtual = Int(CDate(wsDados.Cells(i, colData).Value)) ' Converter para data, ignorando o hor�rio
        Else
            dataAtual = 0 ' Se n�o for uma data v�lida, atribui 0
        End If
        
        tipoFrete = wsDados.Cells(i, colTipoFrete).Value
        localConcatenado = wsDados.Cells(i, colLocalConcatenado).Value

        ' Validar se a data � v�lida
        If dataAtual <> 0 And (tipoFrete = "FRE-NRM" Or tipoFrete = "FRE-EXP") Then
            ' Inicializar soma de quantidades
            somaQuantidade = 0

            ' Verificar e somar as quantidades para o mesmo grupo de data, tipo de frete e local concatenado
            If Not dictFretes.exists(dataAtual & "|" & tipoFrete & "|" & localConcatenado) Then
                For j = 2 To ultimaLinhaDados
                    ' Verificar se a c�lula cont�m uma data v�lida
                    If IsDate(wsDados.Cells(j, colData).Value) Then
                        If Int(CDate(wsDados.Cells(j, colData).Value)) = dataAtual And _
                           wsDados.Cells(j, colTipoFrete).Value = tipoFrete And _
                           wsDados.Cells(j, colLocalConcatenado).Value = localConcatenado Then
                            somaQuantidade = somaQuantidade + wsDados.Cells(j, colQuantidade).Value
                        End If
                    End If
                Next j

                ' Calcular o valor do frete
                If somaQuantidade <= 10 Then
                    valorFrete = 1 ' Se a quantidade for menor ou igual a 10, o valor do frete � 1
                Else
                    valorFrete = somaQuantidade / 10 ' Caso contr�rio, divide por 10
                End If

                ' Registrar o valor na primeira ocorr�ncia e formatar
                dictFretes.Add dataAtual & "|" & tipoFrete & "|" & localConcatenado, valorFrete

                ' Se o valor for 0, deixar em branco
                If valorFrete = 0 Then
                    wsDados.Cells(i, colResultado).ClearContents ' Deixa a c�lula em branco
                ElseIf valorFrete = Int(valorFrete) Then
                    ' Se for n�mero inteiro, exibir sem casas decimais
                    wsDados.Cells(i, colResultado).Value = Int(valorFrete)
                Else
                    ' Se n�o for inteiro, exibir com 1 casa decimal
                    wsDados.Cells(i, colResultado).Value = Round(valorFrete, 1)
                End If
            End If
        End If
    Next i

    'MsgBox "C�lculo de frete conclu�do!", vbInformation
        ActiveWorkbook.Save
        
  
        
    Call CalcularPagamentoItens
    
End Sub

Sub CalcularPagamentoItens()

    ' Copia a quantidade para o pagamento dos tipos de item aplicaveis.
    Dim wsDados As Worksheet
    Dim ultimaLinhaDados As Long
    Dim i As Long
    Dim colTipoItem As Long, colQuantidade As Long, colPagamento As Long
    Dim tipoItem As String, valorQuantidade As Double

    ' Inicializar a planilha
    Set wsDados = ThisWorkbook.Sheets("Dados")

    ' Determinar o n�mero de linhas
    ultimaLinhaDados = wsDados.Cells(wsDados.Rows.Count, "K").End(xlUp).Row

    ' Definir �ndices de colunas
    colTipoItem = 31 ' Coluna AE (Tipo de Item)
    colQuantidade = 13 ' Coluna M (Quantidade)
    colPagamento = 27 ' Coluna AA (Pagamento)

    ' Percorrer cada linha da aba Dados
    For i = 2 To ultimaLinhaDados
        tipoItem = wsDados.Cells(i, colTipoItem).Value

        ' Verificar se o tipo de item corresponde a um dos itens listados
        If tipoItem = "INS-ITEM" Or tipoItem = "PES-ITEM" Or tipoItem = "DES-BEM" Or _
           tipoItem = "TRANSF" Or tipoItem = "MIGR-UA" Or tipoItem = "CAT-BEM" Or _
           tipoItem = "CAT-ITEM" Or tipoItem = "HIG-MI" Or tipoItem = "HIG-MF" Or _
           tipoItem = "HIG-PER" Or tipoItem = "CONV-MID" Or tipoItem = "GRV-MID" Or _
           tipoItem = "COP-MID" Or tipoItem = "COP-VID" Or tipoItem = "IND-DOC" Then

            ' A quantidade da coluna M ser� copiada para a coluna AA (Pagamento)
            valorQuantidade = wsDados.Cells(i, colQuantidade).Value

            ' Registrar o valor na coluna AA
            wsDados.Cells(i, colPagamento).Value = valorQuantidade

            ' Verificar se o valor � inteiro ou decimal
            If valorQuantidade = Int(valorQuantidade) Then
                wsDados.Cells(i, colPagamento).NumberFormat = "0" ' Se for inteiro, sem casas decimais
            Else
                wsDados.Cells(i, colPagamento).NumberFormat = "0.0" ' Se for decimal, com 1 casa decimal
            End If
        End If
    Next i

    'MsgBox "C�lculo de pagamento de itens conclu�do!", vbInformation
    
    
        ActiveWorkbook.Save
        
        
    Call CalcularOrgDoc
    
End Sub


Sub CalcularOrgDoc()

    ' Calcula o pagamento de organizacao de documentos conforme a descricao.
    Dim wsDados As Worksheet
    Dim ultimaLinhaDados As Long
    Dim i As Long
    Dim colTipoItem As Long, colDescricao As Long, colQuantidade As Long, colOrgDoc As Long
    Dim tipoItem As String, descricaoItem As String
    Dim quaSimples As Double, quaAnalitica As Double
    Dim orgDoc As Double

    ' Inicializar a planilha
    Set wsDados = ThisWorkbook.Sheets("Dados")

    ' Determinar o n�mero de linhas
    ultimaLinhaDados = wsDados.Cells(wsDados.Rows.Count, "K").End(xlUp).Row

    ' Definir �ndices de colunas
    colTipoItem = 31 ' Coluna AE (Tipo de Item)
    colDescricao = 2 ' Coluna B (Descri��o)
    colQuantidade = 13 ' Coluna M (Quantidade)
    colOrgDoc = 27 ' Coluna AA (Onde o resultado ser� registrado)

    ' Percorrer cada linha da aba Dados
    For i = 2 To ultimaLinhaDados
        tipoItem = wsDados.Cells(i, colTipoItem).Value
        descricaoItem = wsDados.Cells(i, colDescricao).Value

        ' Verificar se o tipo de item � "ORG-DOC"
        If tipoItem = "ORG-DOC" Then
            ' Verificar se a descri��o � "ORGANIZACAO DE DOCUMENTO SIMPLES"
            If descricaoItem = "ORGANIZACAO DE DOCUMENTO SIMPLES" Then
                ' Multiplicar a quantidade por 0,5 e registrar na coluna AA
                wsDados.Cells(i, colOrgDoc).Value = wsDados.Cells(i, colQuantidade).Value * 0.5
            End If

            ' Verificar se a descri��o � "ORGANIZACAO DE DOCUMENTO ANALITICA"
            If descricaoItem = "ORGANIZACAO DE DOCUMENTO ANALITICA" Then
                ' Registrar a quantidade diretamente na coluna AA
                wsDados.Cells(i, colOrgDoc).Value = wsDados.Cells(i, colQuantidade).Value
            End If
        End If
    Next i

    'MsgBox "C�lculo de ORG-DOC conclu�do!", vbInformation
    
        ActiveWorkbook.Save
        
        
    Call CalcularDigitalizacao
    
End Sub

Sub CalcularDigitalizacao()

    ' Busca o preco de digitalizacao e calcula quantidade vezes preco.

    Dim wsOrigem As Worksheet
    Dim wsTabelas As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim valorC As Variant
    Dim valorV As Variant
    Dim valorM As Double
    Dim resultado As Double
    Dim chaveBusca As Range
    Dim valorZ As Double

    ' Definir as planilhas
    Set wsOrigem = ThisWorkbook.Sheets("dados") ' Nome da planilha de origem
    Set wsTabelas = ThisWorkbook.Sheets("Tabelas") ' Nome da planilha Tabelas

    ' Encontrar a �ltima linha com dados na planilha de origem
    ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "AE").End(xlUp).Row

    ' Loop pelas linhas da planilha de origem
    For i = 2 To ultimaLinha ' Inicia na linha 2, considerando que a primeira linha tenha cabe�alhos
        ' Verifica se a c�lula da coluna AE tem os valores desejados
        If wsOrigem.Cells(i, "AE").Value = "DIG-MI" Or wsOrigem.Cells(i, "AE").Value = "DIG-MF" Or wsOrigem.Cells(i, "AE").Value = "DIG-DOC" Then
            
            ' Pega o valor da coluna C da linha atual
            valorC = wsOrigem.Cells(i, "C").Value
            
            ' Procura o valor na coluna T da planilha Tabelas
            Set chaveBusca = wsTabelas.Range("T:T").Find(valorC, LookIn:=xlValues, LookAt:=xlWhole)
            
            ' Se o valor for encontrado
            If Not chaveBusca Is Nothing Then
                ' Pega o valor correspondente na coluna V
                valorV = chaveBusca.Offset(0, 2).Value
                
                ' Insere o valor na coluna Z da planilha de origem como n�mero com 2 casas decimais
                wsOrigem.Cells(i, "Z").Value = valorV
                wsOrigem.Cells(i, "Z").NumberFormat = "0.00"
                
                ' Calcula o valor na coluna AA (coluna M * coluna Z)
                valorM = wsOrigem.Cells(i, "M").Value
                valorZ = wsOrigem.Cells(i, "Z").Value
                resultado = valorM * valorZ
                
                ' Insere o resultado na coluna AA e formata com 2 casas decimais
                wsOrigem.Cells(i, "AA").Value = resultado
                wsOrigem.Cells(i, "AA").NumberFormat = "0.00"
            End If
        End If
    Next i

    ActiveWorkbook.Save
    
    
Call AtualizarUnidade


End Sub


Sub AtualizarUnidade()

    ' Preenche a unidade com base na atividade e no item do atendimento.

    Dim wsDados As Worksheet
    Dim wsTabelas As Worksheet
    Dim ultimaLinhaDados As Long
    Dim ultimaLinhaTabelas As Long
    Dim i As Long
    Dim valorB As String
    Dim valorC As String
    Dim valorEncontrado As Variant
    Dim linhaEncontrada As Long
    Dim encontrou As Boolean
    
    ' Definir as planilhas
    Set wsDados = ThisWorkbook.Sheets("Dados")
    Set wsTabelas = ThisWorkbook.Sheets("Tabelas")
    
    ' Obter a �ltima linha com dados nas planilhas
    ultimaLinhaDados = wsDados.Cells(wsDados.Rows.Count, "B").End(xlUp).Row
    ultimaLinhaTabelas = wsTabelas.Cells(wsTabelas.Rows.Count, "H").End(xlUp).Row
    
    ' Loop pelas linhas da planilha "Dados"
    For i = 2 To ultimaLinhaDados ' Pressupondo que a linha 1 cont�m cabe�alhos
        ' Obter os valores das colunas "B" e "C" da guia "Dados"
        valorB = wsDados.Cells(i, "B").Value
        valorC = wsDados.Cells(i, "C").Value
        encontrou = False
        
        ' Verificar se as localidades n�o est�o vazias
        If Trim(valorB) <> "" And Trim(valorC) <> "" Then
            ' Buscar na guia "Tabelas" para as duas condi��es
            For linhaEncontrada = 2 To ultimaLinhaTabelas ' Come�ar da linha 2, assumindo cabe�alhos
                If wsTabelas.Cells(linhaEncontrada, "H").Value = valorB And wsTabelas.Cells(linhaEncontrada, "I").Value = valorC Then
                    ' Se encontrar, pegar o valor da coluna "J" e inserir na coluna "AB" da guia "Dados"
                    wsDados.Cells(i, "AB").Value = wsTabelas.Cells(linhaEncontrada, "J").Value
                    encontrou = True
                    Exit For
                End If
            Next linhaEncontrada
        End If
        
        ' Se n�o encontrar, inserir a mensagem "N�o encontrado"
        If Not encontrou Then
            wsDados.Cells(i, "AB").Value = "N�o encontrado"
        End If
    Next i
    
      Range("AB2").Select
    ActiveCell.FormulaR1C1 = "Unidade"
    
    ActiveWorkbook.Save
    
    
Call Preencher_Codigo_Centro

End Sub

Sub Preencher_Codigo_Centro()

    ' Preenche e formata o codigo do centro associado a cada localidade.

    Dim wsDados As Worksheet
    Dim wsTabelas As Worksheet
    Dim ultimaLinhaDados As Long
    Dim ultimaLinhaTabelas As Long
    Dim i As Long
    Dim localidade As String
    Dim valorEncontrado As Variant
    Dim celula As Range
    Dim valorConvertido As String
    Dim encontrado As Boolean
    
    ' Definir as planilhas
    Set wsDados = ThisWorkbook.Sheets("Dados")
    Set wsTabelas = ThisWorkbook.Sheets("Tabelas")
    
    ' Obter a �ltima linha com dados nas planilhas
    ultimaLinhaDados = wsDados.Cells(wsDados.Rows.Count, "AC").End(xlUp).Row
    ultimaLinhaTabelas = wsTabelas.Cells(wsTabelas.Rows.Count, "N").End(xlUp).Row
    
    ' Loop pelas linhas da planilha "Dados"
    For i = 2 To ultimaLinhaDados ' Supondo que a linha 1 cont�m cabe�alhos
        ' Obter a localidade da coluna AC
        localidade = wsDados.Cells(i, "AC").Value
        encontrado = False
        
        ' Pesquisar a localidade na coluna N da planilha "Tabelas"
        For Each celula In wsTabelas.Range("N2:N" & ultimaLinhaTabelas)
            If celula.Value = localidade Then
                ' Se encontrado, pegar o valor da coluna O e inserir na coluna AD
                valorEncontrado = wsTabelas.Cells(celula.Row, "O").Value
                wsDados.Cells(i, "AD").Value = valorEncontrado
                encontrado = True
                Exit For
            End If
        Next celula
        
        ' Se n�o encontrado, inserir a mensagem "N�o encontrado"
        If Not encontrado Then
            wsDados.Cells(i, "AD").Value = "N�o encontrado"
        End If
        
        ' Garantir que os valores na coluna AD estejam no formato de 4 d�gitos
        Set celula = wsDados.Cells(i, "AD")
        If IsNumeric(celula.Value) And Not IsEmpty(celula.Value) Then
            ' Aplicar o formato de 4 d�gitos (zeros � esquerda)
            Application.CutCopyMode = False
            celula.NumberFormat = "0000"
        End If
    Next i
    
        Range("AD2").Select
    ActiveCell.FormulaR1C1 = "C�digo Centro"
    
    'Formata Planilha
        Cells.Select
    Range("T1").Activate
    Cells.EntireColumn.AutoFit
    With Selection
        .HorizontalAlignment = xlGeneral
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
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
    End With

    ThisWorkbook.Sheets("Dados").Columns("AF").Hidden = True

    Range("B1:C1").Select
    'ActiveWorkbook.Save
    
    
    ActiveWorkbook.Save
    
    
 Call CopiarDadosParaFretes

End Sub

Sub CopiarDadosParaFretes()

    ' Copia para Fretes somente registros com tipo de frete e valor calculado.
    Dim wsDados As Worksheet
    Dim wsFretes As Worksheet
    Dim ultimaLinhaDados As Long
    Dim ultimaLinhaFretes As Long
    Dim i As Long
    Dim valorAE As String
    Dim valorAA As Variant
    
    ' Definir as planilhas de origem e destino
    Set wsDados = ThisWorkbook.Sheets("Dados")
    Set wsFretes = ThisWorkbook.Sheets("Fretes")
    
    ' Limpar dados antigos na planilha "Fretes"
    Sheets("Fretes").Select
    Range("A2:Q10000").Select
    Selection.ClearContents

    Range("A2:Q10000").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

    Range("A2:Q10000").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    
    
    ' Encontrar a �ltima linha preenchida na guia "Dados"
    ultimaLinhaDados = wsDados.Cells(wsDados.Rows.Count, "B").End(xlUp).Row
    
    ' Encontrar a �ltima linha preenchida na guia "Fretes" (onde os dados ser�o colados)
    ultimaLinhaFretes = wsFretes.Cells(wsFretes.Rows.Count, "A").End(xlUp).Row
    
    ' Iniciar o loop para copiar os dados a partir da linha 3 da guia "Dados"
    For i = 3 To ultimaLinhaDados
        ' Obter o valor da c�lula na coluna AE da guia "Dados"
        valorAE = wsDados.Cells(i, 31).Value ' Coluna AE � a 31� coluna
        ' Obter o valor da c�lula na coluna AA da guia "Dados"
        valorAA = wsDados.Cells(i, 27).Value ' Coluna AA � a 27� coluna
        
        ' Verificar se o valor da coluna AE � "FRE-NRM" ou "FRE-EXP"
        ' E se o valor da coluna AA for diferente de zero ou vazio
        If (valorAE = "FRE-NRM" Or valorAE = "FRE-EXP") And valorAA <> 0 And Not IsEmpty(valorAA) Then
            ' Incrementar a �ltima linha de "Fretes" para garantir que os dados sejam copiados sem lacunas
            ultimaLinhaFretes = ultimaLinhaFretes + 1
            
            ' Copiar os dados conforme a tabela fornecida
            wsFretes.Cells(ultimaLinhaFretes, 1).Value = wsDados.Cells(i, 2).Value  ' Coluna B -> Coluna A
            wsFretes.Cells(ultimaLinhaFretes, 2).Value = wsDados.Cells(i, 3).Value  ' Coluna C -> Coluna B
            wsFretes.Cells(ultimaLinhaFretes, 3).Value = wsDados.Cells(i, 6).Value  ' Coluna F -> Coluna C
            wsFretes.Cells(ultimaLinhaFretes, 4).Value = wsDados.Cells(i, 7).Value  ' Coluna G -> Coluna D
            wsFretes.Cells(ultimaLinhaFretes, 5).Value = wsDados.Cells(i, 8).Value  ' Coluna H -> Coluna E
            wsFretes.Cells(ultimaLinhaFretes, 6).Value = wsDados.Cells(i, 11).Value ' Coluna K -> Coluna F
            wsFretes.Cells(ultimaLinhaFretes, 7).Value = wsDados.Cells(i, 19).Value ' Coluna S -> Coluna G
            wsFretes.Cells(ultimaLinhaFretes, 8).Value = wsDados.Cells(i, 20).Value ' Coluna T -> Coluna H
            wsFretes.Cells(ultimaLinhaFretes, 9).Value = wsDados.Cells(i, 21).Value ' Coluna U -> Coluna I
            wsFretes.Cells(ultimaLinhaFretes, 10).Value = wsDados.Cells(i, 22).Value ' Coluna V -> Coluna J
            wsFretes.Cells(ultimaLinhaFretes, 11).Value = wsDados.Cells(i, 23).Value ' Coluna W -> Coluna K
            wsFretes.Cells(ultimaLinhaFretes, 12).Value = wsDados.Cells(i, 24).Value ' Coluna X -> Coluna L
            wsFretes.Cells(ultimaLinhaFretes, 13).Value = wsDados.Cells(i, 25).Value ' Coluna Y -> Coluna M
            wsFretes.Cells(ultimaLinhaFretes, 14).Value = wsDados.Cells(i, 27).Value ' Coluna AA -> Coluna N
            wsFretes.Cells(ultimaLinhaFretes, 15).Value = wsDados.Cells(i, 28).Value ' Coluna AB -> Coluna O
            wsFretes.Cells(ultimaLinhaFretes, 16).Value = wsDados.Cells(i, 29).Value ' Coluna AC -> Coluna P
            wsFretes.Cells(ultimaLinhaFretes, 17).Value = wsDados.Cells(i, 31).Value ' Coluna AE -> Coluna Q
        End If
    Next i

    ' Mensagem de conclus�o
    'MsgBox "Os dados foram copiados com sucesso para a guia 'Fretes'.", vbInformation
    
    ActiveWorkbook.Save
    
    
    Call AtualizarValores
    
    
End Sub

Sub AtualizarValores()

    ' Remove linhas sem valor e identifica fretes com quilometragem adicional.

    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    
    ' Definir a planilha onde os dados precisam ser ajustados
    Set ws = ThisWorkbook.Sheets("Fretes") ' Altere para o nome da sua planilha, se necess�rio
    
    ' Encontrar a �ltima linha preenchida na planilha
    ultimaLinha = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row
    
    ' Loop para percorrer as linhas a partir da linha 2 at� a �ltima linha preenchida
    For i = ultimaLinha To 2 Step -1 ' Vamos percorrer de baixo para cima para evitar problemas ao excluir linhas
        ' Verificar se a c�lula na coluna N est� vazia
        If IsEmpty(ws.Cells(i, "N").Value) Then
            ' Se a c�lula estiver vazia, deletar a linha
            ws.Rows(i).Delete
        Else
            ' Verificar se o valor da coluna M � maior que zero
            If ws.Cells(i, "M").Value > 0 Then
                ' Verificar se a coluna Q tem o valor "FRE-NRM"
                If ws.Cells(i, "Q").Value = "FRE-NRM" Then
                    ' Alterar o valor da coluna Q para "FRE-NRM/KM-Adicional"
                    ws.Cells(i, "Q").Value = "FRE-NRM/KM-Adicional"
                ' Verificar se a coluna Q tem o valor "FRE-EXP"
                ElseIf ws.Cells(i, "Q").Value = "FRE-EXP" Then
                    ' Alterar o valor da coluna Q para "FRE-EXP/KM-Adicional"
                    ws.Cells(i, "Q").Value = "FRE-EXP/KM-Adicional"
                End If
            End If
        End If
    Next i

    Range("A1").Select
    Sheets("MC").Select
    ActiveWorkbook.Save
    
    
    Call preencher_valores_MC2
    
End Sub


Sub preencher_valores_MC2()

    ' Preenche os precos da medicao conforme o lote selecionado na guia MC.
    ' Declarar vari�veis
    Dim wsMC As Worksheet
    Dim lote As String
    
    ' Definir a guia MC
    Set wsMC = ThisWorkbook.Sheets("MC")
    
    ' Obter o valor de J5 para identificar o lote
    lote = wsMC.Range("J5").Value
    
    ' Se for Lote 1 (RJ, SP, DF, ES)
    If lote = "Lote 1 - RJ" Or lote = "Lote 1 - SP" Or lote = "Lote 1 - DF" Or lote = "Lote 1 - ES" Then
        ' Preencher as c�lulas de L13 a L52 para Lote 1
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
        ' Preencher as c�lulas de L13 a L52 para Lote 2
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
    
    ' Aplicar o formato de n�mero nas c�lulas de L13 a L52
    wsMC.Range("L13:L52").NumberFormat = "0.00"
    
    'formatar valores para exibi��o de forma cont�bil
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
        
        
    ' Informar ao usu�rio que o processo foi conclu�do
    MsgBox "Dados de atendimento conclu�do!", vbInformation
    
    
    Call importar_e_calcular_ARM
    
End Sub


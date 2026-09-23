Attribute VB_Name = "subMCGuardaExterna"
'================================================================================================= ===================
'Módulo procedimento da memória de cálculo de guarda externa
'Obs.: Todo desenvolvimento para a memória de cálculo de guarda externa
'================================================================================================= ===================

Option Explicit
Const MSG_TITULO_MCA As String = "Painel de Controle - Memória de Cálculo"

Sub btMCGuardaExterna()
'Procedimento para gerar a memória de cálculo de guarda externa

'VARIÁVEIS
Dim sContratoA As String

Dim iCont As Integer

Dim iPerMes As Integer
Dim sPerMes As String
Dim iPerAno As Integer
Dim sPeriodo As String
Dim sMensagem As String
Dim sRe() As String
Dim iTSPR1 As Integer
Dim iTSPE1 As Integer
Dim iTSPR2 As Integer
Dim iTSPE2 As Integer

Dim iULin As Integer
Dim i As Integer
Dim sContratoN As String

Dim sAgrupamento As String
Dim sSituacao As String

Dim dirTemplate As String

Dim dIAPFARQ1 As Double
Dim dIAPFARQ2 As Double
Dim dFDM1 As Double
Dim dFDM2 As Double

Dim wbTemplate As Workbook
Dim wsTempl_MC As Worksheet
Dim wsTempl_ARM As Worksheet
Dim wsTempl_DADOS As Worksheet
Dim wsTempl_MEDICAO As Worksheet

Dim sDescrAtividade As String
Dim dQExec As Double
Dim dFA As Double

Dim sUnidade As String
Dim sLinhaPPU As String
Dim dKM As Double

Dim sNomeArquivo As String
Dim vSalvarComo As Variant

Dim sEstado As String
Dim sCentro As String
Dim sCodcentro As String
Dim sDado1 As String
Dim sDado2 As String
Dim iRefFDM As Integer
Dim bEncontrou As Boolean
Dim n As Integer
Dim sPerAnt1 As String
Dim sPerAnt2 As String

'Verifica se existe alguma célula obrigatória em branco
'################################################################################################# ###################
If verificar_celulaEmBranco() Then
    MsgBox _
    "Não será possível gerar a memória de cálculo. Foi encontrado células obrigatórias em branco.", _
    vbCritical, MSG_TITULO_MCA
    Exit Sub
End If

'VERIFICAR LOG
'################################################################################################# ###################
iCont = WorksheetFunction.CountA(Planilha7.Range("AG3:AG2003"))
If iCont > 0 Then
    MsgBox _
    "Não será possível gerar a memória de cálculo. A coluna LOG possui pendência(s). Por favor, verificar cada linha.", _
    vbCritical, MSG_TITULO_MCA
    Exit Sub
End If
iCont = Empty


'CONTRATO PARA PREENCHIMENTO DE DADOS
'#####################################################################################################################
sContratoA = Planilha6.Range("E18").Value


'VERIFICAR STATUS PARA AGRUPAMENTO
'#####################################################################################################################
iULin = Planilha7.Range("Q2003").End(xlUp).Row
For i = 3 To iULin
    sContratoN = Planilha7.Cells(i, 3)
    If sContratoA = sContratoN Then
        sAgrupamento = Planilha7.Cells(i, 17)
        If sAgrupamento <> "" Then
            sSituacao = Planilha7.Cells(i, 2)
            If sSituacao <> "CO" Then
                MsgBox "Não será permitido gerar a Memória de Cálculo." & _
                " Algumas linhas de agrupamento possuem situação diferente de concluído (CO).", _
                vbExclamation, MSG_TITULO_MCA
                Exit Sub
            End If
        End If
    End If
Next i
sContratoN = ""
sAgrupamento = ""
sSituacao = ""


'MENSAGEM INFORMANDO PERÍODO A COLHER DADOS
'################################################################################################# ###################
sPerMes = UCase(Mid(Planilha6.Range("C9").Value, 4, 20)) 'Descrição Mês - Período informado na planilha Menu
iPerMes = Left(Planilha6.Range("C9").Value, 2) 'Número Mês - Período informado na planilha Menu
iPerAno = Planilha6.Range("C10").Value 'Ano - Período informado na planilha Menu

sMensagem = MsgBox("Dados a serem exportados para a memória de cálculo" & Chr(10) & "Período: " & sPerMes & "/" & iPerAno & _
Chr(10) & "Contrato: " & sContratoA & Chr(10) & Chr(10) & "INICIAR?", vbYesNo, MSG_TITULO_MCA)
If sMensagem = vbNo Then
    MsgBox "Exportação cancelada.", vbInformation, MSG_TITULO_MCA
    Exit Sub
End If

sPeriodo = gerar_periodo(iPerMes, iPerAno)

Application.StatusBar = "Procedimento em progresso. Preenchendo cabeçalho..."

'CALCULANDO IAPEFARQ E FDM
'################################################################################################# ###################
'Verificando linhas do painel para calcular TSPE e TSPR
sRe() = Split(calcular_TSPE_TSPR(sContratoA), "|")
iTSPE1 = sRe(0)
iTSPR1 = sRe(1)
iTSPE2 = sRe(2)
iTSPR2 = sRe(3)

'calculando medidas para serviço 1 da PPU (1.1 a 1.14)
dIAPFARQ1 = Format(calcular_IAPFARQ(iTSPE1, iTSPR1), "#.00")
dFDM1 = calcular_FDM(dIAPFARQ1)

'calculando medidas para serviço 2 da PPU (2.1 a 2.14 (a 2.13 do lote 2)
dIAPFARQ2 = Format(calcular_IAPFARQ(iTSPE2, iTSPR2), "#.00")
dFDM2 = calcular_FDM(dIAPFARQ2)

'RECALCULANDO FDM CONFORME PERÍODOS ANTERIOS (Item 8.4.3 do contrato de guarda externa)
'Referente ao primeiro Mês anterior ao período vigente
If iPerMes - 1 = 0 Then
    sPerAnt1 = iPerAno - 1 & "#" & "12"
    
Else
    sPerAnt1 = iPerAno & "#" & Format(iPerMes - 1, "00")
    
End If


'Referente ao segundo Mês anterior ao período vigente
If iPerMes - 2 = -1 Then
    sPerAnt2 = iPerAno - 1 & "#" & "11"
    
ElseIf iPerMes - 2 = 0 Then
    sPerAnt2 = iPerAno - 1 & "#" & "12"
    
Else
    sPerAnt2 = iPerAno & "#" & Format(iPerMes - 2, "00")
    
End If

'Verifica se períodos anteriores foram inseridos na tabela
If Not validar_periodo_anterior(sPerAnt1) Then
    MsgBox "Não será permitido gerar a Memória de Cálculo. Dados do período anterior não foram encontratos. Siga as instruções abaixo:" & _
    Chr(10) & Chr(10) & _
    "1) Atualize o arquivo TabelaB na rede com dados: " & sPerAnt1 & Chr(10) & _
    "2) Atualize o painel de controle e tente novamente.", vbExclamation, MSG_TITULO_MCA
    Exit Sub
End If

If Not validar_periodo_anterior(sPerAnt2) Then
    MsgBox "Não será permitido gerar a Memória de Cálculo. Dados do período anterior não foram encontratos. Siga as instruções abaixo:" & _
    Chr(10) & Chr(10) & _
    "1) Atualize o arquivo TabelaB na rede com dados: " & sPerAnt2 & Chr(10) & _
    "2) Atualize o painel de controle e tente novamente.", vbExclamation, MSG_TITULO_MCA
    Exit Sub
End If


'FDM recalculado conforme períodos anteriores
dFDM1 = recalcular_FDM(iPerAno & "#" & Format(iPerMes, "00"), sPerAnt1, sPerAnt2, sContratoA, dFDM1, 1)
dFDM2 = recalcular_FDM(iPerAno & "#" & Format(iPerMes, "00"), sPerAnt1, sPerAnt2, sContratoA, dFDM2, 2)


'VERIFICAR DADOS DA REDE PARA LEITURA DE PASTAS
'#####################################################################################################################
Select Case sContratoA
    Case Is = "IRON-LT1"
        dirTemplate = Planilha5.Range("C14")
    Case Is = "PA-LT2"
        dirTemplate = Planilha5.Range("C23")
End Select

'Verifica se arquivo existe na pasta/diretório
If Dir(dirTemplate) = "" Then
    MsgBox "Procedimento interrompido. Arquivo e/ou diretório não existe." & Chr(10) & _
    Chr(10) & _
    "Obs.: Verifique diretório informado e tente novamente.", vbCritical, MSG_TITULO_MCA
    Exit Sub
End If


'PREENCHIMENTO DO TEMPLATE COM DADOS DO PAINEL DE CONTROLE
'#####################################################################################################################
Set wbTemplate = Workbooks.Open(dirTemplate) 'Abre template
Set wsTempl_MC = wbTemplate.Worksheets(1)
Set wsTempl_ARM = wbTemplate.Worksheets(2)
Set wsTempl_DADOS = wbTemplate.Worksheets(3)
Set wsTempl_MEDICAO = wbTemplate.Worksheets(4)

'Preenchimento referente ao cabeçalho do template > planilha MC
wsTempl_MC.Range("J7") = sPeriodo
wsTempl_MC.Range("N4") = dIAPFARQ1
wsTempl_MC.Range("N5") = dFDM1
wsTempl_MC.Range("N6") = dIAPFARQ2
wsTempl_MC.Range("N7") = dFDM2

'Exportação de dados DE: Painel de controle > Armazenamento, PARA: Template > planilha ARM
Application.StatusBar = "Procedimento em progresso. Preenchendo planilha ARM..."
i = 2
Do While Planilha8.Cells(i, 1) <> ""
    If InStr(1, Planilha8.Cells(i, 1), sContratoA) > 0 Then
        iULin = wsTempl_ARM.Range("A5000").End(xlUp).Row + 1
        wsTempl_ARM.Cells(iULin, 1).Value = Planilha8.Cells(i, 1).Value
        wsTempl_ARM.Cells(iULin, 2).Value = Planilha8.Cells(i, 2).Value
        wsTempl_ARM.Cells(iULin, 3).Value = Planilha8.Cells(i, 3).Value
        wsTempl_ARM.Cells(iULin, 4).Value = Planilha8.Cells(i, 4).Value
        wsTempl_ARM.Cells(iULin, 5).Value = Planilha8.Cells(i, 5).Value
        wsTempl_ARM.Cells(iULin, 6).Value = Planilha8.Cells(i, 6).Value
        wsTempl_ARM.Cells(iULin, 7).Value = Planilha8.Cells(i, 7).Value
        wsTempl_ARM.Cells(iULin, 8).Value = Planilha8.Cells(i, 8).Value
        wsTempl_ARM.Cells(iULin, 9).Value = WorksheetFunction.Round(Planilha8.Cells(i, 9).Value, 3)
        wsTempl_ARM.Cells(iULin, 10).Value = wsTempl_ARM.Cells(iULin, 8).Value * wsTempl_ARM.Cells(iULin, 9).Value
        wsTempl_ARM.Cells(iULin, 11).Value = _
        pesquisar_linhaPPU_arm("ARMAZENAMENTO DE ACERVO DOCUMENTAL", wsTempl_ARM.Cells(iULin, 3)) 'erro encontrado
        
    End If
i = i + 1
Loop

                
'Exportação de dados DE: Painel de controle > Painel, PARA: Template > planilha DADOS
Application.StatusBar = "Procedimento em progresso. Preenchendo planilha DADOS..."
i = 3
Do While Planilha7.Cells(i, 7) <> ""
    sContratoN = Planilha7.Cells(i, 3)
    If sContratoA = sContratoN Then
        sSituacao = Planilha7.Cells(i, 2)
        If sSituacao = "CO" Then
            sDescrAtividade = Planilha7.Cells(i, 4)
            If pesquisar_referenciaQExec(sDescrAtividade) = "S" Then
                iULin = wsTempl_DADOS.Range("B5000").End(xlUp).Row + 1
                wsTempl_DADOS.Cells(iULin, 2).Value = Planilha7.Cells(i, 4).Value 'Descrição da atividade
                wsTempl_DADOS.Cells(iULin, 3).Value = Planilha7.Cells(i, 5).Value 'Item
                wsTempl_DADOS.Cells(iULin, 4).Value = Planilha7.Cells(i, 6).Value 'Aplicação
                wsTempl_DADOS.Cells(iULin, 5).Value = Planilha7.Cells(i, 7).Value 'Código da solicitação
                wsTempl_DADOS.Cells(iULin, 6).Value = Planilha7.Cells(i, 9).Value 'D. solicitação
                wsTempl_DADOS.Cells(iULin, 7).Value = Planilha7.Cells(i, 10).Value 'Gerência solicitante
                wsTempl_DADOS.Cells(iULin, 8).Value = Planilha7.Cells(i, 11).Value 'Qtd. (solicitada)
                wsTempl_DADOS.Cells(iULin, 9).Value = Planilha7.Cells(i, 12).Value 'Código OS
                wsTempl_DADOS.Cells(iULin, 10).Value = Planilha7.Cells(i, 13).Value 'D. abertura
                wsTempl_DADOS.Cells(iULin, 11).Value = Planilha7.Cells(i, 14).Value 'D. fechamento
                wsTempl_DADOS.Cells(iULin, 12).Value = Planilha7.Cells(i, 15).Value 'Prazo combinado
                wsTempl_DADOS.Cells(iULin, 13).Value = Planilha7.Cells(i, 16).Value 'Qtd. (Atendida)
                wsTempl_DADOS.Cells(iULin, 14).Value = Planilha7.Cells(i, 17).Value 'Agrupamento
                wsTempl_DADOS.Cells(iULin, 15).Value = Planilha7.Cells(i, 18).Value 'Prazo
                wsTempl_DADOS.Cells(iULin, 16).Value = Planilha7.Cells(i, 19).Value 'FDM
                wsTempl_DADOS.Cells(iULin, 17).Value = Planilha7.Cells(i, 20).Value 'Obs. isenção
                wsTempl_DADOS.Cells(iULin, 18).Value = Planilha7.Cells(i, 21).Value 'FA
                wsTempl_DADOS.Cells(iULin, 19).Value = Planilha7.Cells(i, 22).Value 'Dados PETROBRAS Localidade
                wsTempl_DADOS.Cells(iULin, 20).Value = Planilha7.Cells(i, 23).Value 'Dados PETROBRAS Município
                wsTempl_DADOS.Cells(iULin, 21).Value = Planilha7.Cells(i, 24).Value 'Dados PETROBRAS UF
                wsTempl_DADOS.Cells(iULin, 22).Value = Planilha7.Cells(i, 25).Value 'Dados galpão Localidade
                wsTempl_DADOS.Cells(iULin, 23).Value = Planilha7.Cells(i, 26).Value 'Dados galpão município
                wsTempl_DADOS.Cells(iULin, 24).Value = Planilha7.Cells(i, 27).Value 'Dados galpão UF
                
                'Conforme solicitação da PETROBRAS, a atividade MATERIAL PARA ARQUIVAMENTO entrará para a memória de cálculo, mas não terá nenhum cálculo
                If wsTempl_DADOS.Cells(iULin, 2) = "MATERIAL PARA ARQUIVAMENTO" Then GoTo proximaLinha
                
                'Funcionalidade KM adicional
                If validar_localizacao(wsTempl_DADOS.Cells(iULin, 2)) Then
                    sCentro = pesquisar_centro(wsTempl_DADOS.Cells(iULin, 19))
                    dKM = calcular_KMAdicional(sCentro, wsTempl_DADOS.Cells(iULin, 22))
                    If dKM > 0 Then
                        wsTempl_DADOS.Cells(iULin, 25) = dKM
                    End If
                Else
                    sCentro = pesquisar_centro(wsTempl_DADOS.Cells(iULin, 19))
                End If
                
                'Funcionalidade FC
                
                
                'Funcionalidade QExec
                dQExec = _
                calcular_QExec(wsTempl_DADOS.Cells(iULin, 2), wsTempl_DADOS.Cells(iULin, 3), wsTempl_DADOS.Cells(iULin, 13), wsTempl_DADOS.Cells(iULin, 14))
                'Verifica o fator de ajuste inserido
                If wsTempl_DADOS.Cells(iULin, 18).Value = Empty Then
                    dFA = 1
                Else
                    dFA = wsTempl_DADOS.Cells(iULin, 18).Value
                End If
                wsTempl_DADOS.Cells(iULin, 27) = dQExec * dFA
                
                'Funcionalidade Unidade
                sUnidade = pesquisar_unidade(wsTempl_DADOS.Cells(iULin, 2))
                
                'Funcionalidade Linha PPU
                sLinhaPPU = pesquisar_linhaPPU(wsTempl_DADOS.Cells(iULin, 2))
                
                wsTempl_DADOS.Cells(iULin, 28) = sUnidade
                
                'Funcionalidade Centro
                wsTempl_DADOS.Cells(iULin, 29) = sCentro
                
                'Funcionalidade Código Centro
                'Obs.: Variável sCentro utilizada da pesquisa anterior
                wsTempl_DADOS.Cells(iULin, 30) = pesquisar_codCentro(sCentro)
                
                'Funcionalidade Linha de serviço PPU
                wsTempl_DADOS.Cells(iULin, 31) = sLinhaPPU
                
            End If
        End If
    End If
proximaLinha:
sCentro = ""
i = i + 1
Loop

sDescrAtividade = ""
dQExec = Empty
sLinhaPPU = ""

'Preenchimento referente ao template > planilha MEDICAO
Application.StatusBar = "Procedimento em progresso. Preenchendo planilha MEDICAO..."
i = 3
Do While wsTempl_DADOS.Cells(i, 2) <> ""
    sDescrAtividade = wsTempl_DADOS.Cells(i, 2)
    If sDescrAtividade = "MATERIAL PARA ARQUIVAMENTO" Then GoTo proximaLinha2
    sEstado = wsTempl_DADOS.Cells(i, 21)
    dQExec = wsTempl_DADOS.Cells(i, 27)
    sCentro = wsTempl_DADOS.Cells(i, 29)
    sCodcentro = wsTempl_DADOS.Cells(i, 30)
    sLinhaPPU = wsTempl_DADOS.Cells(i, 31)
    sDado1 = sEstado & "#" & sCentro & "#" & sDescrAtividade
    
    iCont = WorksheetFunction.CountIf(wsTempl_MEDICAO.Range("L:L"), sDado1)
    
    If iCont = 0 Then 'Preenchendo nova linha no caso de não encontrado
        iULin = wsTempl_MEDICAO.Range("D5000").End(xlUp).Row + 1
        wsTempl_MEDICAO.Cells(iULin, 2) = sEstado 'Estado
        wsTempl_MEDICAO.Cells(iULin, 3) = sCentro & "#" & sCodcentro 'Centro#Cód. centro
        wsTempl_MEDICAO.Cells(iULin, 4) = sDescrAtividade 'Descrição da atividade
        wsTempl_MEDICAO.Cells(iULin, 5) = pesquisar_itemPPU(sLinhaPPU) 'Item PPU (Pesquisa na tabelaA)
        wsTempl_MEDICAO.Cells(iULin, 6) = sLinhaPPU 'Linha de serviço PPU
        wsTempl_MEDICAO.Cells(iULin, 7) = dQExec 'QExec
        iRefFDM = pesquisar_referenciaFDM(sDescrAtividade) 'FDM (Pesquisa na tabelaA)
        If sDescrAtividade = "TRANSFERENCIA DE ACERVO" Then
            wsTempl_MEDICAO.Cells(iULin, 8) = 1
        Else
            If iRefFDM = 1 Then
                wsTempl_MEDICAO.Cells(iULin, 8) = dFDM1
            ElseIf iRefFDM = 2 Then
                wsTempl_MEDICAO.Cells(iULin, 8) = dFDM2
            End If
        End If
        wsTempl_MEDICAO.Cells(iULin, 9).FormulaLocal = "=G" & iULin & "*H" & iULin 'QFinal
        wsTempl_MEDICAO.Cells(iULin, 10) = pesquisar_precoUnitario(sContratoA, sLinhaPPU) 'preço unitário  (Pesquisa na tabelaB)
        wsTempl_MEDICAO.Cells(iULin, 11).FormulaLocal = "=I" & iULin & "*J" & iULin 'Subtotal
        wsTempl_MEDICAO.Cells(iULin, 12) = sDado1 'dados concatenados (Referência de consulta)
    Else
    
        bEncontrou = False
        n = 4
        Do While wsTempl_MEDICAO.Cells(n, 12) <> "" And Not bEncontrou
        sDado2 = wsTempl_MEDICAO.Cells(n, 12)
        
            If sDado1 = sDado2 Then
            wsTempl_MEDICAO.Cells(n, 7) = wsTempl_MEDICAO.Cells(n, 7) + dQExec
            bEncontrou = True
            End If
          
        n = n + 1
        Loop

    End If
proximaLinha2:
i = i + 1
Loop

'Limpar dados não necessários para a memória de cálculo . planilha MEDICAO
wsTempl_MEDICAO.Range("L:L").ClearContents

'Funcionalidade para salvar arquivo
Application.StatusBar = "Procedimento em progresso. Salvando Arquivo Como..."
sNomeArquivo = _
wsTempl_MC.Range("D5") & "-PLA-MemoriaPetrobras-" & iPerAno & iPerMes & "-" & wsTempl_MC.Range("J5") & _
"#" & Year(Now) & "-" & Month(Now) & "-" & Day(Now) & "_" & Hour(Now) & "h" & Minute(Now) & "min"

vSalvarComo = Application.GetSaveAsFilename(InitialFileName:=sNomeArquivo, filefilter:="Excel Files (*.xlsx),*.xlsx")
If vSalvarComo <> False Then
    ActiveWorkbook.SaveAs filename:=vSalvarComo
End If

wbTemplate.Close (False)

Set wbTemplate = Nothing
Set wsTempl_MC = Nothing
Set wsTempl_ARM = Nothing
Set wsTempl_DADOS = Nothing
Set wsTempl_MEDICAO = Nothing

Application.StatusBar = False

End Sub

Sub btMC_GE_individual()
'Procedimento para gerar memória de cálculo de guarda externa individual

'VARIÁVEIS
Dim sContratoA As String

Dim iCont As Integer

Dim iPerMes As Integer
Dim sPerMes As String
Dim iPerAno As Integer
Dim sPeriodo As String
Dim sMensagem As String
Dim sRe() As String
Dim iTSPR1 As Integer
Dim iTSPE1 As Integer
Dim iTSPR2 As Integer
Dim iTSPE2 As Integer

Dim iULin As Integer
Dim i As Integer
Dim sContratoN As String

Dim sAgrupamento As String
Dim sSituacao As String

Dim dirTemplate As String

Dim dIAPFARQ1 As Double
Dim dIAPFARQ2 As Double
Dim dFDM1 As Double
Dim dFDM2 As Double

Dim wbTemplate As Workbook
Dim wsTempl_MC As Worksheet
Dim wsTempl_ARM As Worksheet
Dim wsTempl_DADOS As Worksheet
Dim wsTempl_MEDICAO As Worksheet

Dim sDescrAtividade As String
Dim dQExec As Double
Dim dFA As Double

Dim sUnidade As String
Dim sLinhaPPU As String
Dim dKM As Double

Dim sNomeArquivo As String
Dim vSalvarComo As Variant

Dim sEstado As String
Dim sCentro As String
Dim sCodcentro As String
Dim sDado1 As String
Dim sDado2 As String
Dim iRefFDM As Integer
Dim bEncontrou As Boolean
Dim n As Integer
Dim sPerAnt1 As String
Dim sPerAnt2 As String

Dim sCntrtEstd As String '<------------------------

'Verifica se existe alguma célula obrigatória em branco
'################################################################################################# ###################
If verificar_celulaEmBranco() Then
    MsgBox _
    "Não será possível gerar a memória de cálculo. Foi encontrado células obrigatórias em branco.", _
    vbCritical, MSG_TITULO_MCA
    Exit Sub
End If


'VERIFICAR LOG
'################################################################################################# ###################
iCont = WorksheetFunction.CountA(Planilha7.Range("AG3:AG2003"))
If iCont > 0 Then
    MsgBox _
    "Não será possível gerar a memória de cálculo. A coluna LOG possui pendência(s). Por favor, verificar cada linha.", _
    vbCritical, MSG_TITULO_MCA
    Exit Sub
End If
iCont = Empty


'CONTRATO PARA PREENCHIMENTO DE DADOS
'#####################################################################################################################
sContratoA = Left(Planilha6.Range("E22").Value, Len(Planilha6.Range("E22").Value) - 3) '--
sCntrtEstd = Right(Planilha6.Range("E22").Value, 2) '--

'VERIFICAR STATUS PARA AGRUPAMENTO
'#####################################################################################################################
iULin = Planilha7.Range("Q2003").End(xlUp).Row
For i = 3 To iULin
    sContratoN = Planilha7.Cells(i, 3)
    If sContratoA = sContratoN Then
        sAgrupamento = Planilha7.Cells(i, 17)
        If sAgrupamento <> "" Then
            sSituacao = Planilha7.Cells(i, 2)
            If sSituacao <> "CO" Then
                MsgBox "Não será permitido gerar a Memória de Cálculo." & _
                " Algumas linhas de agrupamento possuem situação diferente de concluído (CO).", _
                vbExclamation, MSG_TITULO_MCA
                Exit Sub
            End If
        End If
    End If
Next i
sContratoN = ""
sAgrupamento = ""
sSituacao = ""


'MENSAGEM INFORMANDO PERÍODO A COLHER DADOS
'################################################################################################# ###################
sPerMes = UCase(Mid(Planilha6.Range("C9").Value, 4, 20)) 'Descrição Mês - Período informado na planilha Menu
iPerMes = Left(Planilha6.Range("C9").Value, 2) 'Número Mês - Período informado na planilha Menu
iPerAno = Planilha6.Range("C10").Value 'Ano - Período informado na planilha Menu

sMensagem = MsgBox("Dados a serem exportados para a memória de cálculo." & _
Chr(10) & _
"Período: " & sPerMes & "/" & iPerAno & _
Chr(10) & _
"Contrato: " & sContratoA & "-" & sCntrtEstd & _
Chr(10) & Chr(10) & _
"<<A T E N Ç Ã O>>: Dados exportados para a memória de cálculo serão apenas da região/estado selecionado." & _
Chr(10) & Chr(10) & _
"INICIAR?", vbInformation + vbYesNo, MSG_TITULO_MCA)
If sMensagem = vbNo Then
    MsgBox "Exportação cancelada.", vbCritical, MSG_TITULO_MCA
    Exit Sub
End If

sPeriodo = gerar_periodo(iPerMes, iPerAno)

Application.StatusBar = "Procedimento em progresso. Preenchendo cabeçalho..."

'CALCULANDO IAPEFARQ E FDM
'################################################################################################# ###################
'Verificando linhas do painel para calcular TSPE e TSPR
sRe() = Split(calcular_TSPE_TSPR(sContratoA), "|")
iTSPE1 = sRe(0)
iTSPR1 = sRe(1)
iTSPE2 = sRe(2)
iTSPR2 = sRe(3)

'calculando medidas para serviço 1 da PPU (1.1 a 1.14)
dIAPFARQ1 = Format(calcular_IAPFARQ(iTSPE1, iTSPR1), "#.00")
dFDM1 = calcular_FDM(dIAPFARQ1)

'calculando medidas para serviço 2 da PPU (2.1 a 2.14 (a 2.13 do lote 2)
dIAPFARQ2 = Format(calcular_IAPFARQ(iTSPE2, iTSPR2), "#.00")
dFDM2 = calcular_FDM(dIAPFARQ2)

'RECALCULANDO FDM CONFORME PERÍODOS ANTERIOS (Item 8.4.3 do contrato de guarda externa)
'Referente ao primeiro Mês anterior ao período vigente
If iPerMes - 1 = 0 Then
    sPerAnt1 = iPerAno - 1 & "#" & "12"
    
Else
    sPerAnt1 = iPerAno & "#" & Format(iPerMes - 1, "00")
    
End If


'Referente ao segundo Mês anterior ao período vigente
If iPerMes - 2 = -1 Then
    sPerAnt2 = iPerAno - 1 & "#" & "11"
    
ElseIf iPerMes - 2 = 0 Then
    sPerAnt2 = iPerAno - 1 & "#" & "12"
    
Else
    sPerAnt2 = iPerAno & "#" & Format(iPerMes - 2, "00")
    
End If

'Verifica se períodos anteriores foram inseridos na tabela
If Not validar_periodo_anterior(sPerAnt1) Then
    MsgBox "Não será permitido gerar a Memória de Cálculo. Dados do período anterior não foram encontratos. Siga as instruções abaixo:" & _
    Chr(10) & Chr(10) & _
    "1) Atualize o arquivo TabelaB na rede com dados: " & sPerAnt1 & Chr(10) & _
    "2) Atualize o painel de controle e tente novamente.", vbExclamation, MSG_TITULO_MCA
    Exit Sub
End If

If Not validar_periodo_anterior(sPerAnt2) Then
    MsgBox "Não será permitido gerar a Memória de Cálculo. Dados do período anterior não foram encontratos. Siga as instruções abaixo:" & _
    Chr(10) & Chr(10) & _
    "1) Atualize o arquivo TabelaB na rede com dados: " & sPerAnt2 & Chr(10) & _
    "2) Atualize o painel de controle e tente novamente.", vbExclamation, MSG_TITULO_MCA
    Exit Sub
End If

'FDM recalculado conforme períodos anteriores
dFDM1 = recalcular_FDM(iPerAno & "#" & Format(iPerMes, "00"), sPerAnt1, sPerAnt2, sContratoA, dFDM1, 1)
dFDM2 = recalcular_FDM(iPerAno & "#" & Format(iPerMes, "00"), sPerAnt1, sPerAnt2, sContratoA, dFDM2, 2)

'VERIFICAR DADOS DA REDE PARA LEITURA DE PASTAS
'#####################################################################################################################
Select Case sContratoA
    Case Is = "IRON-LT1"
        dirTemplate = Planilha5.Range("C14")
    Case Is = "PA-LT2"
        dirTemplate = Planilha5.Range("C23")
End Select

'Verifica se arquivo existe na pasta/diretório
If Dir(dirTemplate) = "" Then
    MsgBox "Procedimento interrompido. Arquivo e/ou diretório não existe." & Chr(10) & _
    Chr(10) & _
    "Obs.: Verifique diretório informado e tente novamente.", vbCritical, MSG_TITULO_MCA
    Exit Sub
End If


'PREENCHIMENTO DO TEMPLATE COM DADOS DO PAINEL DE CONTROLE
'#####################################################################################################################
Set wbTemplate = Workbooks.Open(dirTemplate) 'Abre template
Set wsTempl_MC = wbTemplate.Worksheets(1)
Set wsTempl_ARM = wbTemplate.Worksheets(2)
Set wsTempl_DADOS = wbTemplate.Worksheets(3)
Set wsTempl_MEDICAO = wbTemplate.Worksheets(4)

'Preenchimento referente ao cabeçalho do template > planilha MC
wsTempl_MC.Range("J5") = wsTempl_MC.Range("J5") & " - " & sCntrtEstd '<------------------
wsTempl_MC.Range("J7") = sPeriodo
wsTempl_MC.Range("N4") = dIAPFARQ1
wsTempl_MC.Range("N5") = dFDM1
wsTempl_MC.Range("N6") = dIAPFARQ2
wsTempl_MC.Range("N7") = dFDM2

'Exportação de dados DE: Painel de controle > Armazenamento, PARA: Template > planilha ARM
Application.StatusBar = "Procedimento em progresso. Preenchendo planilha ARM..."
i = 2
Do While Planilha8.Cells(i, 1) <> ""
    If InStr(1, Planilha8.Cells(i, 1), sContratoA & "-" & sCntrtEstd) > 0 Then  '<------------------------
        iULin = wsTempl_ARM.Range("A5000").End(xlUp).Row + 1
        wsTempl_ARM.Cells(iULin, 1).Value = Planilha8.Cells(i, 1).Value
        wsTempl_ARM.Cells(iULin, 2).Value = Planilha8.Cells(i, 2).Value
        wsTempl_ARM.Cells(iULin, 3).Value = Planilha8.Cells(i, 3).Value
        wsTempl_ARM.Cells(iULin, 4).Value = Planilha8.Cells(i, 4).Value
        wsTempl_ARM.Cells(iULin, 5).Value = Planilha8.Cells(i, 5).Value
        wsTempl_ARM.Cells(iULin, 6).Value = Planilha8.Cells(i, 6).Value
        wsTempl_ARM.Cells(iULin, 7).Value = Planilha8.Cells(i, 7).Value
        wsTempl_ARM.Cells(iULin, 8).Value = Planilha8.Cells(i, 8).Value
        wsTempl_ARM.Cells(iULin, 9).Value = WorksheetFunction.Round(Planilha8.Cells(i, 9).Value, 3)
        wsTempl_ARM.Cells(iULin, 10).Value = wsTempl_ARM.Cells(iULin, 8).Value * wsTempl_ARM.Cells(iULin, 9).Value
        wsTempl_ARM.Cells(iULin, 11).Value = _
        pesquisar_linhaPPU_arm("ARMAZENAMENTO DE ACERVO DOCUMENTAL", wsTempl_ARM.Cells(iULin, 3))
    End If
i = i + 1
Loop

                
'Exportação de dados DE: Painel de controle > Painel, PARA: Template > planilha DADOS
Application.StatusBar = "Procedimento em progresso. Preenchendo planilha DADOS..."
i = 3
Do While Planilha7.Cells(i, 7) <> ""
    sContratoN = Planilha7.Cells(i, 3)
    If sContratoA = sContratoN And sCntrtEstd = Planilha7.Cells(i, 24) Then  '<------------------------
        sSituacao = Planilha7.Cells(i, 2)
        If sSituacao = "CO" Then
            sDescrAtividade = Planilha7.Cells(i, 4)
            If pesquisar_referenciaQExec(sDescrAtividade) = "S" Then
                iULin = wsTempl_DADOS.Range("B5000").End(xlUp).Row + 1
                wsTempl_DADOS.Cells(iULin, 2).Value = Planilha7.Cells(i, 4).Value 'Descrição da atividade
                wsTempl_DADOS.Cells(iULin, 3).Value = Planilha7.Cells(i, 5).Value 'Item
                wsTempl_DADOS.Cells(iULin, 4).Value = Planilha7.Cells(i, 6).Value 'Aplicação
                wsTempl_DADOS.Cells(iULin, 5).Value = Planilha7.Cells(i, 7).Value 'Código da solicitação
                wsTempl_DADOS.Cells(iULin, 6).Value = Planilha7.Cells(i, 9).Value 'D. solicitação
                wsTempl_DADOS.Cells(iULin, 7).Value = Planilha7.Cells(i, 10).Value 'Gerência solicitante
                wsTempl_DADOS.Cells(iULin, 8).Value = Planilha7.Cells(i, 11).Value 'Qtd. (solicitada)
                wsTempl_DADOS.Cells(iULin, 9).Value = Planilha7.Cells(i, 12).Value 'Código OS
                wsTempl_DADOS.Cells(iULin, 10).Value = Planilha7.Cells(i, 13).Value 'D. abertura
                wsTempl_DADOS.Cells(iULin, 11).Value = Planilha7.Cells(i, 14).Value 'D. fechamento
                wsTempl_DADOS.Cells(iULin, 12).Value = Planilha7.Cells(i, 15).Value 'Prazo combinado
                wsTempl_DADOS.Cells(iULin, 13).Value = Planilha7.Cells(i, 16).Value 'Qtd. (Atendida)
                wsTempl_DADOS.Cells(iULin, 14).Value = Planilha7.Cells(i, 17).Value 'Agrupamento
                wsTempl_DADOS.Cells(iULin, 15).Value = Planilha7.Cells(i, 18).Value 'Prazo
                wsTempl_DADOS.Cells(iULin, 16).Value = Planilha7.Cells(i, 19).Value 'FDM
                wsTempl_DADOS.Cells(iULin, 17).Value = Planilha7.Cells(i, 20).Value 'Obs. isenção
                wsTempl_DADOS.Cells(iULin, 18).Value = Planilha7.Cells(i, 21).Value 'FA
                wsTempl_DADOS.Cells(iULin, 19).Value = Planilha7.Cells(i, 22).Value 'Dados PETROBRAS Localidade
                wsTempl_DADOS.Cells(iULin, 20).Value = Planilha7.Cells(i, 23).Value 'Dados PETROBRAS Município
                wsTempl_DADOS.Cells(iULin, 21).Value = Planilha7.Cells(i, 24).Value 'Dados PETROBRAS UF
                wsTempl_DADOS.Cells(iULin, 22).Value = Planilha7.Cells(i, 25).Value 'Dados galpão Localidade
                wsTempl_DADOS.Cells(iULin, 23).Value = Planilha7.Cells(i, 26).Value 'Dados galpão município
                wsTempl_DADOS.Cells(iULin, 24).Value = Planilha7.Cells(i, 27).Value 'Dados galpão UF
                
                'Conforme solicitação da PETROBRAS, a atividade MATERIAL PARA ARQUIVAMENTO entrará para a memória de cálculo, mas não terá nenhum cálculo
                If wsTempl_DADOS.Cells(iULin, 2) = "MATERIAL PARA ARQUIVAMENTO" Then GoTo proximaLinha
                
                'Funcionalidade KM adicional
                If validar_localizacao(wsTempl_DADOS.Cells(iULin, 2)) Then
                    sCentro = pesquisar_centro(wsTempl_DADOS.Cells(iULin, 19))
                    dKM = calcular_KMAdicional(sCentro, wsTempl_DADOS.Cells(iULin, 22))
                    If dKM > 0 Then
                        wsTempl_DADOS.Cells(iULin, 25) = dKM
                    End If
                Else
                    sCentro = pesquisar_centro(wsTempl_DADOS.Cells(iULin, 19))
                End If
                
                'Funcionalidade FC
                
                
                'Funcionalidade QExec
                dQExec = _
                calcular_QExec(wsTempl_DADOS.Cells(iULin, 2), wsTempl_DADOS.Cells(iULin, 3), wsTempl_DADOS.Cells(iULin, 13), wsTempl_DADOS.Cells(iULin, 14))
                'Verifica o fator de ajuste inserido
                If wsTempl_DADOS.Cells(iULin, 18).Value = Empty Then
                    dFA = 1
                Else
                    dFA = wsTempl_DADOS.Cells(iULin, 18).Value
                End If
                wsTempl_DADOS.Cells(iULin, 27) = dQExec * dFA
                
                'Funcionalidade Unidade
                sUnidade = pesquisar_unidade(wsTempl_DADOS.Cells(iULin, 2))
                
                'Funcionalidade Linha PPU
                sLinhaPPU = pesquisar_linhaPPU(wsTempl_DADOS.Cells(iULin, 2))
                
                wsTempl_DADOS.Cells(iULin, 28) = sUnidade
                
                'Funcionalidade Centro
                wsTempl_DADOS.Cells(iULin, 29) = sCentro
                
                'Funcionalidade Código Centro
                'Obs.: Variável sCentro utilizada da pesquisa anterior
                wsTempl_DADOS.Cells(iULin, 30) = pesquisar_codCentro(sCentro)
                
                'Funcionalidade Linha de serviço PPU
                wsTempl_DADOS.Cells(iULin, 31) = sLinhaPPU
                
            End If
        End If
    End If
proximaLinha:
sCentro = ""
i = i + 1
Loop

sDescrAtividade = ""
dQExec = Empty
sLinhaPPU = ""

'Preenchimento referente ao template > planilha MEDICAO
Application.StatusBar = "Procedimento em progresso. Preenchendo planilha MEDICAO..."
i = 3
Do While wsTempl_DADOS.Cells(i, 2) <> ""
    sDescrAtividade = wsTempl_DADOS.Cells(i, 2)
    If sDescrAtividade = "MATERIAL PARA ARQUIVAMENTO" Then GoTo proximaLinha2
    sEstado = wsTempl_DADOS.Cells(i, 21)
    dQExec = wsTempl_DADOS.Cells(i, 27)
    sCentro = wsTempl_DADOS.Cells(i, 29)
    sCodcentro = wsTempl_DADOS.Cells(i, 30)
    sLinhaPPU = wsTempl_DADOS.Cells(i, 31)
    sDado1 = sEstado & "#" & sCentro & "#" & sDescrAtividade
    
    iCont = WorksheetFunction.CountIf(wsTempl_MEDICAO.Range("L:L"), sDado1)
    
    If iCont = 0 Then 'Preenchendo nova linha no caso de não encontrado
        iULin = wsTempl_MEDICAO.Range("D5000").End(xlUp).Row + 1
        wsTempl_MEDICAO.Cells(iULin, 2) = sEstado 'Estado
        wsTempl_MEDICAO.Cells(iULin, 3) = sCentro & "#" & sCodcentro 'Centro#Cód. centro
        wsTempl_MEDICAO.Cells(iULin, 4) = sDescrAtividade 'Descrição da atividade
        wsTempl_MEDICAO.Cells(iULin, 5) = pesquisar_itemPPU(sLinhaPPU) 'Item PPU (Pesquisa na tabelaA)
        wsTempl_MEDICAO.Cells(iULin, 6) = sLinhaPPU 'Linha de serviço PPU
        wsTempl_MEDICAO.Cells(iULin, 7) = dQExec 'QExec
        iRefFDM = pesquisar_referenciaFDM(sDescrAtividade) 'FDM (Pesquisa na tabelaA)
        
        If sDescrAtividade = "TRANSFERENCIA DE ACERVO" Then
            wsTempl_MEDICAO.Cells(iULin, 8) = 1
        Else
            If iRefFDM = 1 Then
                wsTempl_MEDICAO.Cells(iULin, 8) = dFDM1
            ElseIf iRefFDM = 2 Then
                wsTempl_MEDICAO.Cells(iULin, 8) = dFDM2
            End If
        End If
        wsTempl_MEDICAO.Cells(iULin, 9).FormulaLocal = "=G" & iULin & "*H" & iULin 'QFinal
        wsTempl_MEDICAO.Cells(iULin, 10) = pesquisar_precoUnitario(sContratoA, sLinhaPPU) 'preço unitário  (Pesquisa na tabelaB)
        wsTempl_MEDICAO.Cells(iULin, 11).FormulaLocal = "=I" & iULin & "*J" & iULin 'Subtotal
        wsTempl_MEDICAO.Cells(iULin, 12) = sDado1 'dados concatenados (Referência de consulta)
    Else
    
        bEncontrou = False
        n = 4
        Do While wsTempl_MEDICAO.Cells(n, 12) <> "" And Not bEncontrou
        sDado2 = wsTempl_MEDICAO.Cells(n, 12)
        
            If sDado1 = sDado2 Then
            wsTempl_MEDICAO.Cells(n, 7) = wsTempl_MEDICAO.Cells(n, 7) + dQExec
            bEncontrou = True
            End If
          
        n = n + 1
        Loop

    End If
proximaLinha2:
i = i + 1
Loop

'Limpar dados não necessários para a memória de cálculo . planilha MEDICAO
wsTempl_MEDICAO.Range("L:L").ClearContents

'Funcionalidade para salvar arquivo  '<------------------------
Application.StatusBar = "Procedimento em progresso. Salvando Arquivo Como..."
sNomeArquivo = _
wsTempl_MC.Range("D5") & "-PLA-MemoriaPetrobras-" & iPerAno & Format(iPerMes, "00") & "-" & wsTempl_MC.Range("J5") & _
"#" & Year(Now) & "-" & Month(Now) & "-" & Day(Now) & "_" & Hour(Now) & "h" & Minute(Now) & "min"

vSalvarComo = Application.GetSaveAsFilename(InitialFileName:=sNomeArquivo, filefilter:="Excel Files (*.xlsx),*.xlsx")
If vSalvarComo <> False Then
    ActiveWorkbook.SaveAs filename:=vSalvarComo
End If

wbTemplate.Close (False)

Set wbTemplate = Nothing
Set wsTempl_MC = Nothing
Set wsTempl_ARM = Nothing
Set wsTempl_DADOS = Nothing
Set wsTempl_MEDICAO = Nothing

Application.StatusBar = False

End Sub


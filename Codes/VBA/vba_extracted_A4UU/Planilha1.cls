Attribute VB_Name = "Planilha1"

Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = False

Attribute VB_PredeclaredId = True

Attribute VB_Exposed = True

Attribute VB_TemplateDerived = False

Attribute VB_Customizable = True

'Procedimentos desenvolvidos por: Thiago Lima de Paiva - XMED



Option Explicit



Private Const MONITORAMENTO1 = "Monitoramento"



Sub analiseMonitoramento()

'Procedimento para análise de dados inseridos na plamilha



'VARIÁVEIS

Dim i As Integer

Dim sMensagem As String

Dim sPerMes As String

Dim iPerMes As Integer

Dim iPerAno As Integer

Dim iPerAnt As Integer

Dim ul As Integer

Dim dtData As Date

Dim iDia As Integer

Dim iMes As Integer

Dim iAno As Integer

Dim iCont As Integer



Dim sValorA As String

Dim sValorB() As String

Dim sMunicipio As String

Dim sUF As String

Dim dtSLA As Date

Dim dtDataSLA As Date

Dim dtHoraSLA As Date

Dim vTempo As Variant

Dim j As Integer

Dim dtReSLA As Date

Dim n As Integer



'FUNCIONALIDADE PARA PERÍODO

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

sPerMes = UCase(Mid(Planilha3.Range("C12").Value, 4, 20)) 'Descrição Mês - Período informado na planilha Menu

iPerMes = Left(Planilha3.Range("C12").Value, 2) 'Número Mês - Período informado na planilha Menu

iPerAno = Planilha3.Range("C13").Value 'Ano - Período informado na planilha Menu



sMensagem = MsgBox("Dados a serem analisados." & Chr(10) & "Período: " & sPerMes & "/" & iPerAno & _

Chr(10) & "Situação: CO - Concluído e AT - Atendido" & Chr(10) & Chr(10) & "INICIAR?", vbYesNo, MONITORAMENTO1)

If sMensagem = vbNo Then

    MsgBox "Análise cancelada.", vbInformation, MONITORAMENTO1

    Exit Sub

End If



iPerAnt = iPerMes - 1

If iPerAnt = 0 Then iPerAnt = 12

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



Planilha1.Range("Z3:Z20000").ClearContents



On Error Resume Next



'Preenchimento e análise de dados concluídos e atendidos

ul = Planilha1.Range("I20000").End(xlUp).Row

For i = 3 To ul

On Error GoTo -1

'##############################################################################################

'ANALISA DADOS DA SITUAÇÃO CONCLUÍDA E DATA DE FECHAMENTO NO PERÍODO

'Obs.: Entrando nos critérios, inicia a análise e preenchimento





'Coluna C

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Planilha1.Cells(i, 3) = "" Then

    GoTo proximaLinha

    

ElseIf Planilha1.Cells(i, 3) <> "CO" And Planilha1.Cells(i, 3) <> "AT" Then

    If WorksheetFunction.CountIf(Planilha2.Range("B:B"), Planilha1.Cells(i, 3)) = 0 Then

        Planilha1.Cells(i, 26) = _

        Planilha1.Cells(i, 26) & "(Coluna C:) " & ExibirMsg("descrInvalida") & "; "

    End If

    GoTo proximaLinha

    

End If







'Coluna B

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Planilha1.Cells(i, 2) <> "" Then 'Célula vazia



    iCont = Len(Planilha1.Cells(i, 2)) - Len(Trim(Planilha1.Cells(i, 2)))

    If iCont > 0 Then

        Planilha1.Cells(i, 26) = _

        Planilha1.Cells(i, 26) & "(Coluna B): " & ExibirMsg("dataInvalida") & "; "

        GoTo proximaLinha

    End If

    

    dtData = Planilha1.Cells(i, 2)

    If Err.Number = 13 Then

        Planilha1.Cells(i, 26) = _

        Planilha1.Cells(i, 26) & "(Coluna B): " & ExibirMsg("dataInvalida") & "; "

        GoTo proximaLinha

    

    ElseIf IsDate(Planilha1.Cells(i, 2)) Then 'Verifica é data

        iDia = Day(dtData)

        iMes = Month(dtData)

        iAno = Year(dtData)

    

        'Verifica se a data está no período

        If iPerMes = iMes And iDia < 26 And iAno = iPerAno Then

        ElseIf iPerAnt = iMes And iDia > 25 And iDia < 32 And iPerAno = iAno Then

        ElseIf iPerAnt = iMes And iDia > 25 And iDia < 32 And iAno = iPerAno - 1 Then

        Else

            GoTo proximaLinha

        End If

    End If

Else

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna B): " & ExibirMsg("campoVazioMed") & "; "

    GoTo proximaLinha

End If

iCont = Empty

Err.Clear







'ANÁLISE E PREENCHIMENTO

'##############################################################################################



'Coluna A

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Planilha1.Cells(i, 1) = "" Then Planilha1.Cells(i, 1) = StrConv(preencher_usuario_rede(), vbUpperCase)

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



    

    

'Coluna B

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'Funcionalidade na parte de conferência de período

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    

   





'Coluna C

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'Funcionalidade na parte de conferência de período

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



  

     



'Coluna D

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Planilha1.Cells(i, 4) <> "" Then

    If WorksheetFunction.CountIf(Planilha2.Range("E:E"), Planilha1.Cells(i, 4)) = 0 Then

        Planilha1.Cells(i, 26) = _

        Planilha1.Cells(i, 26) & "(Coluna D:) " & ExibirMsg("descrInvalida") & "; "

    End If

Else

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna D): " & ExibirMsg("campoVazio") & "; "

End If

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////







'Coluna E

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Planilha1.Cells(i, 5) <> "" Then

    If WorksheetFunction.CountIf(Planilha2.Range("AD:AD"), Planilha1.Cells(i, 5)) = 0 Then

        Planilha1.Cells(i, 26) = _

        Planilha1.Cells(i, 26) & "(Coluna E:) " & ExibirMsg("descrInvalida") & "; "

    End If

Else

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna D): " & ExibirMsg("campoVazio") & "; "

End If



'Colunas D e E

If Not validar_contratoUF(Planilha1.Cells(i, 4), Planilha1.Cells(i, 5)) Then

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna D e E): " & ExibirMsg("contratoXuf") & "; "

End If

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////









'Coluna F

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Planilha1.Cells(i, 6) <> "" Then

    If WorksheetFunction.CountIf(Planilha2.Range("K:K"), Planilha1.Cells(i, 6)) = 0 Then

        Planilha1.Cells(i, 26) = _

        Planilha1.Cells(i, 26) & "(Coluna F): " & ExibirMsg("descrInvalida") & "; "

    End If

Else

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna F): " & ExibirMsg("campoVazio") & "; "

    GoTo pulaColFG

End If

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////









'Coluna G

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Planilha1.Cells(i, 7) <> "" Then

    If WorksheetFunction.CountIf(Planilha2.Range("T:T"), Planilha1.Cells(i, 7)) = 0 Then

        Planilha1.Cells(i, 26) = _

        Planilha1.Cells(i, 26) & "(Coluna G:) " & ExibirMsg("descrInvalida") & "; "

    End If

Else

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna G): " & ExibirMsg("campoVazio") & "; "

    GoTo pulaColFG

End If

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////







'Coluna F e G

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Not validar_itemDescricao(Planilha1.Cells(i, 6), Planilha1.Cells(i, 7)) Then

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna F e G): " & ExibirMsg("itemXAtividade") & "; "

End If

pulaColFG:

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////







'Coluna H

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Planilha1.Cells(i, 8) <> "" Then

    If WorksheetFunction.CountIf(Planilha2.Range("AH:AH"), Planilha1.Cells(i, 8)) = 0 Then

        Planilha1.Cells(i, 26) = _

        Planilha1.Cells(i, 26) & "(Coluna H:) " & ExibirMsg("descrInvalida") & "; "

    End If

Else

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna H): " & ExibirMsg("campoVazio") & "; "

End If

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////







'Coluna I

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Planilha1.Cells(i, 9) = "" Then

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna I): " & ExibirMsg("campoVazio") & "; "

End If



If Len(Planilha1.Cells(i, 9)) - Len(WorksheetFunction.Substitute(Planilha1.Cells(i, 9), " ", "")) > 0 Then

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna I): " & ExibirMsg("caractInvalido") & "Espaço; "

End If

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////







'Coluna J

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Planilha1.Cells(i, 10) = "" Then

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna J): " & ExibirMsg("campoVazio") & "; "



ElseIf Len(Planilha1.Cells(i, 10)) - Len(WorksheetFunction.Substitute(Planilha1.Cells(i, 10), " ", "")) > 0 Then

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna J): " & ExibirMsg("caractInvalido") & "Espaço; "

End If

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////







'Coluna K

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Planilha1.Cells(i, 11) = "" Then

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna K): " & ExibirMsg("campoVazio") & "; "

End If

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////







'Coluna L

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Planilha1.Cells(i, 12) <> "" Then 'Célula vazia



    iCont = Len(Planilha1.Cells(i, 12)) - Len(Trim(Planilha1.Cells(i, 12)))

    If iCont > 0 Then

        Planilha1.Cells(i, 26) = _

        Planilha1.Cells(i, 26) & "(Coluna L): " & ExibirMsg("dataInvalida") & "; "



    End If

    

    dtData = Planilha1.Cells(i, 12)

    If Err.Number = 13 Then

        Planilha1.Cells(i, 26) = _

        Planilha1.Cells(i, 26) & "(Coluna L): " & ExibirMsg("dataInvalida") & "; "

    

    End If

Else

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna L): " & ExibirMsg("campoVazio") & "; "



End If

iCont = Empty

dtData = Empty

Err.Clear

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////









'Coluna M

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Planilha1.Cells(i, 13) = Empty Then

    If Planilha1.Cells(i, 12) = Empty Then

        Planilha1.Cells(i, 26) = _

        Planilha1.Cells(i, 26) & "(Coluna L): " & ExibirMsg("campoVazio") & "; "

        

    Else

        

        sValorA = pesquisar_centro(Planilha1.Cells(i, 22))

        If sValorA = "Não encontrado" Then

            Planilha1.Cells(i, 26) = _

            Planilha1.Cells(i, 26) & "(Coluna M): " & ExibirMsg("infNaoEncontrada") & ": Centro; "

            

        End If

        

        sValorB() = Split(pesquisar_municipioUF(sValorA), "|")

        sMunicipio = sValorB(0)

        sUF = sValorB(1)

        

        dtSLA = Planilha1.Cells(i, 12)

        dtDataSLA = Format(dtSLA, "dd/mm/yyyy")

        dtHoraSLA = Format(dtSLA, "hh:mm")

                

        vTempo = "8:00"

        n = Planilha3.Range("G13")

        For j = 1 To n

            dtReSLA = calcular_dataSLA(dtDataSLA, dtHoraSLA, vTempo, sMunicipio, sUF)

            dtDataSLA = Format(dtReSLA, "dd/mm/yyyy")

            dtHoraSLA = Format(dtReSLA, "hh:mm")

        Next j

        

        Planilha1.Cells(i, 13) = dtDataSLA + dtHoraSLA

        

    End If



ElseIf Planilha1.Cells(i, 13) <> Empty Then 'Célula vazia

    If Not IsDate(Planilha1.Cells(i, 13)) Then 'Verifica é data

        Planilha1.Cells(i, 26) = _

        Planilha1.Cells(i, 26) & "(Coluna M): " & ExibirMsg("dataInvalida") & "; "

        

    End If

End If



sValorA = ""

sMunicipio = ""

sUF = ""

dtSLA = Empty

dtDataSLA = Empty

dtHoraSLA = Empty

dtReSLA = Empty

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////









'Coluna N

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Planilha1.Cells(i, 14) = "" Then

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna N): " & ExibirMsg("CampoVazio") & "; "

    

ElseIf Len(Planilha1.Cells(i, 14)) - Len(WorksheetFunction.Substitute(Planilha1.Cells(i, 14), " ", "")) > 0 Then

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna N): " & ExibirMsg("caractInvalido") & "Espaço; "

    

ElseIf Len(Planilha1.Cells(i, 14)) - Len(WorksheetFunction.Substitute(Planilha1.Cells(i, 14), "_", "")) > 0 Then

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna N): " & ExibirMsg("caractInvalido") & "Underline; "

    

ElseIf Len(Planilha1.Cells(i, 14)) - Len(WorksheetFunction.Substitute(Planilha1.Cells(i, 14), "\", "")) > 0 Then

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna N): " & ExibirMsg("caractInvalido") & "Barra invertida; "

    

End If

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////







'Coluna O

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Planilha1.Cells(i, 15) = 0 Or Planilha1.Cells(i, 15) = Empty Then 'Célula vazia

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna O): " & ExibirMsg("campoVazio") & "; "

    

End If

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////







'Coluna P

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Planilha1.Cells(i, 16) = "" Then

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna P): " & ExibirMsg("campoVazio") & "; "

    

ElseIf Len(Planilha1.Cells(i, 16)) - Len(WorksheetFunction.Substitute(Planilha1.Cells(i, 16), " ", "")) > 0 Then

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna P): " & ExibirMsg("caractInvalido") & "Espaço; "

    

End If

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////







'Coluna Q

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Planilha1.Cells(i, 17) <> "" Then 'Célula vazia



    iCont = Len(Planilha1.Cells(i, 17)) - Len(Trim(Planilha1.Cells(i, 17)))

    If iCont > 0 Then

        Planilha1.Cells(i, 26) = _

        Planilha1.Cells(i, 26) & "(Coluna Q): " & ExibirMsg("dataInvalida") & "; "

        GoTo proximaLinha

    End If

    

    dtData = Planilha1.Cells(i, 17)

    If Err.Number = 13 Then

        Planilha1.Cells(i, 26) = _

        Planilha1.Cells(i, 26) & "(Coluna Q): " & ExibirMsg("dataInvalida") & "; "

        GoTo proximaLinha

    

    End If

Else

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna Q): " & ExibirMsg("campoVazio") & "; "

    GoTo proximaLinha

End If

iCont = Empty

dtData = Empty

Err.Clear

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////







'Coluna R

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Planilha1.Cells(i, 18) <> "" Then 'Célula vazia



    iCont = Len(Planilha1.Cells(i, 18)) - Len(Trim(Planilha1.Cells(i, 18)))

    If iCont > 0 Then

        Planilha1.Cells(i, 26) = _

        Planilha1.Cells(i, 26) & "(Coluna R): " & ExibirMsg("dataInvalida") & "; "

        

    End If

    

    dtData = Planilha1.Cells(i, 18)

    If Err.Number = 13 Then

        Planilha1.Cells(i, 26) = _

        Planilha1.Cells(i, 26) & "(Coluna R): " & ExibirMsg("dataInvalida") & "; "

    

    End If

Else

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna R): " & ExibirMsg("campoVazioMed") & "; "



End If

iCont = Empty

dtData = Empty

Err.Clear

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////







'Coluna S

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Planilha1.Cells(i, 19) = Empty Then

    If pesquisar_referenciaPrazo(Planilha1.Cells(i, 5)) = "Fiscalização" Then

        Planilha1.Cells(i, 26) = _

        Planilha1.Cells(i, 26) & "(Coluna S): " & ExibirMsg("atividadeSemPrazo") & "; "

    End If

End If

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////







'Coluna T

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Planilha1.Cells(i, 20) = 0 Or Planilha1.Cells(i, 20) = Empty Then 'Célula vazia

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna T): " & ExibirMsg("campoVazio") & "; "

End If

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////







'Coluna U

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Planilha1.Cells(i, 21) <> "" Then 'Célula vazia

    If WorksheetFunction.CountIf(Planilha2.Range("Y:Y"), Planilha1.Cells(i, 21)) = 0 Then

        Planilha1.Cells(i, 26) = _

        Planilha1.Cells(i, 26) & "(Coluna U): " & ExibirMsg("descrInvalida") & "; "

    End If

End If

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////







'Coluna V

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Planilha1.Cells(i, 22) = "" Then 'Célula vazia

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna V): " & ExibirMsg("campoVazio") & "; "

Else

    Planilha1.Cells(i, 22) = Trim(WorksheetFunction.Substitute(Planilha1.Cells(i, 22), "  ", " "))

    If WorksheetFunction.CountIf(Planilha2.Range("AK:AK"), Planilha1.Cells(i, 22)) = 0 Then

        Planilha1.Cells(i, 26) = _

        Planilha1.Cells(i, 26) & "(Coluna V): " & ExibirMsg("infNaoEncontrada") & ": Localidade; "

    Else

        If verificar_localidadeAtiva(Planilha1.Cells(i, 22)) = "N" Then

            Planilha1.Cells(i, 26) = _

            Planilha1.Cells(i, 26) & "(Coluna V): " & ExibirMsg("ativarLocalidade") & "; "

        End If

    End If

End If

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////







'Coluna W a X

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'Campo aberto

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////









'Coluna Y

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

If Planilha1.Cells(i, 25) = "" Then

    Planilha1.Cells(i, 26) = _

    Planilha1.Cells(i, 26) & "(Coluna Y): " & ExibirMsg("campoVazio") & "; "

Else

    If UCase(Planilha1.Cells(i, 25)) <> "SIM" And UCase(Planilha1.Cells(i, 25)) <> "NÃO" Then

        Planilha1.Cells(i, 26) = _

        Planilha1.Cells(i, 26) & "(Coluna W): " & ExibirMsg("limiteDesr") & "; "

    Else

        Planilha1.Cells(i, 25) = UCase(Planilha1.Cells(i, 25))

    End If

End If

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



proximaLinha:

Next i



iCont = Empty

iCont = WorksheetFunction.CountA(Planilha1.Range("Z3:Z20003"))

If iCont = 0 Then

    MsgBox "Procedimento realizado com sucesso. Sem pendências no campo LOG", vbInformation, MONITORAMENTO1

    

Else

    MsgBox "Procedimento realizado. Pendências no campo LOG", vbExclamation, MONITORAMENTO1

    

End If



End Sub



Sub faturamentoPorPeriodo()

'Cálcula quantidade de solicitações atendidas no período



Dim i As Integer

Dim sMensagem As String

Dim sPerMes As String

Dim iPerMes As Integer

Dim iPerAno As Integer

Dim iPerAnt As Integer

Dim ul As Integer

Dim dtData As Date

Dim iDia As Integer

Dim iMes As Integer

Dim iAno As Integer

Dim iCont As Integer



'FUNCIONALIDADE PARA PERÍODO

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

sPerMes = UCase(Mid(Planilha3.Range("C12").Value, 4, 20)) 'Descrição Mês - Período informado na planilha Menu

iPerMes = Left(Planilha3.Range("C12").Value, 2) 'Número Mês - Período informado na planilha Menu

iPerAno = Planilha3.Range("C13").Value 'Ano - Período informado na planilha Menu



sMensagem = MsgBox("Dados do período de faturamento serão contabilizados." & Chr(10) & "Período: " & sPerMes & "/" & iPerAno & _

Chr(10) & "INICIAR?", vbYesNo, MONITORAMENTO1)

If sMensagem = vbNo Then

    MsgBox "Análise cancelada.", vbInformation, MONITORAMENTO1

    Exit Sub

End If



iPerAnt = iPerMes - 1

If iPerAnt = 0 Then iPerAnt = 12

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



ul = Planilha1.Range("A20000").End(xlUp).Row

For i = 3 To ul



If Planilha1.Cells(i, 2) <> Empty Then 'Célula vazia

    If IsDate(Planilha1.Cells(i, 2)) Then 'Verifica é data

        dtData = Planilha1.Cells(i, 2)

        iDia = Day(dtData)

        iMes = Month(dtData)

        iAno = Year(dtData)

    

        'Verifica se a data está no período

        If iPerMes = iMes And iDia < 26 And iAno = iPerAno Then

        ElseIf iPerAnt = iMes And iDia > 25 And iDia < 32 And iPerAno = iAno Then

        ElseIf iPerAnt = iMes And iDia > 25 And iDia < 32 And iAno = iPerAno - 1 Then

        Else

            GoTo pulaProcesso

        End If

    End If

    

End If

iCont = iCont + 1



pulaProcesso:

Next i



MsgBox "Segue quantidade contabilizada: " & iCont, vbInformation, MONITORAMENTO1



End Sub


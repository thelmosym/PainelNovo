Attribute VB_Name = "subDadosSaidaExportacao"
'================================================================================================= ===================
'Módulo procedimento para dados de saída
'Obs.: Todos os dados de saída desenvolvidos para guarda externa
'================================================================================================= ===================
Option Explicit

Sub exportacaoDadosRL01()
'Procedimento

Dim iULin As Integer
Dim iULinT As Integer
Dim i As Integer
Dim sDescrAtividade As String
Dim iNum As Integer
Dim sMunicipio As String
Dim sContrato As String
Dim sUF As String
Dim sGrupo As String
Dim dPrazo As Double
Dim dtSLA As Date
Dim dtDataSLA As Date
Dim dtHoraSLA As Date
Dim dHora As Double
Dim dtDataX As Date
Dim j As Integer
Dim iCont As Integer
Dim dirTemplate As String
Dim wbTemplate As Workbook
Dim wsTempl_dd As Worksheet
Dim wsTempl_ia As Worksheet
Dim sMensagem As String
Dim sDado1 As String
Dim sDado2 As String
Dim sRefPrazo As String
Dim dtDAbertura As Date
Dim dtDFechamento As Date
Dim vTempo As Variant
Dim dtReSLA As Date
Dim dtPrazoCombinado  As Date
Dim iCont1 As Integer
Dim bEncontrou As Boolean
Dim sDadoN As String
Dim iCont2 As Integer
Dim sNomeArquivo As String
Dim vSalvarComo As Variant


sMensagem = MsgBox("I M P O R T A N T E:" & _
Chr(10) & _
Chr(10) & _
"Nao gerar este relatorio antes de passar por todos os procedimentos de FDM e MEMORIA DE CALCULO (Ver sequencia 1, 2, 3 e 4)." & _
Chr(10) & _
"Obs.: Estes dados nao passarao por atualizacao e analise. Certifique-se de todos os procedimentos antes de gerar este relatorio." & Chr(10) & Chr(10) & _
"Deseja continuar?", vbCritical + vbYesNo, MSG_PN_EXP0001)

'Funcionalidade para salvar arquivo
Application.ScreenUpdating = True
Application.StatusBar = "Procedimento em progresso. Preenchendo dados do painel ..."
Application.ScreenUpdating = False

If sMensagem = vbNo Then
    MsgBox "Preenchimento cancelado.", vbInformation, MSG_PN_EXP0001
    Exit Sub
End If

dirTemplate = Planilha5.Range("C59")

Set wbTemplate = Workbooks.Open(dirTemplate) 'Abre template
Set wsTempl_dd = wbTemplate.Worksheets(1)
Set wsTempl_ia = wbTemplate.Worksheets(2)

'Executa laço de repetição com base nos dados solicitados
iULin = Planilha7.Range("G2003").End(xlUp).Row
For i = 3 To iULin
    
        iULinT = wsTempl_dd.Range("C2003").End(xlUp).Row + 1
        wsTempl_dd.Cells(iULinT, 1) = Planilha7.Cells(i, 3)
        wsTempl_dd.Cells(iULinT, 2) = Planilha7.Cells(i, 4)
        wsTempl_dd.Cells(iULinT, 3) = Planilha7.Cells(i, 5)
        wsTempl_dd.Cells(iULinT, 4) = Planilha7.Cells(i, 7)
        wsTempl_dd.Cells(iULinT, 5) = Planilha7.Cells(i, 9)
        wsTempl_dd.Cells(iULinT, 6) = Planilha7.Cells(i, 10)
        wsTempl_dd.Cells(iULinT, 7) = Planilha7.Cells(i, 12)
        wsTempl_dd.Cells(iULinT, 8) = Planilha7.Cells(i, 13)
        wsTempl_dd.Cells(iULinT, 9) = Planilha7.Cells(i, 14)
        wsTempl_dd.Cells(iULinT, 10) = Planilha7.Cells(i, 15)
        wsTempl_dd.Cells(iULinT, 11) = Planilha7.Cells(i, 16)
        wsTempl_dd.Cells(iULinT, 12) = Planilha7.Cells(i, 22)
        wsTempl_dd.Cells(iULinT, 13) = Planilha7.Cells(i, 23)
        wsTempl_dd.Cells(iULinT, 14) = Planilha7.Cells(i, 24)
        wsTempl_dd.Cells(iULinT, 15) = Planilha7.Cells(i, 25)
        wsTempl_dd.Cells(iULinT, 16) = Planilha7.Cells(i, 26)
        wsTempl_dd.Cells(iULinT, 17) = Planilha7.Cells(i, 27)
        wsTempl_dd.Cells(iULinT, 18) = Planilha7.Cells(i, 18)
        If Planilha7.Cells(i, 19) = "NP" Then
            wsTempl_dd.Cells(iULinT, 19) = "No prazo"
        ElseIf Planilha7.Cells(i, 19) = "FP" Then
            wsTempl_dd.Cells(iULinT, 19) = "Fora do prazo"
        Else
            wsTempl_dd.Cells(iULinT, 19) = ""
        End If
        wsTempl_dd.Cells(iULinT, 20) = Planilha7.Cells(i, 20)
        
        sDescrAtividade = Planilha7.Cells(i, 4) 'Atividade (Serviço executado)
        'Verificar tipo da atividade
        iNum = pesquisar_referenciaFDM(sDescrAtividade)
        wsTempl_dd.Cells(iULinT, 21) = iNum
        'Para atividades dentro do escopo de mediçãode FDM, a classificalão será "Normal" ou "Expresso"
        If iNum > 0 Then
            'Busca pela referência de prazo
            sRefPrazo = pesquisar_referenciaPrazo(sDescrAtividade)
            wsTempl_dd.Cells(iULinT, 22) = sRefPrazo
            'Tabela: são prazos pré-definidos conforme contrato
            'Localizados na planilha tabelaA
            If sRefPrazo = "Tabela" Then
                sContrato = Planilha7.Cells(i, 3)
                dtDAbertura = Planilha7.Cells(i, 13)
                dtDFechamento = Planilha7.Cells(i, 14)
                sMunicipio = Planilha7.Cells(i, 23)
                sUF = Planilha7.Cells(i, 24)
                
                'Executa SLA
                'Índice que aponta o tempo de atendimento
                sGrupo = pesquisar_grupoMunicipio(sMunicipio, sUF)
                wsTempl_dd.Cells(iULinT, 23) = sGrupo

                dPrazo = pesquisar_prazo(sContrato, sDescrAtividade, sGrupo)
                wsTempl_dd.Cells(iULinT, 24) = dPrazo
                dtSLA = dtDAbertura
                dtDataSLA = Format(dtSLA, "dd/mm/yyyy")
                dtHoraSLA = Format(dtSLA, "hh:mm")
                
                'Os prazos encontram-se na planilha tabelaA
                Select Case dPrazo
                    'Prazo com zero não será aceito determinado serviço para...
                    '...grupo de município ou localidade
                    Case 0
                        wsTempl_dd.Cells(iULinT, 26) = "Sem prazo para este grupo (Ver contrato)"
                        GoTo proximaLinha:
                    'Prazo 1/2
                    Case 0.5
                        vTempo = "4:00"
                        dtReSLA = calcular_dataSLA(dtDataSLA, dtHoraSLA, vTempo, sMunicipio, sUF)
                        
                        If Format(dtDAbertura, "hh:mm") >= "08:00" And Format(dtDAbertura, "hh:mm") <= "11:00" Then
                            dHora = CDbl(CDate("17:00"))
                            dtDataX = Format(dtReSLA, "dd/mm/yyyy")
                            dtDataX = dtDataX + dHora
                            dtReSLA = dtDataX
                            wsTempl_dd.Cells(iULinT, 25) = dtReSLA
                            
                        ElseIf Format(dtDAbertura, "hh:mm") > "11:00" And Format(dtDAbertura, "hh:mm") <= "16:00" Then
                            dHora = CDbl(CDate("12:00"))
                            dtDataX = Format(dtReSLA, "dd/mm/yyyy")
                            dtDataX = dtDataX + dHora
                            dtReSLA = dtDataX
                            wsTempl_dd.Cells(iULinT, 25) = dtReSLA
                            
                        Else
                            dHora = CDbl(CDate("17:00"))
                            dtDataX = Format(dtReSLA, "dd/mm/yyyy")
                            dtDataX = dtDataX + dHora
                            dtReSLA = dtDataX
                            wsTempl_dd.Cells(iULinT, 25) = dtReSLA
                            
                        End If
                         
                    Case 1
                        vTempo = "8:00"
                        dtReSLA = calcular_dataSLA(dtDataSLA, dtHoraSLA, vTempo, sMunicipio, sUF)
                        wsTempl_dd.Cells(iULinT, 25) = dtReSLA
                        
                    Case 2
                        vTempo = "16:00"
                        dtReSLA = calcular_dataSLA(dtDataSLA, dtHoraSLA, vTempo, sMunicipio, sUF)
                        wsTempl_dd.Cells(iULinT, 25) = dtReSLA
                        
                    Case 3
                        vTempo = "8:00"
                        For j = 1 To 3
                            dtReSLA = calcular_dataSLA(dtDataSLA, dtHoraSLA, vTempo, sMunicipio, sUF)
                            dtDataSLA = Format(dtReSLA, "dd/mm/yyyy")
                            dtHoraSLA = Format(dtReSLA, "hh:mm")
                        Next j
                        wsTempl_dd.Cells(iULinT, 25) = dtReSLA
                        
                    Case 4
                        vTempo = "8:00"
                        For j = 1 To 4
                            dtReSLA = calcular_dataSLA(dtDataSLA, dtHoraSLA, vTempo, sMunicipio, sUF)
                            dtDataSLA = Format(dtReSLA, "dd/mm/yyyy")
                            dtHoraSLA = Format(dtReSLA, "hh:mm")
                        Next j
                        wsTempl_dd.Cells(iULinT, 25) = dtReSLA
                        
                    Case 5
                        vTempo = "8:00"
                        For j = 1 To 5
                            dtReSLA = calcular_dataSLA(dtDataSLA, dtHoraSLA, vTempo, sMunicipio, sUF)
                            dtDataSLA = Format(dtReSLA, "dd/mm/yyyy")
                            dtHoraSLA = Format(dtReSLA, "hh:mm")
                        Next j
                        wsTempl_dd.Cells(iULinT, 25) = "Sem prazo para este grupo (Ver contrato)"
                        
                End Select
                
                'Resultado SLA dentro do prazo
                If dtReSLA >= dtDFechamento Then
                    wsTempl_dd.Cells(iULinT, 26) = "Resultado dentro do prazo (Prazo calculado maior/igual data fechamento)"

                'Resultado SLA fora do prazo
                ElseIf dtReSLA < dtDFechamento Then
                    wsTempl_dd.Cells(iULinT, 26) = "Resultado fora do prazo (Prazo calculado menor/igual data fechamento)"
                    
                End If

             ElseIf sRefPrazo = "Fiscalização" Then
                dtDFechamento = Planilha7.Cells(i, 14)
                dtPrazoCombinado = Planilha7.Cells(i, 15)

                If dtPrazoCombinado >= dtDFechamento Then
                    wsTempl_dd.Cells(iULinT, 26) = "Resultado dentro do prazo (Prazo combinado maior/igual data fechamento)"

                Else
                    wsTempl_dd.Cells(iULinT, 26) = "Resultado fora do prazo (Prazo combinado menor/igual data fechamento)"

                End If
                    
            End If
        Else
            wsTempl_dd.Cells(iULinT, 21) = iNum
            wsTempl_dd.Cells(iULinT, 26) = "Esta atividade nao possui prazo de calculo a ser realizado"
        End If
    
         'If wsTempl_dd.Cells(iULinT, 2) = "MATERIAL PARA ARQUIVAMENTO" Then GoTo proximaLinha
    
    
    
proximaLinha:
sDescrAtividade = ""
iNum = Empty
sRefPrazo = ""
dtDAbertura = Empty
dtDFechamento = Empty
sMunicipio = ""
dtPrazoCombinado = Empty
sContrato = ""
sUF = ""
sGrupo = ""
dPrazo = Empty
dtSLA = Empty
dtDataSLA = Empty
dtHoraSLA = Empty
dtReSLA = Empty
dHora = Empty
dtDataX = Empty
Next i

Application.ScreenUpdating = True
Application.StatusBar = "Procedimento em progresso. Preenchendo dados do FDM e IAPARQ ..."
Application.ScreenUpdating = False


i = 4
Do While wsTempl_dd.Cells(i, 2) <> ""
    sDado1 = wsTempl_dd.Cells(i, 1) & wsTempl_dd.Cells(i, 14) & wsTempl_dd.Cells(i, 19)
    sDado2 = wsTempl_dd.Cells(i, 1) & wsTempl_dd.Cells(i, 19)
    
    If InStr(sDado1, "No prazo") = 0 And InStr(sDado2, "No prazo") = 0 And InStr(sDado2, "Fora do prazo") = 0 And InStr(sDado2, "Fora do prazo") = 0 Then
        GoTo pulaLinha:
    End If
    
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    iCont1 = WorksheetFunction.CountIf(wsTempl_ia.Range("F:F"), sDado1)
    If iCont1 = 0 Then
        iULin = wsTempl_ia.Range("B100").End(xlUp).Row + 1
        wsTempl_ia.Cells(iULin, 2) = wsTempl_dd.Cells(i, 1)
        wsTempl_ia.Cells(iULin, 3) = wsTempl_dd.Cells(i, 14)
        wsTempl_ia.Cells(iULin, 4) = wsTempl_dd.Cells(i, 19)
        wsTempl_ia.Cells(iULin, 5) = 1
        wsTempl_ia.Cells(iULin, 6) = sDado1
        
    Else
        j = 3
        bEncontrou = False
        Do While wsTempl_ia.Cells(j, 2) <> "" And Not bEncontrou
            sDadoN = wsTempl_ia.Cells(j, 6)
            If sDadoN = sDado1 Then
                wsTempl_ia.Cells(j, 5) = wsTempl_ia.Cells(j, 5) + 1
                bEncontrou = True
            End If
        j = j + 1
        Loop
    End If
    
    'yyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyy
    iCont2 = WorksheetFunction.CountIf(wsTempl_ia.Range("J:J"), sDado2)
    If iCont2 = 0 Then
        iULin = wsTempl_ia.Range("G100").End(xlUp).Row + 1
        wsTempl_ia.Cells(iULin, 7) = wsTempl_dd.Cells(i, 1)
        wsTempl_ia.Cells(iULin, 8) = wsTempl_dd.Cells(i, 19)
        wsTempl_ia.Cells(iULin, 9) = 1
        wsTempl_ia.Cells(iULin, 10) = sDado2
        
    Else
        j = 3
        bEncontrou = False
        Do While wsTempl_ia.Cells(j, 7) <> "" And Not bEncontrou
            sDadoN = wsTempl_ia.Cells(j, 10)
            If sDadoN = sDado2 Then
                wsTempl_ia.Cells(j, 9) = wsTempl_ia.Cells(j, 9) + 1
                bEncontrou = True
            End If
        j = j + 1
        Loop
    End If
    
pulaLinha:
i = i + 1
Loop

wsTempl_ia.Range("F:F").ClearContents
wsTempl_ia.Range("J:J").ClearContents

'Funcionalidade para salvar arquivo
Application.ScreenUpdating = True
Application.StatusBar = "Procedimento em progresso. Salvando Arquivo Como..."
Application.ScreenUpdating = False
sNomeArquivo = "RL-Petrobras-" & Year(Now) & "-" & Month(Now) & "-" & Day(Now) & "_" & Hour(Now) & "h" & Minute(Now) & "min"

vSalvarComo = Application.GetSaveAsFilename(InitialFileName:=sNomeArquivo, filefilter:="Excel Files (*.xlsx),*.xlsx")
If vSalvarComo <> False Then
    ActiveWorkbook.SaveAs filename:=vSalvarComo
End If

wbTemplate.Close (False)
Application.ScreenUpdating = True
Application.StatusBar = ""

End Sub



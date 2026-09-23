Attribute VB_Name = "fncEditalGuarda"
'================================================================================================= ===================
'M�dulo fun��es do edital de guarda
'Obs.: Todas as fun��es desenvolvidas conforme edital de guarda externa
'================================================================================================= ===================
 Option Explicit
 
Function calcular_TSPE_TSPR(sContratoA As String) As String
'Fun��o para calcular TSPR e TSPE
'TSPR = Total de solicita��es atendimentos dentro do prazo
'TSPE = Total de solicita��es atendidas no per�odo
'Contrato item n� 8.4.2

'VARI�VEIS
Dim iULin      As Integer
Dim i          As Integer
Dim sContratoN As String
Dim sSituacao  As String
Dim sAtividade As String
Dim iContTSPE1 As Integer
Dim iContTSPE2 As Integer
Dim iContTSPR1 As Integer
Dim iContTSPR2 As Integer
Dim iCont2     As Integer
Dim j          As Integer
Dim bEncontrou As Boolean
Dim iRef       As Integer

iULin = Planilha7.Cells(Planilha7.Rows.Count, 2).End(xlUp).Row '�ltima linha preenchida
For i = 3 To iULin
    sContratoN = Planilha7.Cells(i, 3) 'Descri��o de contrato
    'Verifica descri��o do contrato
    If sContratoA = sContratoN Then
        sSituacao = Planilha7.Cells(i, 2) 'Situa��o da solicita��o
        If sSituacao = "CO" Then
        
            sAtividade = Planilha7.Cells(i, 4) 'Atividade solicitada pelo cliente
            
            'Pesquisa refer�ncia de servi�o a se calculado
            iRef = 0
            j = 2
            bEncontrou = False
            Do While Planilha3.Cells(j, 11) <> "" And Not bEncontrou
                If sAtividade = Planilha3.Cells(j, 11) Then
                   iRef = Planilha3.Cells(j, 12)
                   bEncontrou = True
                End If
            j = j + 1
            Loop
            
            Select Case iRef
            
            Case 1 '(Servi�o de Guarda de Documentos e Informa��es)
                If Planilha7.Cells(i, 19) = "NP" Or Planilha7.Cells(i, 19) = "FP" Then
                    iContTSPE1 = iContTSPE1 + 1
                End If
                If Planilha7.Cells(i, 19) = "NP" Then iContTSPR1 = iContTSPR1 + 1
            
            Case 2 '(Servi�o de Gerenciamento de Documentos e Informa��es)
                If Planilha7.Cells(i, 19) = "NP" Or Planilha7.Cells(i, 19) = "FP" Then
                    iContTSPE2 = iContTSPE2 + 1
                End If
                If Planilha7.Cells(i, 19) = "NP" Then iContTSPR2 = iContTSPR2 + 1
            
            End Select
            
            ''==>> Case das vers�es anteriores que n�o estava desprezando o valor "FDM definido"
            'Select Case iRef
            '    Case 1 '(Servi�o de Guarda de Documentos e Informa��es)
            '        iContTSPE1 = iContTSPE1 + 1
            '        If Planilha7.Cells(i, 19) = "NP" Then iContTSPR1 = iContTSPR1 + 1
            '    Case 2 '(Servi�o de Gerenciamento de Documentos e Informa��es)
            '        iContTSPE2 = iContTSPE2 + 1
            '        If Planilha7.Cells(i, 19) = "NP" Then iContTSPR2 = iContTSPR2 + 1
            'End Select
            
        End If
    End If
Next i

calcular_TSPE_TSPR = iContTSPE1 & "|" & iContTSPR1 & "|" & iContTSPE2 & "|" & iContTSPR2

 End Function

Function calcular_IAPFARQ(iTSPE As Integer, iTSPR As Integer) As Double
'Fun��o para calcular IAPFARQ (xxxx)
'Contrato item n� 8.4.2

Dim dIAPFARQ As Double

If iTSPR = 0 Then
If iTSPE = 0 Then
dIAPFARQ = 100
Else
    dIAPFARQ = 1
    End If
Else
    'dIAPFARQ = (iTSPR * 100) / iTSPE
    dIAPFARQ = (iTSPR / iTSPE) * 100
End If

calcular_IAPFARQ = dIAPFARQ

End Function

Function calcular_FDM(dIAPFARQ As Double) As Double
'Fun��o para calcular FDM (Fator de Desempenho Mensal)
'Contrato item n� 8.4.3

Dim dFDM As Double

Select Case dIAPFARQ
    Case Is < 90
        dFDM = 0.97
    Case Is < 95       ' ? corrigido: era "< 94.99"
        dFDM = 0.98
    Case Is < 99       ' ? corrigido: era "< 98.99"
        dFDM = 0.99
    Case Is >= 99
        dFDM = 1
End Select

calcular_FDM = dFDM

End Function

Function calcular_QExec(sDrescrAtividade As String, sItem As String, lQtd As Long, sAgrupamento As String) As Double
'Fun��o para calcular quantidade atendida da guarda externa para efeitos de medi��o
'Obs.: Ver resultado em: template mem�ria de c�lculo > Planilha DADOS > Coluna AA

Dim dQExec As Double
Dim dFC As Double

Const FC_EMB As Double = 0.1
Const FC_ITEM As Double = 0.017

Select Case sDrescrAtividade

    Case "DEVOLUCAO DE EMPRESTIMO"
        If sItem = "Embalagem" Then
            If lQtd <= 10 And sAgrupamento = "" Then
                dQExec = 1
            ElseIf lQtd <= 10 And sAgrupamento <> "" Then
                dQExec = Format(FC_EMB * lQtd, "#.000")
            Else
                dQExec = Format(FC_EMB * lQtd, "#.000")
            End If
        ElseIf sItem = "Item avulso" Then
            If lQtd <= 60 And sAgrupamento = "" Then
                dQExec = 1
            ElseIf lQtd <= 60 And sAgrupamento <> "" Then
                dQExec = Format(FC_ITEM * lQtd, "#.000")
            Else
                dQExec = Format(FC_ITEM * lQtd, "#.000")
            End If
        End If
    Case "COLETA DE EMBALAGEM"
        If lQtd <= 10 And sAgrupamento = "" Then
            dQExec = 1
        ElseIf lQtd <= 10 And sAgrupamento <> "" Then
            dQExec = Format(FC_EMB * lQtd, "#.000")
        Else
            dQExec = Format(FC_EMB * lQtd, "#.000")
        End If
    Case "COLETA DE ITEM AVULSO"
        If lQtd <= 60 And sAgrupamento = "" Then
            dQExec = 1
        ElseIf lQtd <= 60 And sAgrupamento <> "" Then
            dQExec = Format(FC_ITEM * lQtd, "#.000")
        Else
            dQExec = Format(FC_ITEM * lQtd, "#.000")
        End If
    Case "ENTREGA DE EMBALAGEM EXPRESSO"
        If lQtd <= 10 And sAgrupamento = "" Then
            dQExec = 1
        ElseIf lQtd <= 10 And sAgrupamento <> "" Then
            dQExec = Format(FC_EMB * lQtd, "#.000")
        Else
            dQExec = Format(FC_EMB * lQtd, "#.000")
        End If
    Case "ENTREGA DE EMBALAGEM NORMAL"
        If lQtd <= 10 And sAgrupamento = "" Then
            dQExec = 1
        ElseIf lQtd <= 10 And sAgrupamento <> "" Then
            dQExec = Format(FC_EMB * lQtd, "#.000")
        Else
            dQExec = Format(FC_EMB * lQtd, "#.000")
        End If
    Case "ENTREGA DE ITEM EXPRESSO"
        If lQtd <= 60 And sAgrupamento = "" Then
            dQExec = 1
        ElseIf lQtd <= 60 And sAgrupamento <> "" Then
            dQExec = Format(FC_ITEM * lQtd, "#.000")
        Else
            dQExec = Format(FC_ITEM * lQtd, "#.000")
        End If
    Case "ENTREGA DE ITEM NORMAL"
        If lQtd <= 60 And sAgrupamento = "" Then
            dQExec = 1
        ElseIf lQtd <= 60 And sAgrupamento <> "" Then
            dQExec = FC_ITEM * lQtd
        Else
            dQExec = FC_ITEM * lQtd
        End If
    Case "INSERCAO DE ITEM AVULSO"
            dQExec = Format(lQtd, "#.000")
            
    Case "PESQUISA DE ITEM AVULSO EXPRESSO"
            dQExec = Format(lQtd, "#.000")
            
    Case "PESQUISA DE ITEM AVULSO NORMAL"
            dQExec = Format(lQtd, "#.000")
            
    Case "DESTRUICAO SEGURA DE DOCUMENTO"
            dQExec = Format(lQtd, "#.000")
            
    Case "TRANSFERENCIA DE ACERVO"
            dQExec = Format(lQtd, "#.000")
            
    Case "MIGRACAO DE ACERVO DOCUMENTAL"
            dFC = pesquisar_FC_migracao(sDrescrAtividade, sItem)
            dQExec = Format(dFC * lQtd, "#.000")
            
    Case "CATALOGACAO DE EMBALAGEM"
            dQExec = Format(lQtd, "#.000")
            
    Case "CATALOGACAO DE ITEM AVULSO"
            dQExec = Format(lQtd, "#.000")
            
    Case "ORGANIZACAO DE DOCUMENTO ANALITICA"
            dFC = 1
            dQExec = Format(dFC * lQtd, "#.000")
            
    Case "ORGANIZACAO DE DOCUMENTO SIMPLES"
            dFC = 0.5
            dQExec = Format(dFC * lQtd, "#.000")
            
    Case "HIGIENIZACAO DE DOCUMENTO"
            dQExec = Format(lQtd, "#.000")
            
    Case "DIGITALIZACAO DE MICROFILME"
            dFC = pesquisar_FC_digitalizacao(sDrescrAtividade, sItem)
            dQExec = Format(dFC * lQtd, "#.000")
            
    Case "DIGITALIZACAO DE MICROFICHA"
            dFC = pesquisar_FC_digitalizacao(sDrescrAtividade, sItem)
            dQExec = Format(dFC * lQtd, "#.000")
            
    Case "DIGITALIZACAO DE DOCUMENTO CONTRATUAL"
            dFC = pesquisar_FC_digitalizacao(sDrescrAtividade, sItem)
            dQExec = Format(dFC * lQtd, "#.000")
            
    Case "DIGITALIZACAO DE DOCUMENTO TECNICO"
            dFC = pesquisar_FC_digitalizacao(sDrescrAtividade, sItem)
            dQExec = Format(dFC * lQtd, "#.000")
            
    Case "DIGITALIZACAO DE DOCUMENTO BIBLIOGRAFICO"
            dFC = pesquisar_FC_digitalizacao(sDrescrAtividade, sItem)
            dQExec = Format(dFC * lQtd, "#.000")
            
    Case "DIGITALIZACAO DE DOCUMENTO CONTABIL E FINANCEIRO"
            dFC = pesquisar_FC_digitalizacao(sDrescrAtividade, sItem)
            dQExec = Format(dFC * lQtd, "#.000")
            
    Case "DIGITALIZACAO DE DOCUMENTO ADMINISTRATIVO"
            dFC = pesquisar_FC_digitalizacao(sDrescrAtividade, sItem)
            dQExec = Format(dFC * lQtd, "#.000")
            
    Case "INDEXACAO DE DOCUMENTO"
        If lQtd <= 10 Then
            dQExec = 1
        Else
            dQExec = Format(FC_EMB * lQtd, "#.000")
        End If
        
    Case "CONVERSAO DE MIDIA"
            dQExec = Format(lQtd, "#.000")
            
    Case "GRAVACAO DE MIDIA"
            dQExec = Format(lQtd, "#.000")
            
    Case "COPIA DE MIDIA"
            dQExec = Format(lQtd, "#.000")
            
    Case "COPIA DE VIDEO"
            dQExec = Format(lQtd, "#.000")
    
End Select

calcular_QExec = dQExec

End Function

Function validarAtividadeOxD(sAtividade As String) As Boolean
'Fun��o para validar atividade conforme ordem x destino

Dim iCont As Integer

iCont = WorksheetFunction.CountIf(Planilha3.Range("BH:BH"), sAtividade)

If iCont > 0 Then
    validarAtividadeOxD = True
End If

End Function

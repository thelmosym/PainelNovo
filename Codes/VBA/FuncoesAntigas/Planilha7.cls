VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'================================================================================================= ===================
'M?dulo procedimento do painel de controle
'Obs.: Todo desenvolvimento de procedimentos para o painel de controle de forma geral...
'================================================================================================= ===================
Option Explicit


Private Const MSG_PN_PAINEL1 As String = "Painel de Controle - Funcionalidades"
Private Const MSG_PN_EXP0001 As String = "Dados de sa?da (Exporta??o) - Relatorio 01"

Sub btCodigoDeAgrupamento()
'Procedimento para gerar c?digos de agrupamento

Dim objCelula As Object
Dim vPesq As Variant
Dim iLin As Integer
Dim iCol As Integer
Dim iCont As Integer
Dim bEncontrou As Boolean
Dim sAtividade As String
Dim sAtividadeN As String
Dim sCentro As String
Dim sCentroN As String
Dim sDescricao As String
Dim dtData As Date
Dim lDataVal As Long
Dim sHoraVal As String
Dim sCodArupamento As String

'Verifica c?lulas selecionadas antes de gerar c?digo de agrupamento
For Each objCelula In Selection
    If (objCelula.Value < 3 Or objCelula.Value > 2003) And objCelula.Column <> 17 Then
        MsgBox _
        "Procedimento interrompido. Linhas e/ou colunas n?o permitidas.", vbExclamation, MSG_PN_PAINEL1
        Exit Sub
    End If
    iCont = iCont + 1
Next objCelula

'Interrompe procedimento quando quantidade de c?lulas selecionadas para gerar c?digo ? igual a 1
If iCont <= 1 Then
    MsgBox _
    "Procedimento interrompido. Quantidade a partir de 2 c?lulas selecionadas.", vbExclamation, MSG_PN_PAINEL1
    Exit Sub
End If

'Verifica cada c?lula selecionada
iCont = 0
For Each objCelula In Selection
    iCont = iCont + 1
    
    'Informa linha atual selecionada
    iLin = informar_coluna(Left((objCelula.Address(True, True, xlR1C1)), InStr(objCelula.Address(True, True, xlR1C1), "C")))
    
    'Verifica se situa??o ? diferente de "CO"
    If Planilha7.Cells(iLin, 2) <> "CO" Then
        MsgBox _
        "Procedimento interrompido. Campo situa??o diferente de CO", vbExclamation, MSG_PN_PAINEL1
        Exit Sub
    End If
    
    'Informa reasultado com base na pesquisa de centro
    If iCont = 1 Then
        sAtividade = Planilha7.Cells(iLin, 4) 'Servi?o
        sCentro = pesquisar_centro(Planilha7.Cells(iLin, 22)) 'Centro
        If sCentro = "N?o encontrado" Then
            MsgBox _
            "Procedimento interrompido. Localidade/Centro n?o encontrado na tabelaB.", vbExclamation, MSG_PN_PAINEL1
            Exit Sub
        End If
    ElseIf iCont > 1 Then
        sAtividadeN = Planilha7.Cells(iLin, 4) 'Servi?o
        sCentroN = pesquisar_centro(Planilha7.Cells(iLin, 22)) 'Centro
        If sCentro = "N?o encontrado" Then
            MsgBox _
            "Procedimento interrompido. Localidade/Centro n?o encontrado na tabelaB.", vbExclamation, MSG_PN_PAINEL1
            Exit Sub
        End If
    End If
      
      
    'Verifica a partir da segunda c?lula selecionada atividades compat?veis
    'Obs.: Qualquer dos servi?os podem combinar, exceto servi?os EXPRESSO com outros servi?os
    If sAtividadeN <> "" Then
        bEncontrou = pesquisar_agrupamento(sAtividade, sAtividadeN)
        If Not bEncontrou Then
            MsgBox _
            "Procedimento interrompido. " & sAtividade & " e " & sAtividadeN & " incompat?veis para agrupamento.", vbExclamation, MSG_PN_PAINEL1
            Exit Sub
        End If
    End If
    
    
    'Verifica a partir da segunda c?lula selecionada Centro com mesma descri??o
    If sCentroN <> "" Then
        If sCentro <> sCentroN Then
            MsgBox "N?o foi poss?vel gerar c?digo de agrupamento com centro " & sCentro & " e " & sCentroN & ".", vbExclamation, _
            MSG_PN_PAINEL1
            Exit Sub
        End If
    End If
    
    
    'Verifica c?lula na planilha Obs. Isen??o com descri??o IPF
    If Planilha7.Cells(iLin, 20) = "IPF" Then
        MsgBox "N?o foi poss?vel gerar c?digo de agrupamento na observa??o de isen??o com descri??o IPF na c?lula.", _
        vbExclamation, MSG_PN_PAINEL1
        Exit Sub
    End If
    
Next objCelula
  
'Gera o c?digo de agrupamento
sDescricao = "AGP"
dtData = Format(Now(), "dd/mm/yyyy hh:mm:ss")
lDataVal = Int(CDbl(dtData))
sHoraVal = Format(dtData, "hh") & Format(dtData, "ss")
sCodArupamento = sDescricao & lDataVal & "-" & sHoraVal

'Informa nas c?lulas selecionadas oc?digo de agrupamento
For Each objCelula In Selection
     objCelula.Value = sCodArupamento
Next objCelula

MsgBox "C?digo de agrupamento gerado com sucesso.", vbInformation, MSG_PN_PAINEL1

End Sub

===============================================================================================================================

Sub preenchimentoPrazoFDM()
'Procedimento para verifica??o de Prazo e descri??o de FDM
'Obs.: Esses dados servir?o para calcular o FDM e o IAPFARQ

Dim iULin As Integer
Dim i As Integer
Dim sDescrAtividade As String
Dim iNum As Integer
Dim sRefPrazo As String
Dim dtDAbertura As Date
Dim dtDFechamento As Date
Dim sMunicipio As String
Dim dtPrazoCombinado As Date
Dim sContrato As String
Dim sUF As String
Dim sGrupo As String
Dim dPrazo As Double
Dim dtSLA As Date
Dim vTempo As Variant
Dim dtDataSLA As Date
Dim dtHoraSLA As Date
Dim dtReSLA As Date
Dim dHora As Double
Dim dtDataX As Date
Dim j As Integer
Dim iCont As Integer

'CONSTANTES
Const EXPRESSO As String = "Expresso"
Const ISENTO As String = "Isento"
Const NORMAL As String = "Normal"
Const NOPRAZO As String = "NP"
Const FORADOPRAZO As String = "FP"

'Limpa campo LOG
Planilha7.Range("AG3:AG2003").ClearContents

'Exclui valores, retira cor de fundo, padroniza fonte
Planilha7.Range("R3:S2003").ClearContents
Planilha7.Range("R3:S2003").Interior.ColorIndex = 0
Planilha7.Range("S5:S2003").Font.ColorIndex = 1

'Executa la?o de repeti??o com base nos dados solicitados
iULin = Planilha7.Range("G2003").End(xlUp).Row
For i = 3 To iULin
    'Executa somente em linhas com situa??o conclu?da
    If Planilha7.Cells(i, 2) = "CO" Then
    
        '
        Call concatenar_dados_FDM
        If Planilha7.Cells(i, 35) = "S" Then
           Planilha7.Cells(i, 18) = "FDM definido"
           GoTo proximaLinha:
        End If
        
        sDescrAtividade = Planilha7.Cells(i, 4) 'Atividade (Servi?o executado)

        'Verificar tipo da atividade
        iNum = pesquisar_referenciaFDM(sDescrAtividade)
        
        'Para atividades dentro do escopo de medi??ode FDM, a classifical?o ser? "Normal" ou "Expresso"
        If iNum > 0 Then
            'Busca pela refer?ncia de prazo
            sRefPrazo = pesquisar_referenciaPrazo(sDescrAtividade)
            
            'Tabela: s?o prazos pr?-definidos conforme contrato
            'Localizados na planilha tabelaA
            If sRefPrazo = "Tabela" Then
                sContrato = Planilha7.Cells(i, 3)
                dtDAbertura = Planilha7.Cells(i, 13)
                dtDFechamento = Planilha7.Cells(i, 14)
                sMunicipio = Planilha7.Cells(i, 23)
                sUF = Planilha7.Cells(i, 24)
                
                'Verifica se campo contrato est? vazio gera log de erro na coluna LOG
                
                If sContrato = "" Then
                    Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
                    "Coluna C" & " (" & DADOS_CAMPOC & "): " & ExibirMsg("CampoVazio") & "; "
                    GoTo proximaLinha:
                End If
                
                '
                If dtDAbertura = Empty Then
                    Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
                    "Coluna M" & " (" & DADOS_CAMPOM & "): " & ExibirMsg("CampoVazio") & "; "
                    GoTo proximaLinha:
                End If
                
                
                If dtDFechamento = Empty Then
                    Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
                    "Coluna N" & " (" & DADOS_CAMPON & "): " & ExibirMsg("CampoVazio") & "; "
                    GoTo proximaLinha:
                End If
                
                
                If sMunicipio = "" Then
                    Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
                    "Coluna W" & " (" & DADOS_CAMPOW & "): " & ExibirMsg("CampoVazio") & "; "
                    GoTo proximaLinha:
                Else
                    If Not validar_grupoMunicipio(sMunicipio) Then
                        Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
                        "Coluna W" & " (" & DADOS_CAMPOW & "): " & ExibirMsg("municipioSemGrupo") & "; "
                        GoTo proximaLinha:
                    End If
                End If
                
                
                If sUF = "" Then
                    Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
                    "Coluna X" & " (" & DADOS_CAMPOX & "): " & ExibirMsg("CampoVazio") & "; "
                    GoTo proximaLinha:
                End If
                

                'Executa SLA
                '?ndice que aponta o tempo de atendimento
                sGrupo = pesquisar_grupoMunicipio(sMunicipio, sUF)
                If sGrupo = "N?o encontrado" Then
                    Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
                    "Coluna X" & " (" & DADOS_CAMPOX & "): " & ExibirMsg("municipioXUF") & "; "
                    GoTo proximaLinha:
                End If
                dPrazo = pesquisar_prazo(sContrato, sDescrAtividade, sGrupo)
                dtSLA = dtDAbertura
                dtDataSLA = Format(dtSLA, "dd/mm/yyyy")
                dtHoraSLA = Format(dtSLA, "hh:mm")
                
                 'Os prazos encontram-se na planilha tabelaA
                Select Case dPrazo
                    'Prazo com zero n?o ser? aceito determinado servi?o para...
                    '...grupo de munic?pio ou localidade
                    
                    Case 0
                        Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
                        "Coluna S" & " (" & DADOS_CAMPOS & "): " & ExibirMsg("FDMNaoPertimido") & "; "
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
                            
                        ElseIf Format(dtDAbertura, "hh:mm") > "11:00" And Format(dtDAbertura, "hh:mm") <= "16:00" Then
                            dHora = CDbl(CDate("12:00"))
                            dtDataX = Format(dtReSLA, "dd/mm/yyyy")
                            dtDataX = dtDataX + dHora
                            
                            dtDataSLA = Format(dtDataX, "dd/mm/yyyy")
                            dtHoraSLA = Format(dtDataX, "hh:mm")
                            vTempo = "8:00"
                            dtReSLA = calcular_dataSLA(dtDataSLA, dtHoraSLA, vTempo, sMunicipio, sUF)
                            
                        Else
                            dHora = CDbl(CDate("17:00"))
                            dtDataX = Format(dtReSLA, "dd/mm/yyyy")
                            dtDataX = dtDataX + dHora
                            dtReSLA = dtDataX
                            
                        End If
                         
                    Case 1
                        vTempo = "8:00"
                        dtReSLA = calcular_dataSLA(dtDataSLA, dtHoraSLA, vTempo, sMunicipio, sUF)
                        
                    Case 2
                        vTempo = "16:00"
                        dtReSLA = calcular_dataSLA(dtDataSLA, dtHoraSLA, vTempo, sMunicipio, sUF)
                        
                    Case 3
                        vTempo = "8:00"
                        For j = 1 To 3
                            dtReSLA = calcular_dataSLA(dtDataSLA, dtHoraSLA, vTempo, sMunicipio, sUF)
                            dtDataSLA = Format(dtReSLA, "dd/mm/yyyy")
                            dtHoraSLA = Format(dtReSLA, "hh:mm")
                        Next j
                        
                    Case 4
                        vTempo = "8:00"
                        For j = 1 To 4
                            dtReSLA = calcular_dataSLA(dtDataSLA, dtHoraSLA, vTempo, sMunicipio, sUF)
                            dtDataSLA = Format(dtReSLA, "dd/mm/yyyy")
                            dtHoraSLA = Format(dtReSLA, "hh:mm")
                        Next j
                        
                    Case 5
                        vTempo = "8:00"
                        For j = 1 To 5
                            dtReSLA = calcular_dataSLA(dtDataSLA, dtHoraSLA, vTempo, sMunicipio, sUF)
                            dtDataSLA = Format(dtReSLA, "dd/mm/yyyy")
                            dtHoraSLA = Format(dtReSLA, "hh:mm")
                        Next j
                        
                End Select
                
                                
                'Obs.: Significado de siglas encontra-se na planilha instru??o de uso
                'Resultado SLA dentro do prazo
                If dtReSLA >= dtDFechamento Then
                
                    'Verificando classifica??o expressa
                    If InStr(1, sDescrAtividade, "EXPRESSO") > 0 Then
                    
                        If Planilha7.Cells(i, 20) = "IPF" Then
                            Planilha7.Cells(i, 18).Value = NORMAL
                            Planilha7.Cells(i, 19).Value = FORADOPRAZO
                            Planilha7.Cells(i, 19).Interior.ColorIndex = 9
                            Planilha7.Cells(i, 19).Font.ColorIndex = 9
                        Else
                            Planilha7.Cells(i, 18).Value = EXPRESSO
                            Planilha7.Cells(i, 19).Value = NOPRAZO
                            Planilha7.Cells(i, 19).Interior.ColorIndex = 51
                            Planilha7.Cells(i, 19).Font.ColorIndex = 51
                        End If
                        
                    'Verificando classifica??o normal
                    Else
                        Planilha7.Cells(i, 18).Value = NORMAL
                        Planilha7.Cells(i, 19).Value = NOPRAZO
                        Planilha7.Cells(i, 19).Interior.ColorIndex = 51
                        Planilha7.Cells(i, 19).Font.ColorIndex = 51
                    End If
                'Resultado SLA fora do prazo
                ElseIf dtReSLA < dtDFechamento Then
                
                    If Planilha7.Cells(i, 20) = "ABF" Or Planilha7.Cells(i, 20) = "RFT" Then
                    
                        If InStr(1, sDescrAtividade, "EXPRESSO") = 0 Then
                            Planilha7.Cells(i, 18).Value = NORMAL
                            Planilha7.Cells(i, 19).Value = NOPRAZO
                            Planilha7.Cells(i, 19).Interior.ColorIndex = 51
                            Planilha7.Cells(i, 19).Font.ColorIndex = 51
                        Else
                            Planilha7.Cells(i, 18).Value = EXPRESSO
                            Planilha7.Cells(i, 19).Value = NOPRAZO
                            Planilha7.Cells(i, 19).Interior.ColorIndex = 51
                            Planilha7.Cells(i, 19).Font.ColorIndex = 51
                        End If
                        
                    Else
                        Planilha7.Cells(i, 18).Value = NORMAL
                        Planilha7.Cells(i, 19).Value = FORADOPRAZO
                        Planilha7.Cells(i, 19).Interior.ColorIndex = 9
                        Planilha7.Cells(i, 19).Font.ColorIndex = 9
                    End If
                    
                End If
                
            ElseIf sRefPrazo = "Fiscaliza??o" Then
                dtDFechamento = Planilha7.Cells(i, 14)
                dtPrazoCombinado = Planilha7.Cells(i, 15)
                If dtDFechamento <> Empty And dtPrazoCombinado <> Empty Then
                    '
                    If dtPrazoCombinado >= dtDFechamento Then
                       Planilha7.Cells(i, 18) = NORMAL
                       Planilha7.Cells(i, 19) = NOPRAZO
                       Planilha7.Cells(i, 19).Interior.ColorIndex = 51
                       Planilha7.Cells(i, 19).Font.ColorIndex = 51
                    Else
                       Planilha7.Cells(i, 18) = NORMAL
                       Planilha7.Cells(i, 19) = FORADOPRAZO
                       Planilha7.Cells(i, 19).Interior.ColorIndex = 9
                       Planilha7.Cells(i, 19).Font.ColorIndex = 9
                    End If
                    
                'mensagem no LOG quando campo prazo combinado encontra-se vazio
                'Obs.: Obrigat?rio para refer?ncia de prazo
                ElseIf dtPrazoCombinado = Empty Then
                    Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
                    "Coluna O" & " (" & DADOS_CAMPOO & "): " & ExibirMsg("CampoVazio") & "; "
                    
                'mensagem no LOG quando campo D. fechamento encontra-se vazio
                'Obs.: Obrigat?rio para refer?ncia de prazo
                ElseIf dtDFechamento = Empty Then
                    Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
                    "Coluna N" & " (" & DADOS_CAMPON & "): " & ExibirMsg("CampoVazio") & "; "
                End If
            End If
        Else
            'Para atividades fora do escopo de medi??o de FDM, a classifica??o ser? "Isento"
            Planilha7.Cells(i, 18) = ISENTO
        End If
    End If
    
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

'Limpa dados concatenados
' Planilha7.Range("AH3:AI2003").ClearContents

iCont = WorksheetFunction.CountA(Planilha7.Range("AG3:AG2003"))
If iCont = 0 Then
    MsgBox MSG_INFORMACAO1, vbInformation, MSG_PN_PAINEL1
    
Else
    MsgBox MSG_INFORMACAO2, vbExclamation, MSG_PN_PAINEL1
    
End If

End Sub

Sub concatenar_dados_FDM()
'Procedimento para concatenar dados do painel de controle

'Obs.: A concatena??o ser? necess?ria para gerar c?lculo de FDM, pois consta no painel de controle linhas duplicadas ...
'... por causa dos atendimentos de digitaliza??o que precisam ser separados pelo item (A4, A3, A2 e etc).
'Estes itens dados precisam ser separados para mem?ria de c?lculo.

Dim ul As Integer
Dim i As Integer

ul = WorksheetFunction.Max(Planilha7.Range("B50000").End(xlUp).Row, Planilha7.Range("D50000").End(xlUp).Row, _
    Planilha7.Range("G50000").End(xlUp).Row, Planilha7.Range("L50000").End(xlUp).Row)
For i = 3 To ul
    Planilha7.Cells(i, 34) = Planilha7.Cells(i, 4) & "#" & Planilha7.Cells(i, 7) & "#" & Planilha7.Cells(i, 12)
    If WorksheetFunction.CountIf(Planilha7.Range("AH3:AH" & i), Planilha7.Cells(i, 34)) > 1 Then
        Planilha7.Cells(i, 35) = "S"
    Else
        Planilha7.Cells(i, 35) = "N"
    End If
Next i

End Sub


Sub analiseDeDados()
'Procedimento para analisar dados inseridos conforme campos do painel de controle
'Obs.: A an?lise seguir? conforme preenchimento do campo situa??o = "CO"


'VARI?VEIS
Dim iULin As Integer
Dim i As Integer
Dim iCont As Integer
Dim sCentro As String
Dim sMunicipio As String
Dim sUF As String
Dim sGrupo As String
Dim dPrazo As Double
Dim sRefPrazo As String

'Limpa campo LOG
Planilha7.Range("AG3:AG2003").ClearContents

iULin = Planilha7.Range("G2003").End(xlUp).Row
For i = 3 To iULin
    If Planilha7.Cells(i, 2) = "CO" Then
        
        
        'An?lise Coluna C
        If Planilha7.Cells(i, 3) = "" Then
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna C" & " (" & DADOS_CAMPOC & "): " & ExibirMsg("CampoVazio") & "; "
            
        Else
            iCont = WorksheetFunction.CountIf(Planilha3.Range("E:E"), Planilha7.Cells(i, 3).Value)
            If iCont = 0 Then
                Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
                "Coluna C" & " (" & DADOS_CAMPOC & "): " & ExibirMsg("descricaoTabelaA") & "; "
                
            End If
        End If
            
        
        'An?lise Coluna D
        If Planilha7.Cells(i, 4) = "" Then
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna D" & " (" & DADOS_CAMPOD & "): " & ExibirMsg("CampoVazio") & "; "
            
        Else
            iCont = WorksheetFunction.CountIf(Planilha3.Range("K:K"), Planilha7.Cells(i, 4).Value)
            If iCont = 0 Then
                Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
                "Coluna D" & " (" & DADOS_CAMPOD & "): " & ExibirMsg("descricaoTabelaA") & "; "
                
            End If
        End If
        
        
        'An?lise Coluna E
        If Planilha7.Cells(i, 5) = "" Then
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna E" & " (" & DADOS_CAMPOE & "): " & ExibirMsg("CampoVazio") & "; "
            
        Else
            iCont = WorksheetFunction.CountIf(Planilha3.Range("T:T"), Planilha7.Cells(i, 5).Value)
            If iCont = 0 Then
                Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
                "Coluna E" & " (" & DADOS_CAMPOE & "): " & ExibirMsg("descricaoTabelaA") & "; "
                
            End If
            
            If Not validar_itemDescricao(Planilha7.Cells(i, 4), Planilha7.Cells(i, 5)) Then
                Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
                "Coluna E" & " (" & DADOS_CAMPOE & "): " & ExibirMsg("itemXAtividade") & "; "
                
            End If
        End If
        
        
        
        'An?lise Coluna F
        If Planilha7.Cells(i, 6) = "" Then
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna F" & " (" & DADOS_CAMPOF & "): " & ExibirMsg("CampoVazio") & "; "
            
        End If
        
        
        'An?lise Coluna G
        If Planilha7.Cells(i, 7) = "" Then
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna G" & " (" & DADOS_CAMPOG & "): " & ExibirMsg("CampoVazio") & "; "
            
        End If
        
        
        'An?lise Coluna H
        If Planilha7.Cells(i, 8) = "" Then
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna H" & " (" & DADOS_CAMPOH & "): " & ExibirMsg("CampoVazio") & "; "
            
        ElseIf Len(Planilha7.Cells(i, 8)) <> 4 Then
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna H" & " (" & DADOS_CAMPOH & "): " & ExibirMsg("chaveUsuario") & "; "
            
        End If
        
        
        'An?lise Coluna I
        If Planilha7.Cells(i, 9) = Empty Then
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna I" & " (" & DADOS_CAMPOI & "): " & ExibirMsg("CampoVazio") & "; "
            
        ElseIf Not IsDate(Planilha7.Cells(i, 9)) Then
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna I" & " (" & DADOS_CAMPOI & "): " & ExibirMsg("dataInvalida") & "; "
            
        End If
        
        
        'An?lise Coluna J
        If Planilha7.Cells(i, 10) = "" Then
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna J" & " (" & DADOS_CAMPOJ & "): " & ExibirMsg("CampoVazio") & "; "
            
        ElseIf Len(Planilha7.Cells(i, 10)) - Len(WorksheetFunction.Substitute(Planilha7.Cells(i, 10), " ", "")) > 0 Then
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna J" & " (" & DADOS_CAMPOJ & "): " & ExibirMsg("textoInvalido") & "Espa?o; "
            
        ElseIf Len(Planilha7.Cells(i, 10)) - Len(WorksheetFunction.Substitute(Planilha7.Cells(i, 10), "_", "")) > 0 Then
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna J" & " (" & DADOS_CAMPOJ & "): " & ExibirMsg("textoInvalido") & "Underline; "
            
        ElseIf Len(Planilha7.Cells(i, 10)) - Len(WorksheetFunction.Substitute(Planilha7.Cells(i, 10), "\", "")) > 0 Then
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna J" & " (" & DADOS_CAMPOJ & "): " & ExibirMsg("textoInvalido") & "Barra invertida; "
            
        End If


        'An?lise Coluna K
        If Planilha7.Cells(i, 11) = Empty Then
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna K" & " (" & DADOS_CAMPOK & "): " & ExibirMsg("CampoVazio") & "; "
            
        ElseIf Not IsNumeric(Planilha7.Cells(i, 11)) Then
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna K" & " (" & DADOS_CAMPOK & "): " & ExibirMsg("valorInvalido") & "; "
            
        End If


        'An?lise Coluna L
        If Planilha7.Cells(i, 12) = "" Then
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna L" & " (" & DADOS_CAMPOM & "): " & ExibirMsg("CampoVazio") & "; "
            
        End If
        

        'An?lise Coluna M
        If Planilha7.Cells(i, 13) = Empty Then
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna M" & " (" & DADOS_CAMPOM & "): " & ExibirMsg("CampoVazio") & "; "
            
        ElseIf Not IsDate(Planilha7.Cells(i, 13)) Then
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna M" & " (" & DADOS_CAMPOM & "): " & ExibirMsg("dataInvalida") & "; "
            
        End If


        'An?lise Coluna N
        If Planilha7.Cells(i, 14) = Empty Then
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna N" & " (" & DADOS_CAMPON & "): " & ExibirMsg("CampoVazio") & "; "
            
        ElseIf Not IsDate(Planilha7.Cells(i, 14)) Then
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna N" & " (" & DADOS_CAMPON & "): " & ExibirMsg("dataInvalida") & "; "
            
        End If


        'An?lise Coluna O
        sRefPrazo = pesquisar_referenciaPrazo(Planilha7.Cells(i, 4))
        If sRefPrazo = "Fiscaliza??o" Then
             If Planilha7.Cells(i, 15) = Empty Then
                 Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
                "Coluna O" & " (" & DADOS_CAMPOO & "): " & ExibirMsg("atividadeSemPrazo") & "; "
                
             End If
        ElseIf sRefPrazo = "Tabela" Then
             If Planilha7.Cells(i, 15) <> Empty Then
                 Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
                "Coluna O" & " (" & DADOS_CAMPOO & "): " & ExibirMsg("dadoNaoObrigatorio") & "; "
                
             End If
             
        End If


        'An?lise Coluna P
        If Planilha7.Cells(i, 16) = Empty Then
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna P" & " (" & DADOS_CAMPOP & "): " & ExibirMsg("CampoVazio") & "; "
            
        ElseIf Not IsNumeric(Planilha7.Cells(i, 16)) Then
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna P" & " (" & DADOS_CAMPOP & "): " & ExibirMsg("valorInvalido") & "; "
            
        End If


        'An?lise Coluna Q a S
        'Obs.: campos opcionais de preencimento ou de c?lculo FDM autom?tico
        
        
        'An?lise Coluna T
        If Planilha7.Cells(i, 20) <> "" Then
            iCont = WorksheetFunction.CountIf(Planilha3.Range("BV:BV"), Planilha7.Cells(i, 20).Value)
            If iCont = 0 Then
                Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
                "Coluna T" & " (" & DADOS_CAMPOD & "): " & ExibirMsg("descricaoTabelaA") & "; "
                
            End If
        End If
        iCont = Empty
        
        
        'An?lise Coluna U
        'Obs.: campo opcional de preencimento ou de c?lculo FDM autom?tico
   
        
        'An?lise Coluna V
        'Obs.:as colunas W e X s?o preenchidas conforme coluna V
        If Planilha7.Cells(i, 22) = "" Then
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna V" & " (" & DADOS_CAMPOV & "): " & ExibirMsg("CampoVazio") & "; "
        End If
  
      
        'An?lise Coluna V
        'Obs.: Valida a localidade para centro origem X destino
        sCentro = pesquisar_centro(Planilha7.Cells(i, 22))
        If Not validar_centroOrigDest(sCentro) Then
            If validarAtividadeOxD(Planilha7.Cells(i, 4)) Then
                Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
                "Coluna V" & " (" & DADOS_CAMPOV & "): " & ExibirMsg("origemDestino") & "; "
            End If
        End If

        
        'An?lise Coluna V
        'Obs.: Valida contrato, atividade centro
        sMunicipio = pesquisar_municipio(sCentro)
        sUF = pesquisar_uf(sCentro)
        sGrupo = pesquisar_grupoMunicipio(sMunicipio, sUF)
        dPrazo = pesquisar_prazo(Planilha7.Cells(i, 3), Planilha7.Cells(i, 4), sGrupo) ' contrato + descri??o de atividade + gupo municipio tabela b
        sRefPrazo = pesquisar_referenciaPrazo(Planilha7.Cells(i, 4)) ' retorna tabela ou fiscaliza??o
        If dPrazo = 0 And sRefPrazo = "Tabela" Then
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna V" & " (" & DADOS_CAMPOV & "): " & ExibirMsg("atividadeLocalidade") & "; "
        End If
        
        
        
        'An?lise Coluna W
        If Planilha7.Cells(i, 23) <> "" Then
            If Not validar_municipioCalend(Planilha7.Cells(i, 23)) Then
                Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
                "Coluna W" & " (" & DADOS_CAMPOW & "): " & ExibirMsg("CalendarioMunicipio") & "; "
            End If
        End If
        
        
        'An?lise Coluna X
        'Obs.: Analisa contrato e UF
        If Planilha7.Cells(i, 3) <> "" And Planilha7.Cells(i, 24) <> "" Then
            If Not validar_contratoUF(Planilha7.Cells(i, 3), Planilha7.Cells(i, 24)) Then
                Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
                "Coluna W" & " (" & DADOS_CAMPOX & "): " & ExibirMsg("contratoUF") & "; "
            End If
        End If
        
        
    End If
    
Next i

iCont = WorksheetFunction.CountA(Planilha7.Range("AG3:AG2003"))
If iCont = 0 Then
    MsgBox MSG_INFORMACAO1, vbInformation, MSG_PN_PAINEL1
    
Else
    MsgBox MSG_INFORMACAO2, vbExclamation, MSG_PN_PAINEL1
    
End If

End Sub

Sub btPreenchDadosPETROBRAS()
'procedimento para preenchimento de dados da planilha dados PETROBRAS no painel de controle

Dim sMensagem As String

sMensagem = MsgBox("Preencher dados do sistema " & Planilha6.Range("L22") & "." & _
Chr(10) & "INICIAR?", vbYesNo, MSG_PN_PAINEL1)
If sMensagem = vbNo Then
    MsgBox "Preenchimento cancelado.", vbInformation, MSG_PN_PAINEL1
    Exit Sub
End If

Select Case Planilha6.Range("L22")
    Case "Click"
        Call Planilha2.preenchimentoDadosClick
    Case "ServiceNow"
        '@@@Criar script com o Service Now ativo
        MsgBox "Script n?o existe. Entre em contato com a gest?o de atendimento.", vbExclamation, MSG_PN_PAINEL1
End Select

End Sub

Sub atualizacaoDeDados()
'Procedimento para atualiza??o de dados do painel

Dim iULin As Integer
Dim i As Integer
Dim sCentro As String

iULin = Planilha7.Range("G5000").End(xlUp).Row
For i = 3 To iULin
    
    'Preenche coluna B Situa??o

    If Planilha7.Cells(i, 2) = "" Then
        Planilha7.Cells(i, 2) = "N/C"
    End If
    
    
    'Preenche coluna W Munic?pio PETROBRAS

    If Planilha7.Cells(i, 23) = "" Then
        If Planilha7.Cells(i, 22) <> "" Then
            sCentro = pesquisar_centro(Planilha7.Cells(i, 22))
            Planilha7.Cells(i, 23) = pesquisar_municipio(sCentro)
        Else
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna W" & " (" & DADOS_CAMPOP & "): " & ExibirMsg("ValorNaoEncontrado") & "; "
        End If
    End If
    
    
    'Preenche coluna X UF PETROBRAS
    If Planilha7.Cells(i, 24) = "" Then
        If Planilha7.Cells(i, 22) <> "" Then
            sCentro = pesquisar_centro(Planilha7.Cells(i, 22))
            Planilha7.Cells(i, 24) = pesquisar_uf(sCentro)
        Else
            Planilha7.Cells(i, 33) = Planilha7.Cells(i, 33) & _
            "Coluna W" & " (" & DADOS_CAMPOP & "): " & ExibirMsg("ValorNaoEncontrado") & "; "
        End If
    End If
    
    'Preenche coluna Y, Z e AA
    If Planilha7.Cells(i, 25) = "" Then
        If Planilha7.Cells(i, 4) <> "" Then
            If validar_localizacao(Planilha7.Cells(i, 4)) Then
                If Planilha7.Cells(i, 20) = "GBA" Or Planilha7.Cells(i, 20) = "GDF" Or _
                   Planilha7.Cells(i, 20) = "GES" Or Planilha7.Cells(i, 20) = "GRJ" Or _
                   Planilha7.Cells(i, 20) = "GSP" Then
                   
                    Select Case Planilha7.Cells(i, 20)
                        'Siglas para coleta transferida de uma localidade espec?fica ao galp?o desejado
                        Case "GBA" 'Galp?o Bahia
                            Planilha7.Cells(i, 25) = "PA-BA"
                            Planilha7.Cells(i, 26) = "Salvador"
                            Planilha7.Cells(i, 27) = "BA"
                            
                        Case "GDF" 'Galp?o Bras?lia
                            Planilha7.Cells(i, 25) = "IRON-DF"
                            Planilha7.Cells(i, 26) = "Distrito Federal"
                            Planilha7.Cells(i, 27) = "DF"
                            
                        Case "GES" 'Galp?o Esp?rito Santo
                            Planilha7.Cells(i, 25) = "IRON-ES"
                            Planilha7.Cells(i, 26) = "Serra"
                            Planilha7.Cells(i, 27) = "ES"
                            
                        Case "GRJ" 'Galp?o Rio de Janeiro
                            Planilha7.Cells(i, 25) = "IRON-RJ"
                            Planilha7.Cells(i, 26) = "Rio de Janeiro"
                            Planilha7.Cells(i, 27) = "RJ"

                        Case "GSP" 'Galp?o S?o Paulo
                            Planilha7.Cells(i, 25) = "IRON-SP"
                            Planilha7.Cells(i, 26) = "Cotia"
                            Planilha7.Cells(i, 27) = "SP"

                    End Select
                    
                Else
                    Select Case Planilha7.Cells(i, 24)
                        'Meta de transfer?ncia da empresa de guarda externa
                        Case "RJ", "MG" 'Lote 1 Galp?o Rio de Janeiro
                            Planilha7.Cells(i, 25) = "IRON-RJ"
                            Planilha7.Cells(i, 26) = "Rio de Janeiro"
                            Planilha7.Cells(i, 27) = "RJ"
                        Case "SP", "PR", "SC", "RS" 'Lote 1 Galp?o S?o Paulo
                            Planilha7.Cells(i, 25) = "IRON-SP"
                            Planilha7.Cells(i, 26) = "Cotia"
                            Planilha7.Cells(i, 27) = "SP"
                        Case "ES" 'Lote 1 Galp?o Esp?rito Santo
                            Planilha7.Cells(i, 25) = "IRON-ES"
                            Planilha7.Cells(i, 26) = "Serra"
                            Planilha7.Cells(i, 27) = "ES"
                        Case "DF", "GO", "MT", "MS" 'Lote 1 Galp?o Bras?lia
                            Planilha7.Cells(i, 25) = "IRON-DF"
                            Planilha7.Cells(i, 26) = "Distrito Federal"
                            Planilha7.Cells(i, 27) = "DF"
                        Case Else 'Lote 2 galp?o Salvador
                            Planilha7.Cells(i, 25) = "PA-BA"
                            Planilha7.Cells(i, 26) = "Salvador"
                            Planilha7.Cells(i, 27) = "BA"
                    End Select
                End If
            Else
                Planilha7.Cells(i, 25) = "N/A"
                Planilha7.Cells(i, 26) = "N/A"
                Planilha7.Cells(i, 27) = "N/A"
            End If
        End If
    End If
        
    'Preenche coluna C
    If Planilha7.Cells(i, 3) = "" Then
        If Planilha7.Cells(i, 24) <> "" Then
            Planilha7.Cells(i, 3) = pesquisar_descr_contrato(Planilha7.Cells(i, 24))
        End If
    End If
        
Next i

MsgBox "Procedimento realizado com sucesso.", vbInformation, MSG_PN_PAINEL1

End Sub

Sub importacaoMonitoramentoInd()
'Procedimento para importar dados individuais de monitoramento ao painel de controle
'Obs.: os dados s?o do monitoramento individual

Dim oFSO As Object
Dim oPasta As Object
Dim oArquivo As Object
Dim wbArquivo As Workbook
Dim wsMonitora As Worksheet
Dim sDir As String
Dim i As Integer
Dim iCont As Integer
Dim iULin As Integer
Dim sMensagem As Variant
Dim sPerMes As String
Dim iPerMes As Integer
Dim iPerAno As Integer
Dim iAntMes As Integer
Dim dtData As Date
Dim iDia As Integer
Dim iMes As Integer
Dim iAno As Integer
Dim sCodOS As String
Dim sConcat As String

'Diret?rio para leitura
sDir = Planilha5.Range("C50")

'Verifica se vari?vel est? vazia
If sDir = "" Then
    MsgBox _
    "Procedimento interrompido. Diret?rio vazio", vbExclamation, MSG_PN_PAINEL1
    Exit Sub
Else
    'Verifica se diret?rio existe
    If Dir(sDir, vbDirectory) = "" Then
        MsgBox _
        "Procedimento interrompido. Diret?rio n?o existe", vbExclamation, MSG_PN_PAINEL1
        Exit Sub
    End If
End If

'
sPerMes = UCase(Mid(Planilha6.Range("C9").Value, 4, 20)) 'Descri??o M?s - Per?odo informado na planilha Menu
iPerMes = Left(Planilha6.Range("C9").Value, 2) 'N?mero M?s - Per?odo informado na planilha Menu
iPerAno = Planilha6.Range("C10").Value 'Ano - Per?odo informado na planilha Menu
iAntMes = Left(Planilha6.Range("C9").Value, 2) - 1
If iAntMes = 0 Then iAntMes = 12

'Solicita
sMensagem = MsgBox("? sugerido que todos os arquivos de monitoramento estejam atualizados e fechados." & _
Chr(10) & Chr(10) & "Iniciar consulta de per?odo para: " & sPerMes & "/" & iPerAno & "?" & _
Chr(10) & Chr(10) & "Obs.: Dados do painel ser?o exclu?dos.", vbYesNo, MSG_PN_PAINEL1)
If sMensagem = vbNo Then
    MsgBox "Exporta??o cancelada.", vbInformation, MSG_PN_PAINEL1
    Exit Sub
End If

'Dados do painel ser?o apagados antes da importa??o
Planilha7.Range("A3:AG2003").ClearContents
Planilha7.Range("A3:AF2003").Interior.ColorIndex = 0
Planilha7.Range("A3:AF2003").Font.ColorIndex = 1

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oPasta = oFSO.GetFolder(sDir)

'Verifica cada arquivo dentro de uma pasta
For Each oArquivo In oPasta.Files
    If InStr(1, oArquivo.Name, "~$") = 0 Then
        Set wbArquivo = Workbooks.Open(filename:=oArquivo.Path)
        Set wsMonitora = wbArquivo.Worksheets(1)
        
        i = 3
        Do While wsMonitora.Cells(i, 9) <> ""
            
            'Se a data de fechamento estiver vazia, pula o processo
            If wsMonitora.Cells(i, 2) = Empty Then
                GoTo proxima
            End If
            
            'Se contrato n?o for da SOS,pula o processo
            If wsMonitora.Cells(i, 4) <> "IRON-LT1" And wsMonitora.Cells(i, 4) <> "PA-LT2" Then
                GoTo proxima
            End If
            
            'verificar Per?odo
            dtData = wsMonitora.Cells(i, 2).Value
            iDia = Day(dtData)
            iMes = Month(dtData)
            iAno = Year(dtData)
            If iPerMes = iMes And iDia < 26 And iPerAno = iAno Then
            ElseIf iAntMes = iMes And iDia > 25 And iDia < 32 And iPerAno = iAno Then
            ElseIf iAntMes = iMes And iDia > 25 And iDia < 32 And iAno = iPerAno - 1 Then
            Else
                GoTo proxima
            End If
        'se situa??o n?o for igual a "CO" pula
        
            If wsMonitora.Cells(i, 3) = "CO" Then
                sConcat = wsMonitora.Cells(i, 9) & "#" & wsMonitora.Cells(i, 16)
                iCont = WorksheetFunction.CountIf(Planilha7.Range("AH:AH"), sConcat)
                If iCont = 0 Then
                    iULin = Planilha7.Range("H2003").End(xlUp).Row + 1
                    'Colunas do painel          Colunas do arquivo de monitoramento
                    Planilha7.Cells(iULin, 2) = wsMonitora.Cells(i, 3) 'Situa??o
                    Planilha7.Cells(iULin, 3) = wsMonitora.Cells(i, 4) 'Contrato
                    Planilha7.Cells(iULin, 4) = wsMonitora.Cells(i, 6) 'Descri??o da atividade
                    Planilha7.Cells(iULin, 5) = wsMonitora.Cells(i, 7) 'Item
                    Planilha7.Cells(iULin, 6) = wsMonitora.Cells(i, 8) 'Aplica??o
                    Planilha7.Cells(iULin, 7) = wsMonitora.Cells(i, 9) 'C?digo da solicita??o
                    Planilha7.Cells(iULin, 8) = wsMonitora.Cells(i, 10) 'Chave solicitante
                    Planilha7.Cells(iULin, 9) = wsMonitora.Cells(i, 12) 'D. solicita??o
                    Planilha7.Cells(iULin, 10) = wsMonitora.Cells(i, 14) 'Ger?ncia solicitante
                    Planilha7.Cells(iULin, 11) = wsMonitora.Cells(i, 15) 'Qtd (solicitada)
                    sCodOS = wsMonitora.Cells(i, 16)
                    Planilha7.Cells(iULin, 12) = sCodOS 'C?digo OS
                    Planilha7.Cells(iULin, 13) = wsMonitora.Cells(i, 17) 'D. abertura
                    Planilha7.Cells(iULin, 14) = wsMonitora.Cells(i, 18) 'D. fechamento
                    Planilha7.Cells(iULin, 15) = wsMonitora.Cells(i, 19) 'Prazo combinado
                    Planilha7.Cells(iULin, 16) = wsMonitora.Cells(i, 20) 'Qtd (atendida)
                    Planilha7.Cells(iULin, 20) = wsMonitora.Cells(i, 21) 'Obs. sigla
                    Planilha7.Cells(iULin, 22) = wsMonitora.Cells(i, 22) 'Localidade
                    Planilha7.Cells(iULin, 31) = wsMonitora.Cells(i, 23) 'Coment?rios
                    Planilha7.Cells(iULin, 32) = wsMonitora.Cells(i, 24) 'Observa??es
                End If
            End If
proxima:
            
        i = i + 1
        Loop
        wbArquivo.Close (False)
    End If
Next oArquivo

'Limpa dados n?o utilizados
Planilha7.Range("AH:AH").ClearContents

'Limpa atribui??o
Set oFSO = Nothing
Set oPasta = Nothing
Set oArquivo = Nothing
Set wbArquivo = Nothing
Set wsMonitora = Nothing

MsgBox "procedimento realizado com sucesso.", vbInformation, MSG_PN_PAINEL1

Application.ScreenUpdating = True
End Sub

Sub limpezaDadosPainel()
'Procedimento para limpar todos os dados do painel de controle

Dim vResultado As VbMsgBoxResult

vResultado = MsgBox("Confirmar limpeza TODOS os dados do Painel de Controle?", vbOKCancel, MSG_PN_PAINEL1)

If vResultado = vbOK Then
    'Exclui valores, retira cor de fundo, padroniza fonte
    Planilha7.Range("A3:AG2003").ClearContents
    Planilha7.Range("A3:AF2003").Interior.ColorIndex = 0
    Planilha7.Range("A3:AF2003").Font.ColorIndex = 1
    MsgBox "Todos os dados foram exclu?dos", vbInformation, MSG_PN_PAINEL1
ElseIf vResultado = vbCancel Then
    MsgBox "Procedimento interrompido", vbInformation, MSG_PN_PAINEL1
End If

End Sub


Attribute VB_Name = "fncTabelaB"
'================================================================================================= ===================
'M?dulo fun??es tabelaB
'Obs.: Todas as fun??es desenvolvidas na tabela B
'================================================================================================= ===================
 Option Explicit


Public Function pesquisar_centro(sLocalidade As String) As String
'Fun??o para pesquisar Centro na planilha tabelaB

Dim vPesq As Variant
Dim iColA As Integer
Dim iColB As Integer
Dim nValor As String
Dim bEncontrou As Boolean
Dim i As Integer
Dim nResultado As String

vPesq = Planilha4.Rows("A1:XFD1").Find("Localidade#Localidade", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColA = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

vPesq = Planilha4.Rows("A1:XFD1").Find("Localidade#Centro", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColB = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

bEncontrou = False
i = 2
Do While Planilha4.Cells(i, iColA) <> "" And Not bEncontrou
    nValor = Planilha4.Cells(i, iColA)
    If sLocalidade = nValor Then
        nResultado = Planilha4.Cells(i, iColB)
        bEncontrou = True
    End If
i = i + 1
Loop

If bEncontrou = False Then
    nResultado = "N?o encontrado"
End If

pesquisar_centro = nResultado

End Function

--------------------------------------------------------------------------------------------------------------------------------------------

Public Function pesquisar_codCentro(sLocalidade As String) As String
'Fun??o para pesquisar Centro na planilha tabelaB

Dim vPesq As Variant
Dim iColA As Integer
Dim iColB As Integer
Dim nValor As String
Dim bEncontrou As Boolean
Dim i As Integer
Dim nResultado As String

vPesq = Planilha4.Rows("A1:XFD1").Find("Centro#Centro", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColA = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

vPesq = Planilha4.Rows("A1:XFD1").Find("Centro#C?digo Centro", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColB = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

bEncontrou = False
i = 2
Do While Planilha4.Cells(i, iColA) <> "" And Not bEncontrou
    nValor = Planilha4.Cells(i, iColA)
    If sLocalidade = nValor Then
        nResultado = Planilha4.Cells(i, iColB)
        bEncontrou = True
    End If
i = i + 1
Loop

If bEncontrou = False Then
    nResultado = "N?o encontrado"
End If

pesquisar_codCentro = nResultado

End Function


--------------------------------------------------------------------------------------------------------------------------------------------

Function calcular_KMAdicional(sOrigem As String, sDestino As String) As Double
'Fun??o para c?lculo de KM adicional

Dim vPesq As Variant
Dim iColA As Integer
Dim iColB As Integer
Dim iColC As Integer
Dim iColD As Integer
Dim iColE As Integer
Dim nValor1 As String
Dim nValor2 As String
Dim nValor3 As String
Dim nValor4 As String
Dim nResultado As Double
Dim bEncontrou As Boolean
Dim i As Integer

Const RAIO As Integer = 100

vPesq = Planilha4.Rows("A1:XFD1").Find("Origem_Destino#DE_LocalCentro", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColA = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

vPesq = Planilha4.Rows("A1:XFD1").Find("Origem_Destino#PARA_LocalCentro", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColB = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

vPesq = Planilha4.Rows("A1:XFD1").Find("Origem_Destino#DE_RegMetrop", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColC = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

vPesq = Planilha4.Rows("A1:XFD1").Find("Origem_Destino#PARA_RegMetrop", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColD = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

vPesq = Planilha4.Rows("A1:XFD1").Find("Origem_Destino#KM", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColE = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

bEncontrou = False
i = 2
Do While Planilha4.Cells(i, iColA) <> "" And Not bEncontrou
    nValor1 = Planilha4.Cells(i, iColA)
    nValor2 = Planilha4.Cells(i, iColB)
    If sOrigem = nValor1 And sDestino = nValor2 Then
        nValor3 = Planilha4.Cells(i, iColC)
        nValor4 = Planilha4.Cells(i, iColD)
        If nValor3 = nValor4 Then
            nResultado = 0
            bEncontrou = True
        Else
            nResultado = Planilha4.Cells(i, iColE)
            If nResultado <= RAIO Then
                nResultado = 0
            Else
                nResultado = nResultado - RAIO
                bEncontrou = True
            End If
        End If
    End If
i = i + 1
Loop

calcular_KMAdicional = nResultado

End Function

--------------------------------------------------------------------------------------------------------------------------------------------

Function pesquisar_precoUnitario(sContratoA As String, sLinhaPPU As String) As Double
'Fun??o para pesquisar pre?o unit?tio conforme contrato e linha de servi?o PPU

Dim vPesq As Variant
Dim iColA As Integer
Dim iColB As Integer
Dim iColC As Integer
Dim nValor1 As String
Dim nValor2 As String
Dim bEncontrou As Boolean
Dim i As Integer
Dim nResultado As Double

vPesq = Planilha4.Rows("A1:XFD1").Find("PPU#Contrato", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColA = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

vPesq = Planilha4.Rows("A1:XFD1").Find("PPU#Linha de Servi?o PPU", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColB = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

vPesq = Planilha4.Rows("A1:XFD1").Find("PPU#P. Unit?rio (R$)", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColC = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

bEncontrou = False
i = 2
Do While Planilha4.Cells(i, iColA) <> "" And Not bEncontrou
    nValor1 = Planilha4.Cells(i, iColA)
    nValor2 = Planilha4.Cells(i, iColB)
    If sContratoA = nValor1 And sLinhaPPU = nValor2 Then
        nResultado = Planilha4.Cells(i, iColC)
        bEncontrou = True
    End If
i = i + 1
Loop

If bEncontrou = False Then
    nResultado = 0
End If

pesquisar_precoUnitario = nResultado

End Function


--------------------------------------------------------------------------------------------------------------------------------------------

Function pesquisar_grupoMunicipio(sMunicipio As String, sUF As String) As String
'Fun??o para pesquisar grupo de munic?pio
'Obs.: A sigla do grupo de munic?pio ? refer?ncia para prazo das atividades executadas pela guarda externa

Dim vPesq As Variant
Dim iColA As Integer
Dim iColB As Integer
Dim iColC As Integer
Dim nValor1 As String
Dim nValor2 As String
Dim bEncontrou As Boolean
Dim i As Integer
Dim nResultado As String

vPesq = Planilha4.Rows("A1:XFD1").Find("Grupo#Munic?pio", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColA = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

vPesq = Planilha4.Rows("A1:XFD1").Find("Grupo#UF", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColB = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

vPesq = Planilha4.Rows("A1:XFD1").Find("Grupo#Grupo", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColC = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

bEncontrou = False
i = 2
Do While Planilha4.Cells(i, iColA) <> "" And Not bEncontrou
    nValor1 = Planilha4.Cells(i, iColA)
    nValor2 = Planilha4.Cells(i, iColB)
    If sMunicipio = nValor1 And sUF = nValor2 Then
        nResultado = Planilha4.Cells(i, iColC)
        bEncontrou = True
    End If
i = i + 1
Loop

If bEncontrou = False Then
    nResultado = "N?o encontrado"
End If

pesquisar_grupoMunicipio = nResultado

End Function

--------------------------------------------------------------------------------------------------------------------------------------------

Function pesquisar_EFeriado(dDataInicial As Date, sMunicipio As String, sUF As String) As Boolean
'Fun??o pesquisa se Data e munic?pio s?o feriados


Dim vPesq As Variant
Dim iColA As Integer
Dim iColB As Integer
Dim iColC As Integer
Dim nValor1 As Date
Dim nValor2 As String
Dim nValor3 As String
Dim bEncontrou As Boolean
Dim i As Integer
Dim nResultado As Boolean

'vPesq = Planilha4.Rows("A1:XFD1").Find("Calend?rio#Data", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
vPesq = Planilha4.Rows("A1:XFD1").Find("Calendario#Data", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColA = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

'vPesq = Planilha4.Rows("A1:XFD1").Find("Calend?rio#Munic?pio", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
vPesq = Planilha4.Rows("A1:XFD1").Find("Calendario#Munic?pio", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColB = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

'vPesq = Planilha4.Rows("A1:XFD1").Find("Calend?rio#UF", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
vPesq = Planilha4.Rows("A1:XFD1").Find("Calendario#UF", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColC = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

bEncontrou = False
i = 2
Do While Planilha4.Cells(i, iColA) <> "" And Not bEncontrou
    nValor1 = Planilha4.Cells(i, iColA)
    nValor2 = Planilha4.Cells(i, iColB)
    nValor3 = Planilha4.Cells(i, iColC)
    If dDataInicial = nValor1 And sMunicipio = nValor2 And sUF = nValor3 Then
        nResultado = True
        bEncontrou = True
    End If
i = i + 1
Loop

pesquisar_EFeriado = nResultado

End Function

Function pesquisar_municipio(sCentro As String) As String
'Fun??o pesquisa munic?pio


Dim vPesq As Variant
Dim iColA As Integer
Dim iColB As Integer
Dim nValor1 As String
Dim bEncontrou As Boolean
Dim i As Integer
Dim nResultado As String

vPesq = Planilha4.Rows("A1:XFD1").Find("Centro#Centro", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColA = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

vPesq = Planilha4.Rows("A1:XFD1").Find("Centro#Munic?pio", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColB = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

bEncontrou = False
i = 2
Do While Planilha4.Cells(i, iColA) <> "" And Not bEncontrou
    nValor1 = Planilha4.Cells(i, iColA)
    If sCentro = nValor1 Then
        nResultado = Planilha4.Cells(i, iColB)
        bEncontrou = True
    End If
i = i + 1
Loop

pesquisar_municipio = nResultado

End Function

Function pesquisar_uf(sCentro As String) As String
'Fun??o pesquisa unidade federativa


Dim vPesq As Variant
Dim iColA As Integer
Dim iColB As Integer
Dim nValor1 As String
Dim bEncontrou As Boolean
Dim i As Integer
Dim nResultado As String

vPesq = Planilha4.Rows("A1:XFD1").Find("Centro#Centro", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColA = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

vPesq = Planilha4.Rows("A1:XFD1").Find("Centro#UF", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColB = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

bEncontrou = False
i = 2
Do While Planilha4.Cells(i, iColA) <> "" And Not bEncontrou
    nValor1 = Planilha4.Cells(i, iColA)
    If sCentro = nValor1 Then
        nResultado = Planilha4.Cells(i, iColB)
        bEncontrou = True
    End If
i = i + 1
Loop

pesquisar_uf = nResultado

End Function

Function pesquisar_descr_contrato(sUF As String) As String
'Fun??o pesquisa unidade federativa


Dim vPesq As Variant
Dim iColA As Integer
Dim iColB As Integer
Dim nValor1 As String
Dim bEncontrou As Boolean
Dim i As Integer
Dim nResultado As String

vPesq = Planilha4.Rows("A1:XFD1").Find("Grupo#UF", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColA = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

vPesq = Planilha4.Rows("A1:XFD1").Find("Grupo#Contrato", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColB = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

bEncontrou = False
i = 2
Do While Planilha4.Cells(i, iColA) <> "" And Not bEncontrou
    nValor1 = Planilha4.Cells(i, iColA)
    If sUF = nValor1 Then
        nResultado = Planilha4.Cells(i, iColB)
        bEncontrou = True
    End If
i = i + 1
Loop

pesquisar_descr_contrato = nResultado

End Function

Public Function validar_centroOrigDest(sCentro As String) As Boolean
'Fun??o para validar Centro para origem X destino

Dim vPesq As Variant
Dim iColA As Integer
Dim sColA As String
Dim iCont As Integer
Dim nResultado As String
Dim sIntervalo As String

vPesq = Planilha4.Rows("A1:XFD1").Find("Origem_Destino#DE_LocalCentro", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColA = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

nResultado = False
sColA = converter_numParaLetra(iColA)
sIntervalo = "" & sColA & ":" & sColA & ""
iCont = WorksheetFunction.CountIf(Planilha4.Range(sIntervalo), sCentro)
If iCont > 0 Then
    nResultado = True
End If

validar_centroOrigDest = nResultado

End Function


Public Function validar_municipioCalend(sMunicipio As String) As Boolean
'Fun??o para validar Munic?pio no calend?rio

Dim vPesq As Variant
Dim iColA As Integer
Dim sColA As String
Dim iCont As Integer
Dim nResultado As String
Dim sIntervalo As String

vPesq = Planilha4.Rows("A1:XFD1").Find("Calendario#Munic?pio", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColA = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

nResultado = False
sColA = converter_numParaLetra(iColA)
sIntervalo = "" & sColA & ":" & sColA & ""
iCont = WorksheetFunction.CountIf(Planilha4.Range(sIntervalo), sMunicipio)
If iCont > 0 Then
    nResultado = True
End If

validar_municipioCalend = nResultado

End Function

Public Function validar_grupoMunicipio(sMunicipio As String) As Boolean
'Fun??o para validar Munic?pio no calend?rio

Dim vPesq As Variant
Dim iColA As Integer
Dim sColA As String
Dim iCont As Integer
Dim nResultado As String
Dim sIntervalo As String

vPesq = Planilha4.Rows("A1:XFD1").Find("Grupo#Munic?pio", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColA = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

nResultado = False
sColA = converter_numParaLetra(iColA)
sIntervalo = "" & sColA & ":" & sColA & ""
iCont = WorksheetFunction.CountIf(Planilha4.Range(sIntervalo), sMunicipio)
If iCont > 0 Then
    nResultado = True
End If

validar_grupoMunicipio = nResultado

End Function

Function recalcular_FDM(sPerVig As String, sPerAnt1 As String, sPerAnt2 As String, sContrato As String, dFDM As Double, iItem As Integer) As Double
'

Dim dFDMIAnt1 As Double
Dim dFDMIAnt2 As Double

Select Case dFDM
    Case Is > 0.97
        recalcular_FDM = dFDM
        Exit Function
        
    Case Is = 0.97
        dFDMIAnt1 = validar_FDMAnterior(sPerAnt1, sContrato, iItem)
        dFDMIAnt2 = validar_FDMAnterior(sPerAnt2, sContrato, iItem)
        
        If dFDMIAnt1 >= 0.98 Then
            recalcular_FDM = dFDM
            Exit Function
            
        ElseIf dFDMIAnt1 = 0.97 And dFDMIAnt2 > 0.97 Then
            recalcular_FDM = 0.96
            Exit Function
            
        ElseIf dFDMIAnt1 = 0.97 And dFDMIAnt2 = 0.97 Then
            recalcular_FDM = 0.95
            Exit Function
            
        ElseIf dFDMIAnt1 = 0.97 And dFDMIAnt2 < 0.97 Then
            recalcular_FDM = 0.95
            Exit Function
            
        ElseIf dFDMIAnt1 = 0.96 Then
            recalcular_FDM = 0.95
            Exit Function
            
        ElseIf dFDMIAnt1 = 0.95 Then
            recalcular_FDM = 0.95
            Exit Function
            
        End If
        
End Select


End Function

Public Function validar_FDMAnterior(sPerAnt As String, sContrato As String, iItem As Integer) As Double
'Fun??o para validar FDM anterior do per?odo vigente

Dim vPesq As Variant
Dim iColA As Integer
Dim iColB As Integer
Dim iColC As Integer
Dim iColD As Integer
Dim nValor1 As String
Dim nValor2 As String
Dim nValor3 As Integer
Dim bEncontrou As Boolean
Dim nResultado As Double
Dim i As Integer

vPesq = Planilha4.Rows("A1:XFD1").Find("FDM_IAPFARQ#Medi??o", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColA = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

vPesq = Planilha4.Rows("A1:XFD1").Find("FDM_IAPFARQ#Contrato", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColB = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

vPesq = Planilha4.Rows("A1:XFD1").Find("FDM_IAPFARQ#Item contrato", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColC = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

vPesq = Planilha4.Rows("A1:XFD1").Find("FDM_IAPFARQ#FDM", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColD = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

bEncontrou = False
i = 2
Do While Planilha4.Cells(i, iColA) <> "" And Not bEncontrou
    nValor1 = Planilha4.Cells(i, iColA)
    nValor2 = Planilha4.Cells(i, iColB)
    nValor3 = Planilha4.Cells(i, iColC)
    If sPerAnt = nValor1 And sContrato = nValor2 And iItem = nValor3 Then
        nResultado = Planilha4.Cells(i, iColD)
        bEncontrou = True
    End If
i = i + 1
Loop

validar_FDMAnterior = nResultado

End Function

Public Function validar_periodo_anterior(sPerAnterior As String) As Boolean
'Fun??o para validar per?odo anterior

Dim vPesq As Variant
Dim iColA As Integer
Dim sColA As String
Dim iCont As Integer
Dim nResultado As String
Dim sIntervalo As String

vPesq = Planilha4.Rows("A1:XFD1").Find("FDM_IAPFARQ#Medi??o", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
iColA = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))

nResultado = False
sColA = converter_numParaLetra(iColA)
sIntervalo = "" & sColA & ":" & sColA & ""
iCont = WorksheetFunction.CountIf(Planilha4.Range(sIntervalo), sPerAnterior)
If iCont > 0 Then
    nResultado = True
End If

validar_periodo_anterior = nResultado

End Function



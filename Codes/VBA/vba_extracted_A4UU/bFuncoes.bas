Attribute VB_Name = "bFuncoes"

'================================================================================================= ===================

'Módulo funções tabelaB

'Obs.: Todas as funções desenvolvidas na tabela B

'================================================================================================= ===================

 Option Explicit







Public Function pesquisar_centro(sLocalidade As String) As String

'Função para pesquisar Centro na planilha tabelaB



Dim vPesq As Variant

Dim iColA As Integer

Dim iColB As Integer

Dim nValor As String

Dim bEncontrou As Boolean

Dim i As Integer

Dim nResultado As String



vPesq = Planilha5.Rows("A1:XFD1").Find("Localidade#Localidade", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)

iColA = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))



vPesq = Planilha5.Rows("A1:XFD1").Find("Localidade#Centro", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)

iColB = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))



bEncontrou = False

i = 2

Do While Planilha5.Cells(i, iColA) <> "" And Not bEncontrou

    nValor = Planilha5.Cells(i, iColA)

    If sLocalidade = nValor Then

        nResultado = Planilha5.Cells(i, iColB)

        bEncontrou = True

    End If

i = i + 1

Loop



If bEncontrou = False Then

    nResultado = "Não encontrado"

End If



pesquisar_centro = nResultado



End Function



Function pesquisar_EFeriado(dDataInicial As Date, sMunicipio As String, sUF As String) As Boolean

'Função pesquisa se Data e município são feriados



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



vPesq = Planilha5.Rows("A1:XFD1").Find("Calendario#Data", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)

iColA = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))



vPesq = Planilha5.Rows("A1:XFD1").Find("Calendario#Município", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)

iColB = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))



vPesq = Planilha5.Rows("A1:XFD1").Find("Calendario#UF", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)

iColC = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))



bEncontrou = False

i = 2

Do While Planilha5.Cells(i, iColA) <> "" And Not bEncontrou

    nValor1 = Planilha5.Cells(i, iColA)

    nValor2 = Planilha5.Cells(i, iColB)

    nValor3 = Planilha5.Cells(i, iColC)

    If dDataInicial = nValor1 And sMunicipio = nValor2 And sUF = nValor3 Then

        nResultado = True

        bEncontrou = True

    End If

i = i + 1

Loop



pesquisar_EFeriado = nResultado



End Function



Function pesquisar_municipioUF(sCentro As String) As String

'Função pesquisa município





Dim vPesq As Variant

Dim iColA As Integer

Dim iColB As Integer

Dim iColC As Integer

Dim nValor1 As String

Dim bEncontrou As Boolean

Dim i As Integer

Dim nResultado As String



vPesq = Planilha5.Rows("A1:XFD1").Find("Centro#Centro", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)

iColA = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))



vPesq = Planilha5.Rows("A1:XFD1").Find("Centro#Município", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)

iColB = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))



vPesq = Planilha5.Rows("A1:XFD1").Find("Centro#UF", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)

iColC = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))



bEncontrou = False

i = 2

Do While Planilha5.Cells(i, iColA) <> "" And Not bEncontrou

    nValor1 = Planilha5.Cells(i, iColA)

    If sCentro = nValor1 Then

        nResultado = Planilha5.Cells(i, iColB) & "|" & Planilha5.Cells(i, iColC)

        bEncontrou = True

    End If

i = i + 1

Loop



pesquisar_municipioUF = nResultado



End Function





Public Function validar_municipioCalend(sMunicipio As String) As Boolean

'Função para validar Município no calendário



Dim vPesq As Variant

Dim iColA As Integer

Dim sColA As String

Dim iCont As Integer

Dim nResultado As String

Dim sIntervalo As String



vPesq = Planilha5.Rows("A1:XFD1").Find("Calendario#Município", LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)

iColA = informar_coluna(Mid((vPesq), InStr(vPesq, "C")))



nResultado = False

sColA = converter_numParaLetra(iColA)

sIntervalo = "" & sColA & ":" & sColA & ""

iCont = WorksheetFunction.CountIf(Planilha5.Range(sIntervalo), sMunicipio)

If iCont > 0 Then

    nResultado = True

End If



validar_municipioCalend = nResultado



End Function




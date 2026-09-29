Attribute VB_Name = "Planilha5"

Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = False

Attribute VB_PredeclaredId = True

Attribute VB_Exposed = True

Attribute VB_TemplateDerived = False

Attribute VB_Customizable = True

'================================================================================================= ===================

'Módulo procedimento tabelaB

'Obs.: Todo desenvolvimento de procedimentos para a planilha tabelaB de forma geral.

'================================================================================================= ===================

Option Explicit





Private Const MSG_TAB1 As String = "Monitoramento - Importação tabelaB"



Sub btImportacaoTabelaB()

'Copia dados de cada guia/planilha do arquivo "Tabelas" para o monitoramento > planilha tabelaB



'VARIÁVEIS

Dim sDirTabelaB As String

Dim sNomeTabelaB As String



Dim wbTabelaB As Workbook



Dim iWSCont As Integer

Dim iColA As Integer

Dim iPlan As Integer

Dim wsPlan As Worksheet

Dim sNomePlan As String

Dim c As Integer

Dim sColLetra As String

Dim iULin As Integer



Dim sLin1 As String

Dim sLin2 As String



Application.ScreenUpdating = False



'apagando dados anteriores

Planilha5.Range("A1:BZ10000").ClearContents

Planilha5.Columns("A:BZ").ColumnWidth = 9



'Grava diretório conforme planilha "Rede"

sDirTabelaB = Planilha3.Range("C27")



'Verifica se arquivo existe na pasta/diretório

If Dir(sDirTabelaB) = "" Then

    MsgBox "Procedimento interrompido. Arquivo e/ou diretório não existe." & Chr(10) & _

    Chr(10) & _

    "Obs.: Verifique diretório informado e tente novamente.", vbCritical, MSG_TAB1

    Exit Sub

End If





'Atribui variável

Set wbTabelaB = Workbooks.Open(sDirTabelaB)



iWSCont = wbTabelaB.Worksheets.Count 'Conta quantidade de planilhas

iColA = 1

For iPlan = 1 To iWSCont 'Executa laço de repetição para cada planilha

    Set wsPlan = wbTabelaB.Worksheets(iPlan)

    sNomePlan = wsPlan.Name

    c = 1

    Do While wsPlan.Cells(1, c) <> ""

        sColLetra = converter_numParaLetra(c)

        iULin = wsPlan.Range("A10000").End(xlUp).Row

        Planilha5.Cells(1, iColA) = sNomePlan

        wsPlan.Range(sColLetra & "1:" & sColLetra & iULin).Copy

        Planilha5.Cells(2, iColA).PasteSpecial xlPasteValues

        Application.CutCopyMode = False

        iColA = iColA + 1

    c = c + 1

    Loop

Next iPlan



wbTabelaB.Close (False)



'Executa laço de repetição para descrição de cada campo

'Concatena campos (Exemplo: XXXXXX#XXXXXXX)

c = 1

Do While Planilha5.Cells(1, c) <> ""

    sLin1 = Planilha5.Cells(1, c).Value

    sLin2 = Planilha5.Cells(2, c).Value

    Planilha5.Cells(1, c) = sLin1 & "#" & sLin2

c = c + 1

Loop



Planilha5.Rows(2).Delete

Planilha5.Columns("A:BZ").AutoFit

Planilha3.Activate

Planilha3.Range("G16") = Now



Application.ScreenUpdating = True



MsgBox "Procedimento realizado com sucesso", vbInformation, MSG_TAB1



End Sub




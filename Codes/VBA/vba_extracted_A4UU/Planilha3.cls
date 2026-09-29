Attribute VB_Name = "Planilha3"

Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"

Attribute VB_GlobalNameSpace = False

Attribute VB_Creatable = False

Attribute VB_PredeclaredId = True

Attribute VB_Exposed = True

Attribute VB_TemplateDerived = False

Attribute VB_Customizable = True

Option Explicit



Private Const MSG1 = "Agrupamento monitoramento"



Sub agrupamento_arquivo_monitoramento()

'Agrupa os dados de arquivos de monitoramento em um único arquivo



Dim objFSO As Object

Dim objPasta As Object

Dim objArq As Object

Dim sDir As String

Dim wbArq As Workbook

Dim wsPlan As Worksheet

Dim ul1 As Integer

Dim ul2 As Integer



'Variável para definir o diretório

sDir = Application.InputBox("Insira um diretório para leitura de pastas:" & Chr(10) & _

"Obs.: A aplicação não faz leitura de subpastas.", MSG1)



'Verifica se diretório existe

If Dir(sDir, vbDirectory) = "" Then

    MsgBox "Diretório para leitura de pasta não encontrado.", _

    vbExclamation, MSG1

    Exit Sub

End If



'Limpa área do monitoramento

Planilha1.Range("A3:AL50000").ClearContents



Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objPasta = objFSO.GetFolder(sDir)



'Desabilita eventos de mensagens do mapa de produção

Application.EnableEvents = False



For Each objArq In objPasta.Files

    If InStr(1, objArq.Name, "~$") = 0 Then 'Verificação de arquivos temporários

        Set wbArq = Workbooks.Open(Filename:=objArq.Path)

        Set wsPlan = wbArq.Worksheets(1)

        

        'Retira o filtro das planilhas

        If wsPlan.AutoFilterMode Then

            wsPlan.AutoFilter.ShowAllData

        End If

        

        ul1 = wsPlan.Range("I50000").End(xlUp).Row

        

        'Ativa a planilha

        wsPlan.Activate

        

        'Faz cópia das informações

        wsPlan.Range("A3:AL" & ul1).Copy

        

        'Ativa o mapa a ser copiado

        ThisWorkbook.Activate

        Planilha1.Select

        ul2 = Planilha1.Range("I50000").End(xlUp).Row + 1

        Planilha1.Cells(ul2, 1).Select

        Selection.PasteSpecial Paste:=xlPasteValues

        wbArq.Application.CutCopyMode = False

        

        wbArq.Close (False)

    End If

Next objArq



Application.EnableEvents = True



Set objArq = Nothing

Set objPasta = Nothing

Set objFSO = Nothing



MsgBox "Procedimento realizado com sucesso", vbInformation, MSG1



End Sub




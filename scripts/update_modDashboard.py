import os

vba_code = '''Attribute VB_Name = "modDashboard"
Option Explicit

'====================================================================================================
' MÓDULO: modDashboard
' OBJETIVO: Gerenciar a Dashboard Executiva BI (01_DASHBOARD), motor de cálculo (04_CALCULOS),
'           listas de validação (06_LISTAS) e documentação operacional (00_GUIA).
' REGRAS DE NEGÓCIO: Petrobras Contratos PA-LT2 e IRON-LT1 (RJ, SP, ES). Ciclo 26 a 25.
'====================================================================================================

Public Sub AtualizarDashboard()
    Dim wsDash As Worksheet
    Dim wsCalc As Worksheet
    Dim wsGuia As Worksheet
    Dim bEvents As Boolean
    Dim bScreen As Boolean
    Dim appCalc As XlCalculation

    On Error GoTo TratarErro

    bScreen = Application.ScreenUpdating
    bEvents = Application.EnableEvents
    appCalc = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' 1. Atualizar consultas Power Query se desejado
    On Error Resume Next
    ThisWorkbook.RefreshAll
    On Error GoTo TratarErro

    ' 2. Forçar recálculo completo da pasta de trabalho
    Application.CalculateFull

    Application.Calculation = xlCalculationAutomatic
    Application.Calculate

    Set wsDash = ObterPlanilha("01_DASHBOARD")
    If Not wsDash Is Nothing Then
        wsDash.Activate
        wsDash.Range("C6").Select
    End If

    Application.ScreenUpdating = bScreen
    Application.EnableEvents = bEvents

    MsgBox "Dashboard e motor analítico atualizados com sucesso!", vbInformation, "BI Petrobras"
    Exit Sub

TratarErro:
    Application.Calculation = appCalc
    Application.EnableEvents = bEvents
    Application.ScreenUpdating = bScreen
    MsgBox "Erro ao atualizar Dashboard:" & vbCrLf & Err.Number & " - " & Err.Description, vbCritical, "Erro de Atualização"
End Sub

Public Sub ResetarFiltrosDashboard()
    Dim wsDash As Worksheet
    On Error Resume Next
    Set wsDash = ObterPlanilha("01_DASHBOARD")
    If Not wsDash Is Nothing Then
        wsDash.Range("C6").Value = "Todos os Contratos"
        wsDash.Range("F6").Value = "Todas as Atividades"
        wsDash.Range("I6").Value = "Todos"
        Application.Calculate
    End If
End Sub

Private Function ObterPlanilha(ByVal nomeAba As String) As Worksheet
    On Error Resume Next
    Set ObterPlanilha = ThisWorkbook.Worksheets(nomeAba)
    On Error GoTo 0
End Function
'''

target_path = r"Codes/VBA/Funções Novas/modDashboard.bas"
with open(target_path, "wb") as f:
    f.write(vba_code.replace("\r\n", "\n").replace("\n", "\r\n").encode("latin1"))

print(f"Arquivo {target_path} gravado com sucesso com codificação Windows-1252 (ANSI) e quebras CRLF.")

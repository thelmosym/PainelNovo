Attribute VB_Name = "fncPainelDeControle"
'================================================================================================= ===================
'Módulo função geral da memória de cálculo de guarda externa
'Obs.: Todo desenvolvimento de funções para memória de cálculo dos contratos da SBDC de forma geral
'================================================================================================= ===================

Option Explicit
Function verificar_celulaEmBranco() As Boolean
'

Dim i As Integer
Dim j As Integer
Dim iULin As Integer

iULin = Planilha7.Range("C2003").End(xlUp).Row
For i = 3 To iULin

    '
    For j = 2 To 14
        If Planilha7.Cells(i, j) = "" Then
            verificar_celulaEmBranco = True
            Exit Function
        End If
    Next j
        
        '
        If Planilha7.Cells(i, 16) = "" Then
            verificar_celulaEmBranco = True
        End If
    
    '
    For j = 18 To 19
        
        If Planilha7.Cells(i, 18) = "" Then
            verificar_celulaEmBranco = True
        End If
    Next j
    
    '
    For j = 22 To 27
        If Planilha7.Cells(i, j) = "" Then
            verificar_celulaEmBranco = True
        End If
    Next j
    
Next i

End Function

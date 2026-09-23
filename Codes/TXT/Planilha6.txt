VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'================================================================================================= ===================
'Módulo procedimento do menu do painel
'Obs.: Todo desenvolvimento de procedimentos para o menu  do painel de controle de forma geral...
'================================================================================================= ===================
Option Explicit

Const MSG_PN_MENU1 As String = "Painel de Controle - Configurações"

Sub gerenciadorDeNome()
'Procedimento para gerar listas de validação conforme atividades do contrato no painel de controle
'Obs.: serão geradas listas conforme planilha tabelaA

'Variáveis
Dim i As Integer
Dim sAtividade As String
Dim vPqsLin As Variant
Dim iPriLin As Integer
Dim iContLin As Integer
Dim iLinA As Integer
Dim iLinN As Integer
Dim sInterv As String
Dim sConj As String

On Error Resume Next

'Excluir nomes do gerenciador de nomes
ActiveWorkbook.Names("aa").Delete 'ARMAZENAMENTO DE ACERVO DOCUMENTAL
ActiveWorkbook.Names("ba").Delete 'BAIXA PERMANENTE
ActiveWorkbook.Names("ca").Delete 'CATALOGACAO DE EMBALAGEM
ActiveWorkbook.Names("cb").Delete 'CATALOGACAO DE ITEM AVULSO
ActiveWorkbook.Names("cc").Delete 'COLETA DE EMBALAGEM
ActiveWorkbook.Names("cd").Delete 'COLETA DE ITEM AVULSO
ActiveWorkbook.Names("ce").Delete 'CONVERSAO DE MIDIA
ActiveWorkbook.Names("cf").Delete 'COPIA DE MIDIA
ActiveWorkbook.Names("cg").Delete 'COPIA DE VIDEO
ActiveWorkbook.Names("da").Delete 'DESTRUICAO SEGURA DE DOCUMENTO
ActiveWorkbook.Names("db").Delete 'DEVOLUCAO DE EMPRESTIMO
ActiveWorkbook.Names("dc").Delete 'DIGITALIZACAO DE MICROFILME
ActiveWorkbook.Names("dd").Delete 'DIGITALIZACAO DE MICROFICHA
ActiveWorkbook.Names("de").Delete 'DIGITALIZACAO DE DOCUMENTO CONTRATUAL
ActiveWorkbook.Names("df").Delete 'DIGITALIZACAO DE DOCUMENTO TECNICO
ActiveWorkbook.Names("dg").Delete 'DIGITALIZACAO DE DOCUMENTO BIBLIOGRAFICO
ActiveWorkbook.Names("dh").Delete 'DIGITALIZACAO DE DOCUMENTO CONTABIL E FINANCEIRO
ActiveWorkbook.Names("di").Delete 'DIGITALIZACAO DE DOCUMENTO ADMINISTRATIVO
ActiveWorkbook.Names("ea").Delete 'ENTREGA DE EMBALAGEM EXPRESSO
ActiveWorkbook.Names("eb").Delete 'ENTREGA DE EMBALAGEM NORMAL
ActiveWorkbook.Names("ec").Delete 'ENTREGA DE ITEM EXPRESSO
ActiveWorkbook.Names("ed").Delete 'ENTREGA DE ITEM NORMAL
ActiveWorkbook.Names("ga").Delete 'GRAVACAO DE MIDIA
ActiveWorkbook.Names("ha").Delete 'HIGIENIZACAO DE DOCUMENTO
ActiveWorkbook.Names("ia").Delete 'INDEXACAO DE DOCUMENTO
ActiveWorkbook.Names("ib").Delete 'INSERCAO DE ITEM AVULSO
ActiveWorkbook.Names("ma").Delete 'MATERIAL PARA ARQUIVAMENTO
ActiveWorkbook.Names("mb").Delete 'MIGRACAO DE ACERVO DOCUMENTAL
ActiveWorkbook.Names("oa").Delete 'ORGANIZACAO DE DOCUMENTO ANALITICA
ActiveWorkbook.Names("os").Delete 'ORGANIZACAO DE DOCUMENTO SIMPLES
ActiveWorkbook.Names("pa").Delete 'PESQUISA DE ITEM AVULSO EXPRESSO
ActiveWorkbook.Names("pb").Delete 'PESQUISA DE ITEM AVULSO NORMAL
ActiveWorkbook.Names("ra").Delete 'RENOVACAO DE EMPRESTIMO
ActiveWorkbook.Names("ta").Delete 'TRANSFERENCIA DE ACERVO


'consulta descrição da atividade como referência para as listas de validação
i = 2
Do While Planilha3.Cells(i, 11) <> ""
sAtividade = Planilha3.Cells(i, 11)
Select Case sAtividade
Case "ARMAZENAMENTO DE ACERVO DOCUMENTAL"
    sAtividade = "ARMAZENAMENTO DE ACERVO DOCUMENTAL"
    sConj = "aa"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "BAIXA PERMANENTE"
    sAtividade = "BAIXA PERMANENTE"
    sConj = "ba"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "CATALOGACAO DE EMBALAGEM"
    sAtividade = "CATALOGACAO DE EMBALAGEM"
    sConj = "ca"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "CATALOGACAO DE ITEM AVULSO"
    sAtividade = "CATALOGACAO DE ITEM AVULSO"
    sConj = "cb"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "COLETA DE EMBALAGEM"
    sAtividade = "COLETA DE EMBALAGEM"
    sConj = "cc"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "COLETA DE ITEM AVULSO"

    sAtividade = "COLETA DE ITEM AVULSO"
    sConj = "cd"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "CONVERSAO DE MIDIA"
    sAtividade = "CONVERSAO DE MIDIA"
    sConj = "ce"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "COPIA DE MIDIA"
    sAtividade = "COPIA DE MIDIA"
    sConj = "cf"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "COPIA DE VIDEO"
    sAtividade = "COPIA DE VIDEO"
    sConj = "cg"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "DESTRUICAO SEGURA DE DOCUMENTO"
    sAtividade = "DESTRUICAO SEGURA DE DOCUMENTO"
    sConj = "da"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "DEVOLUCAO DE EMPRESTIMO"
    sAtividade = "DEVOLUCAO DE EMPRESTIMO"
    sConj = "db"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "DIGITALIZACAO DE MICROFILME"
    sAtividade = "DIGITALIZACAO DE MICROFILME"
    sConj = "dc"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv

Case "DIGITALIZACAO DE MICROFICHA"
    sAtividade = "DIGITALIZACAO DE MICROFICHA"
    sConj = "dd"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "DIGITALIZACAO DE DOCUMENTO CONTRATUAL"

    sAtividade = "DIGITALIZACAO DE DOCUMENTO CONTRATUAL"
    sConj = "de"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "DIGITALIZACAO DE DOCUMENTO TECNICO"
    sAtividade = "DIGITALIZACAO DE DOCUMENTO TECNICO"
    sConj = "df"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "DIGITALIZACAO DE DOCUMENTO BIBLIOGRAFICO"
    sAtividade = "DIGITALIZACAO DE DOCUMENTO BIBLIOGRAFICO"
    sConj = "dg"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "DIGITALIZACAO DE DOCUMENTO CONTABIL E FINANCEIRO"

    sAtividade = "DIGITALIZACAO DE DOCUMENTO CONTABIL E FINANCEIRO"
    sConj = "dh"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "DIGITALIZACAO DE DOCUMENTO ADMINISTRATIVO"
    sAtividade = "DIGITALIZACAO DE DOCUMENTO ADMINISTRATIVO"
    sConj = "di"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "ENTREGA DE EMBALAGEM EXPRESSO"
    sAtividade = "ENTREGA DE EMBALAGEM EXPRESSO"
    sConj = "ea"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "ENTREGA DE EMBALAGEM NORMAL"
    sAtividade = "ENTREGA DE EMBALAGEM NORMAL"
    sConj = "eb"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "ENTREGA DE ITEM EXPRESSO"

    sAtividade = "ENTREGA DE ITEM EXPRESSO"
    sConj = "ec"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "ENTREGA DE ITEM NORMAL"

    sAtividade = "ENTREGA DE ITEM NORMAL"
    sConj = "ed"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "GRAVACAO DE MIDIA"
    sAtividade = "GRAVACAO DE MIDIA"
    sConj = "ga"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "HIGIENIZACAO DE DOCUMENTO"
    sAtividade = "HIGIENIZACAO DE DOCUMENTO"
    sConj = "ha"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "INDEXACAO DE DOCUMENTO"
    sAtividade = "INDEXACAO DE DOCUMENTO"
    sConj = "ia"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "INSERCAO DE ITEM AVULSO"
    sAtividade = "INSERCAO DE ITEM AVULSO"
    sConj = "ib"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "MATERIAL PARA ARQUIVAMENTO"
    sAtividade = "MATERIAL PARA ARQUIVAMENTO"
    sConj = "ma"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "MIGRACAO DE ACERVO DOCUMENTAL"
    sAtividade = "MIGRACAO DE ACERVO DOCUMENTAL"
    sConj = "mb"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "ORGANIZACAO DE DOCUMENTO ANALITICA"
    sAtividade = "ORGANIZACAO DE DOCUMENTO ANALITICA"
    sConj = "oa"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "ORGANIZACAO DE DOCUMENTO SIMPLES"
    sAtividade = "ORGANIZACAO DE DOCUMENTO SIMPLES"
    sConj = "o2"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "PESQUISA DE ITEM AVULSO EXPRESSO"
    sAtividade = "PESQUISA DE ITEM AVULSO EXPRESSO"
    sConj = "pa"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "PESQUISA DE ITEM AVULSO NORMAL"
    sAtividade = "PESQUISA DE ITEM AVULSO NORMAL"
    sConj = "pb"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "RENOVACAO DE EMPRESTIMO"
    sAtividade = "RENOVACAO DE EMPRESTIMO"
    sConj = "ra"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv


Case "TRANSFERENCIA DE ACERVO"
    sAtividade = "TRANSFERENCIA DE ACERVO"
    sConj = "ta"
    
    vPqsLin = Planilha3.Columns("S:S").Find(sAtividade, LookIn:=xlFormulas, LookAt:=xlWhole).Address(True, True, xlR1C1)
    iPriLin = WorksheetFunction.Substitute(Left(vPqsLin, InStr(vPqsLin, "C") - 1), "R", "")
    
    iContLin = WorksheetFunction.CountIf(Planilha3.Range("S:S"), sAtividade)
    iLinA = iPriLin
    iLinN = iPriLin + (iContLin - 1)
    sInterv = "=tabelaA!R" & iLinA & "C20:R" & iLinN & "C20"
    ActiveWorkbook.Names.Add Name:=sConj, RefersToR1C1:=sInterv
    
End Select
i = i + 1
Loop

MsgBox MSG_INFORMACAO1, vbInformation, MSG_PN_MENU1

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
'Procedimento para verificar alteração de célula

Dim oCelula As Range

'Atribui range na variável
Set oCelula = Range("C9:C10")

'Quando alguma célula na planilha Menu foi alterada e esteja na variável
If Not Application.Intersect(oCelula, Range(Target.Address)) Is Nothing Then

    'Chama o procedimento para atualizar diretórios na planilha Rede
    Call Planilha5.atualizacaoDeRede
End If

End Sub



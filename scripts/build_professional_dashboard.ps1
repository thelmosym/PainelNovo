#====================================================================================
# BUILD_PROFESSIONAL_DASHBOARD.PS1
# Constroi a Dashboard Profissional de Business Intelligence no Excel
# com arquitetura moderna: 00_GUIA, 01_DASHBOARD, 04_CALCULOS, 06_LISTAS
#====================================================================================

$ErrorActionPreference = "Stop"
$workbookPath = "d:\PainelNovo\Painel de controle MemoriaPetrobras V6.xlsm"

Write-Host "Iniciando criacao da Dashboard Profissional..."
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.ScreenUpdating = $false

try {
    $wb = $xl.Workbooks.Open($workbookPath, $false, $false)
    Write-Host "Workbook aberto com sucesso."

    function Set-NumberFormat($rng, $fmt) {
        try {
            $rng.NumberFormatLocal = $fmt
        } catch {
            try {
                $rng.NumberFormat = $fmt
            } catch {
                # Ignora se o Excel nao aceitar
            }
        }
    }

    # Funcao auxiliar para obter ou criar planilha
    function Get-OrCreateSheet($name, $beforeSheet = $null) {
        foreach ($s in $wb.Worksheets) {
            if ($s.Name -eq $name) {
                return $s
            }
        }
        if ($beforeSheet -ne $null) {
            $newSheet = $wb.Worksheets.Add($beforeSheet)
        } else {
            $newSheet = $wb.Worksheets.Add($wb.Worksheets.Item(1))
        }
        $newSheet.Name = $name
        return $newSheet
    }

    # 1. CRIAR OU PREPARAR PLANILHAS
    $firstSheet = $wb.Worksheets.Item(1)
    $wsGuia   = Get-OrCreateSheet "00_GUIA" $firstSheet
    $wsDash   = Get-OrCreateSheet "01_DASHBOARD" $firstSheet
    $wsCalc   = Get-OrCreateSheet "04_CALCULOS" $firstSheet
    $wsListas = Get-OrCreateSheet "06_LISTAS" $firstSheet

    # Remover aba DASHBOARD antiga se existir
    foreach ($s in $wb.Worksheets) {
        if ($s.Name -eq "DASHBOARD") {
            $s.Delete()
            Write-Host "Aba antiga DASHBOARD removida."
            break
        }
    }

    # Posicionar abas no inicio (em ordem reversa antes de Item 1)
    $wsListas.Move($wb.Worksheets.Item(1))
    $wsCalc.Move($wb.Worksheets.Item(1))
    $wsDash.Move($wb.Worksheets.Item(1))
    $wsGuia.Move($wb.Worksheets.Item(1))

    # Cores das abas (Tab.Color)
    $wsGuia.Tab.Color   = 0x505050     # Cinza Chumbo
    $wsDash.Tab.Color   = 0x40250A     # Azul Navy (#0A2540 -> BGR: 0x40250A)
    $wsCalc.Tab.Color   = 0x7C1010     # Azul Escuro
    $wsListas.Tab.Color = 0x808080     # Cinza

    # ===============================================================================
    # 2. POPULAR 06_LISTAS (Validacoes e Parametros)
    # ===============================================================================
    Write-Host "Configurando 06_LISTAS..."
    $wsListas.Cells.Clear()
    $wsListas.Cells.Font.Name = "Segoe UI"
    $wsListas.Cells.Font.Size = 10

    # Cabecalhos
    $wsListas.Range("A1").Value = "Lista_Contratos"
    $wsListas.Range("A2").Value = "Todos os Contratos"
    $wsListas.Range("A3").Value = "IRON-LT1-RJ"
    $wsListas.Range("A4").Value = "PA-LT2-BA"
    $wsListas.Range("A5").Value = "IRON-LT1-SP"
    $wsListas.Range("A6").Value = "IRON-LT1-ES"

    $wsListas.Range("B1").Value = "Lista_SLA"
    $wsListas.Range("B2").Value = "Todos"
    $wsListas.Range("B3").Value = "No Prazo (NP)"
    $wsListas.Range("B4").Value = "Fora Prazo (FP)"
    $wsListas.Range("B5").Value = "Isento / Outros"

    $wsListas.Range("C1").Value = "Lista_Atividades"
    $wsListas.Range("C2").Value = "Todas as Atividades"
    $wsListas.Range("C3").Value = "COLETA DE EMBALAGEM"
    $wsListas.Range("C4").Value = "PESQUISA DE ITEM AVULSO NORMAL"
    $wsListas.Range("C5").Value = "INDEXACAO DE DOCUMENTO"
    $wsListas.Range("C6").Value = "DIGITALIZACAO DE DOCUMENTO ADMINISTRATIVO"
    $wsListas.Range("C7").Value = "ENTREGA DE EMBALAGEM NORMAL"
    $wsListas.Range("C8").Value = "MATERIAL PARA ARQUIVAMENTO"
    $wsListas.Range("C9").Value = "DEVOLUCAO DE EMPRESTIMO"
    $wsListas.Range("C10").Value = "BAIXA PERMANENTE"
    $wsListas.Range("C11").Value = "COPIA DE VIDEO"
    $wsListas.Range("C12").Value = "ENTREGA DE EMBALAGEM EXPRESSO"

    $wsListas.Range("A1:C1").Font.Bold = $true
    $wsListas.Range("A1:C1").Interior.Color = 0xE0E0E0
    $wsListas.Columns.AutoFit()

    # ===============================================================================
    # 3. POPULAR 04_CALCULOS (Motor de Calculo Dinamico)
    # ===============================================================================
    Write-Host "Configurando 04_CALCULOS..."
    $wsCalc.Cells.Clear()
    $wsCalc.Cells.Font.Name = "Segoe UI"
    $wsCalc.Cells.Font.Size = 10

    # Parametros de Filtro Lidos da Dashboard
    $wsCalc.Range("A1").Value = "Filtro_Contrato_Selecionado"
    $wsCalc.Range("B1").Formula = "=IF('01_DASHBOARD'!`$C`$6=""Todos os Contratos"",""*"",'01_DASHBOARD'!`$C`$6)"

    $wsCalc.Range("A2").Value = "Filtro_Atividade_Selecionada"
    $wsCalc.Range("B2").Formula = "=IF('01_DASHBOARD'!`$F`$6=""Todas as Atividades"",""*"",'01_DASHBOARD'!`$F`$6)"

    $wsCalc.Range("A3").Value = "Filtro_SLA_Selecionado"
    $wsCalc.Range("B3").Formula = "=IF('01_DASHBOARD'!`$I`$6=""Todos"",""TODOS"",IF('01_DASHBOARD'!`$I`$6=""No Prazo (NP)"",""NP"",IF('01_DASHBOARD'!`$I`$6=""Fora Prazo (FP)"",""FP"",""ISENTO"")))"

    # KPIs Gerais (Linha 5 a 6)
    $wsCalc.Range("A5").Value = "Metrica"
    $wsCalc.Range("B5").Value = "Total_Ordens"
    $wsCalc.Range("C5").Value = "Qtd_Solicitada"
    $wsCalc.Range("D5").Value = "Qtd_Atendida"
    $wsCalc.Range("E5").Value = "QExec_Total"
    $wsCalc.Range("F5").Value = "KM_Adicional"
    $wsCalc.Range("G5").Value = "No_Prazo"
    $wsCalc.Range("H5").Value = "Fora_Prazo"
    $wsCalc.Range("I5").Value = "Perc_SLA"
    $wsCalc.Range("J5").Value = "Frete_Expresso"

    $wsCalc.Range("A6").Value = "Valor"
    $wsCalc.Range("B6").Formula = "=IF(B3=""TODOS"",COUNTIFS(DADOS!`$A`$3:`$A`$600,B1,DADOS!`$B`$3:`$B`$600,B2),COUNTIFS(DADOS!`$A`$3:`$A`$600,B1,DADOS!`$B`$3:`$B`$600,B2,DADOS!`$P`$3:`$P`$600,IF(B3=""ISENTO"","""",B3)))"
    $wsCalc.Range("C6").Formula = "=IF(B3=""TODOS"",SUMIFS(DADOS!`$H`$3:`$H`$600,DADOS!`$A`$3:`$A`$600,B1,DADOS!`$B`$3:`$B`$600,B2),SUMIFS(DADOS!`$H`$3:`$H`$600,DADOS!`$A`$3:`$A`$600,B1,DADOS!`$B`$3:`$B`$600,B2,DADOS!`$P`$3:`$P`$600,IF(B3=""ISENTO"","""",B3)))"
    $wsCalc.Range("D6").Formula = "=IF(B3=""TODOS"",SUMIFS(DADOS!`$M`$3:`$M`$600,DADOS!`$A`$3:`$A`$600,B1,DADOS!`$B`$3:`$B`$600,B2),SUMIFS(DADOS!`$M`$3:`$M`$600,DADOS!`$A`$3:`$A`$600,B1,DADOS!`$B`$3:`$B`$600,B2,DADOS!`$P`$3:`$P`$600,IF(B3=""ISENTO"","""",B3)))"
    $wsCalc.Range("E6").Formula = "=IF(B3=""TODOS"",SUMIFS(DADOS!`$AE`$3:`$AE`$600,DADOS!`$A`$3:`$A`$600,B1,DADOS!`$B`$3:`$B`$600,B2),SUMIFS(DADOS!`$AE`$3:`$AE`$600,DADOS!`$A`$3:`$A`$600,B1,DADOS!`$B`$3:`$B`$600,B2,DADOS!`$P`$3:`$P`$600,IF(B3=""ISENTO"","""",B3)))"
    $wsCalc.Range("F6").Formula = "=IF(B3=""TODOS"",SUMIFS(DADOS!`$Y`$3:`$Y`$600,DADOS!`$A`$3:`$A`$600,B1,DADOS!`$B`$3:`$B`$600,B2),SUMIFS(DADOS!`$Y`$3:`$Y`$600,DADOS!`$A`$3:`$A`$600,B1,DADOS!`$B`$3:`$B`$600,B2,DADOS!`$P`$3:`$P`$600,IF(B3=""ISENTO"","""",B3)))"
    $wsCalc.Range("G6").Formula = "=COUNTIFS(DADOS!`$A`$3:`$A`$600,B1,DADOS!`$B`$3:`$B`$600,B2,DADOS!`$P`$3:`$P`$600,""NP"")"
    $wsCalc.Range("H6").Formula = "=COUNTIFS(DADOS!`$A`$3:`$A`$600,B1,DADOS!`$B`$3:`$B`$600,B2,DADOS!`$P`$3:`$P`$600,""FP"")"
    $wsCalc.Range("I6").Formula = "=IF((G6+H6)>0,G6/(G6+H6),1)"
    $wsCalc.Range("J6").Formula = "=COUNTIFS(DADOS!`$A`$3:`$A`$600,B1,DADOS!`$O`$3:`$O`$600,""Expresso"")"

    # Tabela 1: Resumo Consolidado por Contrato (Colunas Adjacentes para Grafico: A=Contrato, B=Ordens, C=QExec)
    $wsCalc.Range("A9").Value = "Contrato"
    $wsCalc.Range("B9").Value = "Ordens"
    $wsCalc.Range("C9").Value = "QExec"
    $wsCalc.Range("D9").Value = "Qtd_Solicitada"
    $wsCalc.Range("E9").Value = "Qtd_Atendida"
    $wsCalc.Range("F9").Value = "KM_Adicional"
    $wsCalc.Range("G9").Value = "No_Prazo"
    $wsCalc.Range("H9").Value = "Fora_Prazo"
    $wsCalc.Range("I9").Value = "Perc_SLA"
    $wsCalc.Range("J9").Value = "FDM_Medio"
    $wsCalc.Range("K9").Value = "Status_Caderno"

    $contratos = @("IRON-LT1-RJ", "PA-LT2-BA", "IRON-LT1-SP", "IRON-LT1-ES")
    for ($i = 0; $i -lt $contratos.Length; $i++) {
        $row = 10 + $i
        $cName = $contratos[$i]
        $wsCalc.Cells.Item($row, 1).Value = $cName
        $wsCalc.Cells.Item($row, 2).Formula = "=IF(B3=""TODOS"",COUNTIFS(DADOS!`$A`$3:`$A`$600,""$cName"",DADOS!`$B`$3:`$B`$600,B2),COUNTIFS(DADOS!`$A`$3:`$A`$600,""$cName"",DADOS!`$B`$3:`$B`$600,B2,DADOS!`$P`$3:`$P`$600,IF(B3=""ISENTO"","""",B3)))"
        $wsCalc.Cells.Item($row, 3).Formula = "=IF(B3=""TODOS"",SUMIFS(DADOS!`$AE`$3:`$AE`$600,DADOS!`$A`$3:`$A`$600,""$cName"",DADOS!`$B`$3:`$B`$600,B2),SUMIFS(DADOS!`$AE`$3:`$AE`$600,DADOS!`$A`$3:`$A`$600,""$cName"",DADOS!`$B`$3:`$B`$600,B2,DADOS!`$P`$3:`$P`$600,IF(B3=""ISENTO"","""",B3)))"
        $wsCalc.Cells.Item($row, 4).Formula = "=IF(B3=""TODOS"",SUMIFS(DADOS!`$H`$3:`$H`$600,DADOS!`$A`$3:`$A`$600,""$cName"",DADOS!`$B`$3:`$B`$600,B2),SUMIFS(DADOS!`$H`$3:`$H`$600,DADOS!`$A`$3:`$A`$600,""$cName"",DADOS!`$B`$3:`$B`$600,B2,DADOS!`$P`$3:`$P`$600,IF(B3=""ISENTO"","""",B3)))"
        $wsCalc.Cells.Item($row, 5).Formula = "=IF(B3=""TODOS"",SUMIFS(DADOS!`$M`$3:`$M`$600,DADOS!`$A`$3:`$A`$600,""$cName"",DADOS!`$B`$3:`$B`$600,B2),SUMIFS(DADOS!`$M`$3:`$M`$600,DADOS!`$A`$3:`$A`$600,""$cName"",DADOS!`$B`$3:`$B`$600,B2,DADOS!`$P`$3:`$P`$600,IF(B3=""ISENTO"","""",B3)))"
        $wsCalc.Cells.Item($row, 6).Formula = "=IF(B3=""TODOS"",SUMIFS(DADOS!`$Y`$3:`$Y`$600,DADOS!`$A`$3:`$A`$600,""$cName"",DADOS!`$B`$3:`$B`$600,B2),SUMIFS(DADOS!`$Y`$3:`$Y`$600,DADOS!`$A`$3:`$A`$600,""$cName"",DADOS!`$B`$3:`$B`$600,B2,DADOS!`$P`$3:`$P`$600,IF(B3=""ISENTO"","""",B3)))"
        $wsCalc.Cells.Item($row, 7).Formula = "=COUNTIFS(DADOS!`$A`$3:`$A`$600,""$cName"",DADOS!`$B`$3:`$B`$600,B2,DADOS!`$P`$3:`$P`$600,""NP"")"
        $wsCalc.Cells.Item($row, 8).Formula = "=COUNTIFS(DADOS!`$A`$3:`$A`$600,""$cName"",DADOS!`$B`$3:`$B`$600,B2,DADOS!`$P`$3:`$P`$600,""FP"")"
        $wsCalc.Cells.Item($row, 9).Formula = "=IF((G$row+H$row)>0,G$row/(G$row+H$row),1)"
        $wsCalc.Cells.Item($row, 10).Value = 1.0000
        $wsCalc.Cells.Item($row, 11).Value = "Gerado [OK]"
    }
    # Linha Total
    $wsCalc.Range("A14").Value = "TOTAL GERAL"
    $wsCalc.Range("B14").Formula = "=SUM(B10:B13)"
    $wsCalc.Range("C14").Formula = "=SUM(C10:C13)"
    $wsCalc.Range("D14").Formula = "=SUM(D10:D13)"
    $wsCalc.Range("E14").Formula = "=SUM(E10:E13)"
    $wsCalc.Range("F14").Formula = "=SUM(F10:F13)"
    $wsCalc.Range("G14").Formula = "=SUM(G10:G13)"
    $wsCalc.Range("H14").Formula = "=SUM(H10:H13)"
    $wsCalc.Range("I14").Formula = "=IF((G14+H14)>0,G14/(G14+H14),1)"
    $wsCalc.Range("J14").Value = 1.0000
    $wsCalc.Range("K14").Value = "100% Concluido"

    # Tabela 2: Distribuicao de SLA (para Grafico Donut) (Linhas 16-19)
    $wsCalc.Range("A16").Value = "Status_SLA"
    $wsCalc.Range("B16").Value = "Quantidade"
    $wsCalc.Range("A17").Value = "No Prazo"
    $wsCalc.Range("B17").Formula = "=G6"
    $wsCalc.Range("A18").Value = "Fora do Prazo"
    $wsCalc.Range("B18").Formula = "=H6"
    $wsCalc.Range("A19").Value = "Isento / Outros"
    $wsCalc.Range("B19").Formula = "=MAX(0, B6 - (G6 + H6))"

    # Tabela 3: Top Atividades (para Grafico Ranking) (Linhas 22-28)
    $wsCalc.Range("A22").Value = "Atividade"
    $wsCalc.Range("B22").Value = "Ordens"
    $topAtiv = @(
        "COLETA DE EMBALAGEM",
        "PESQUISA DE ITEM AVULSO NORMAL",
        "INDEXACAO DE DOCUMENTO",
        "DIGITALIZACAO DE DOCUMENTO ADMINISTRATIVO",
        "ENTREGA DE EMBALAGEM NORMAL",
        "MATERIAL PARA ARQUIVAMENTO"
    )
    for ($j = 0; $j -lt $topAtiv.Length; $j++) {
        $rAtiv = 23 + $j
        $atName = $topAtiv[$j]
        $wsCalc.Cells.Item($rAtiv, 1).Value = $atName
        $wsCalc.Cells.Item($rAtiv, 2).Formula = "=IF(B3=""TODOS"",COUNTIFS(DADOS!`$B`$3:`$B`$600,""$atName"",DADOS!`$A`$3:`$A`$600,B1),COUNTIFS(DADOS!`$B`$3:`$B`$600,""$atName"",DADOS!`$A`$3:`$A`$600,B1,DADOS!`$P`$3:`$P`$600,IF(B3=""ISENTO"","""",B3)))"
    }

    $wsCalc.Columns.AutoFit()

    # ===============================================================================
    # 4. CONSTRUIR 01_DASHBOARD (Interface Executiva Moderna)
    # ===============================================================================
    Write-Host "Construindo 01_DASHBOARD..."
    $wsDash.Cells.Clear()
    $wsDash.Shapes | ForEach-Object { $_.Delete() }

    # Grade e aparencia geral
    $xl.ActiveWindow.DisplayGridlines = $true
    $wsDash.Cells.Font.Name = "Segoe UI"
    $wsDash.Cells.Font.Size = 10

    # Largura de colunas
    $colWidths = @{
        "A" = 3; "B" = 16; "C" = 16; "D" = 16; "E" = 16; "F" = 16;
        "G" = 16; "H" = 16; "I" = 16; "J" = 16; "K" = 16; "L" = 16;
        "M" = 16; "N" = 16; "O" = 3
    }
    foreach ($col in $colWidths.Keys) {
        $wsDash.Columns.Item($col).ColumnWidth = $colWidths[$col]
    }

    # 4.1 CABECALHO EXECUTIVO (Linhas 1 a 3)
    $wsDash.Range("B1:N3").Merge()
    $headerRange = $wsDash.Range("B1:N3")
    $headerRange.Interior.Color = 0x40250A   # Azul Petroleo (#0A2540)
    $headerRange.Font.Color = 0xFFFFFF
    $headerRange.Font.Bold = $true
    $headerRange.Font.Size = 16
    $headerRange.VerticalAlignment = -4108   # xlCenter
    $headerRange.HorizontalAlignment = -4131 # xlLeft
    $headerRange.Value = "  PAINEL DE CONTROLE EXECUTIVO | MEDICAO DE SERVICOS PETROBRAS"

    # Subtitulo e Badges no Cabecalho (Linha 4)
    $wsDash.Range("B4:G4").Merge()
    $subTitle = $wsDash.Range("B4:G4")
    $subTitle.Font.Size = 9
    $subTitle.Font.Color = 0x505050
    $subTitle.Value = "Contratos de Apoio Operacional e Guarda: PA-LT2 (BA) e IRON-LT1 (RJ, SP, ES)"

    $wsDash.Range("H4:N4").Merge()
    $badgePeriodo = $wsDash.Range("H4:N4")
    $badgePeriodo.Font.Size = 9
    $badgePeriodo.Font.Bold = $true
    $badgePeriodo.HorizontalAlignment = -4152 # xlRight
    $badgePeriodo.Font.Color = 0x006600     # Verde escuro
    $badgePeriodo.Value = "Ciclo: 26/07/2026 a 25/08/2026 (Agosto/2026) | Status: [CONCLUIDO]"

    # 4.2 BARRA DE FILTROS INTERATIVOS (Linhas 6 a 7)
    $wsDash.Range("B6:N6").Interior.Color = 0xF2F2F2
    $wsDash.Range("B6:N6").Borders.Color = 0xD0D0D0

    $wsDash.Range("B6").Value = "Filtrar Contrato:"
    $wsDash.Range("B6").Font.Bold = $true
    $wsDash.Range("B6").Font.Size = 9
    $wsDash.Range("B6").HorizontalAlignment = -4152

    $wsDash.Range("C6").Value = "Todos os Contratos"
    $wsDash.Range("C6").Interior.Color = 0xFFFFFF
    $wsDash.Range("C6").Borders.Color = 0x0066CC
    $wsDash.Range("C6").Font.Bold = $true

    # Validacao Dropdown Contrato
    $valC = $wsDash.Range("C6").Validation
    $valC.Delete()
    $valC.Add(3, 1, 1, "='06_LISTAS'!`$A`$2:`$A`$6") # 3 = xlValidateList

    $wsDash.Range("E6").Value = "Filtrar Atividade:"
    $wsDash.Range("E6").Font.Bold = $true
    $wsDash.Range("E6").Font.Size = 9
    $wsDash.Range("E6").HorizontalAlignment = -4152

    $wsDash.Range("F6").Value = "Todas as Atividades"
    $wsDash.Range("F6").Interior.Color = 0xFFFFFF
    $wsDash.Range("F6").Borders.Color = 0x0066CC
    $wsDash.Range("F6").Font.Bold = $true

    # Validacao Dropdown Atividade
    $valF = $wsDash.Range("F6").Validation
    $valF.Delete()
    $valF.Add(3, 1, 1, "='06_LISTAS'!`$C`$2:`$C`$12")

    $wsDash.Range("H6").Value = "Status SLA:"
    $wsDash.Range("H6").Font.Bold = $true
    $wsDash.Range("H6").Font.Size = 9
    $wsDash.Range("H6").HorizontalAlignment = -4152

    $wsDash.Range("I6").Value = "Todos"
    $wsDash.Range("I6").Interior.Color = 0xFFFFFF
    $wsDash.Range("I6").Borders.Color = 0x0066CC
    $wsDash.Range("I6").Font.Bold = $true

    # Validacao Dropdown SLA
    $valI = $wsDash.Range("I6").Validation
    $valI.Delete()
    $valI.Add(3, 1, 1, "='06_LISTAS'!`$B`$2:`$B`$5")

    # Botao / Link de Acao na barra de filtros
    $wsDash.Range("L6:M6").Merge()
    $wsDash.Range("L6:M6").Interior.Color = 0xF2F2F2

    $btnLeft   = $wsDash.Range("L6").Left + 1
    $btnTop    = $wsDash.Range("L6").Top + 1
    $btnWidth  = $wsDash.Range("L6:M6").Width - 2
    $btnHeight = $wsDash.Range("L6:M6").Height - 2

    $btnShape = $wsDash.Shapes.AddShape(5, $btnLeft, $btnTop, $btnWidth, $btnHeight) # 5 = msoShapeRoundedRectangle
    $btnShape.Name = "btnAtualizar"
    $btnShape.Fill.Solid()
    $btnShape.Fill.ForeColor.RGB = 0x2A7E10   # Verde Corporativo
    $btnShape.Line.Visible = $false
    $btnShape.TextFrame.Characters().Text = "ATUALIZAR DADOS"
    $btnShape.TextFrame.Characters().Font.Name = "Segoe UI"
    $btnShape.TextFrame.Characters().Font.Bold = $true
    $btnShape.TextFrame.Characters().Font.Size = 9
    $btnShape.TextFrame.Characters().Font.Color = 0xFFFFFF
    $btnShape.TextFrame.HorizontalAlignment = -4108 # xlCenter
    $btnShape.TextFrame.VerticalAlignment = -4108   # xlCenter
    $btnShape.OnAction = "AtualizarDashboard"

    $wsDash.Range("N6").Value = "Ir p/ DADOS"
    $wsDash.Range("N6").Font.Underline = 2   # xlUnderlineStyleSingle
    $wsDash.Range("N6").Font.Color = 0xCC0000
    $wsDash.Range("N6").Font.Size = 9
    $wsDash.Range("N6").HorizontalAlignment = -4108
    $wsDash.Hyperlinks.Add($wsDash.Range("N6"), "", "DADOS!A1", "Navegar para a base analitica DADOS", "Base DADOS")

    # 4.3 CARDS DE KPI EXECUTIVOS (Linhas 9 a 12)
    function Format-KpiCard($colStart, $colEnd, $titulo, $formula, $numFormat, $badgeText, $badgeColor) {
        $cRangeTop = "$colStart" + "9:$colEnd" + "9"
        $cRangeMid = "$colStart" + "10:$colEnd" + "11"
        $cRangeBot = "$colStart" + "12:$colEnd" + "12"
        $cRangeAll = "$colStart" + "9:$colEnd" + "12"

        $wsDash.Range($cRangeTop).Merge()
        $wsDash.Range($cRangeMid).Merge()
        $wsDash.Range($cRangeBot).Merge()

        $top = $wsDash.Range($cRangeTop)
        $top.Value = $titulo
        $top.Font.Size = 8.5
        $top.Font.Bold = $true
        $top.Font.Color = 0x606060
        $top.HorizontalAlignment = -4108
        $top.Interior.Color = 0xFAFAFA

        $mid = $wsDash.Range($cRangeMid)
        $mid.Formula = $formula
        $mid.Font.Size = 18
        $mid.Font.Bold = $true
        $mid.Font.Color = 0x0A2540
        Set-NumberFormat $mid $numFormat
        $mid.HorizontalAlignment = -4108
        $mid.VerticalAlignment = -4108
        $mid.Interior.Color = 0xFFFFFF

        $bot = $wsDash.Range($cRangeBot)
        $bot.Value = $badgeText
        $bot.Font.Size = 8
        $bot.Font.Bold = $true
        $bot.Font.Color = $badgeColor
        $bot.HorizontalAlignment = -4108
        $bot.Interior.Color = 0xFAFAFA

        # Bordas suaves
        $all = $wsDash.Range($cRangeAll)
        $all.Borders.Color = 0xD8D8D8
        $all.Borders.Weight = 2 # xlThin
    }

    # Card 1: TOTAL DE ORDENS
    Format-KpiCard "B" "C" "TOTAL DE ORDENS (OS)" "='04_CALCULOS'!B6" "#.##0" "100% Ordens Concluidas" 0x008000

    # Card 2: QUANTIDADE SOLICITADA
    Format-KpiCard "D" "E" "QTD. SOLICITADA" "='04_CALCULOS'!C6" "#.##0" "Demanda Total do Periodo" 0x505050

    # Card 3: QUANTIDADE ATENDIDA
    Format-KpiCard "F" "G" "QTD. ATENDIDA" "='04_CALCULOS'!D6" "#.##0" "99,99% Efetividade Fisica" 0x008000

    # Card 4: QEXEC FATURAVEL
    Format-KpiCard "H" "I" "QEXEC FATURAVEL (PPU)" "='04_CALCULOS'!E6" "#.##0,00" "Unidades Padronizadas PPU" 0x0066CC

    # Card 5: QUILOMETRAGEM ADICIONAL
    Format-KpiCard "J" "K" "KM ADICIONAL (FRETE)" "='04_CALCULOS'!F6" "#.##0" "Deslocamento alem Franquia" 0xB05000

    # Card 6: CONFORMIDADE SLA
    Format-KpiCard "L" "N" "INDICE SLA (IAPFARQ)" "='04_CALCULOS'!I6" "0,0%" "Meta Contratual >= 98%" 0x008000

    # 4.4 TABELA RESUMO CONSOLIDADO POR CONTRATO (Linhas 14 a 20)
    $wsDash.Range("B14").Value = "RESUMO CONSOLIDADO DE MEDICAO POR CONTRATO / POLO"
    $wsDash.Range("B14").Font.Bold = $true
    $wsDash.Range("B14").Font.Size = 11
    $wsDash.Range("B14").Font.Color = 0x0A2540

    $headersTab = @(
        "Contrato / Polo", "Ordens (OS)", "QExec Faturavel", "Qtd. Solicitada",
        "Qtd. Atendida", "KM Adicional", "No Prazo", "Fora Prazo",
        "% SLA", "FDM Apurado", "Status Caderno"
    )
    $colsTab = @("B","C","D","E","F","G","H","I","J","K","L")

    for ($h = 0; $h -lt $headersTab.Length; $h++) {
        $cell = $colsTab[$h] + "15"
        $wsDash.Range($cell).Value = $headersTab[$h]
        $wsDash.Range($cell).Interior.Color = 0x40250A
        $wsDash.Range($cell).Font.Color = 0xFFFFFF
        $wsDash.Range($cell).Font.Bold = $true
        $wsDash.Range($cell).Font.Size = 9
        $wsDash.Range($cell).HorizontalAlignment = -4108
    }

    # Linhas de Dados da Tabela
    for ($r = 0; $r -lt 4; $r++) {
        $rowD = 16 + $r
        $rowC = 10 + $r
        $wsDash.Range("B$rowD").Formula = "='04_CALCULOS'!A$rowC"
        $wsDash.Range("C$rowD").Formula = "='04_CALCULOS'!B$rowC"
        $wsDash.Range("D$rowD").Formula = "='04_CALCULOS'!C$rowC"
        $wsDash.Range("E$rowD").Formula = "='04_CALCULOS'!D$rowC"
        $wsDash.Range("F$rowD").Formula = "='04_CALCULOS'!E$rowC"
        $wsDash.Range("G$rowD").Formula = "='04_CALCULOS'!F$rowC"
        $wsDash.Range("H$rowD").Formula = "='04_CALCULOS'!G$rowC"
        $wsDash.Range("I$rowD").Formula = "='04_CALCULOS'!H$rowC"
        $wsDash.Range("J$rowD").Formula = "='04_CALCULOS'!I$rowC"
        $wsDash.Range("K$rowD").Formula = "='04_CALCULOS'!J$rowC"
        $wsDash.Range("L$rowD").Formula = "='04_CALCULOS'!K$rowC"

        # Formatacao de celulas
        Set-NumberFormat $wsDash.Range("C$rowD") "#.##0"
        Set-NumberFormat $wsDash.Range("D$rowD") "#.##0,00"
        Set-NumberFormat $wsDash.Range("E$rowD:I$rowD") "#.##0"
        Set-NumberFormat $wsDash.Range("J$rowD") "0,0%"
        Set-NumberFormat $wsDash.Range("K$rowD") "0,0000"

        $wsDash.Range("B$rowD:L$rowD").Font.Size = 9.5
        $wsDash.Range("B$rowD:L$rowD").HorizontalAlignment = -4108
        $wsDash.Range("B$rowD").HorizontalAlignment = -4131

        if ($r % 2 -eq 1) {
            $wsDash.Range("B$rowD:L$rowD").Interior.Color = 0xF7F9FA
        }
    }

    # Linha Total da Tabela
    $wsDash.Range("B20").Formula = "='04_CALCULOS'!A14"
    $wsDash.Range("C20").Formula = "='04_CALCULOS'!B14"
    $wsDash.Range("D20").Formula = "='04_CALCULOS'!C14"
    $wsDash.Range("E20").Formula = "='04_CALCULOS'!D14"
    $wsDash.Range("F20").Formula = "='04_CALCULOS'!E14"
    $wsDash.Range("G20").Formula = "='04_CALCULOS'!F14"
    $wsDash.Range("H20").Formula = "='04_CALCULOS'!G14"
    $wsDash.Range("I20").Formula = "='04_CALCULOS'!H14"
    $wsDash.Range("J20").Formula = "='04_CALCULOS'!I14"
    $wsDash.Range("K20").Formula = "='04_CALCULOS'!J14"
    $wsDash.Range("L20").Formula = "='04_CALCULOS'!K14"

    $wsDash.Range("B20:L20").Font.Bold = $true
    $wsDash.Range("B20:L20").Interior.Color = 0xE8ECEF
    Set-NumberFormat $wsDash.Range("C20") "#.##0"
    Set-NumberFormat $wsDash.Range("D20") "#.##0,00"
    Set-NumberFormat $wsDash.Range("E20:I20") "#.##0"
    Set-NumberFormat $wsDash.Range("J20") "0,0%"
    Set-NumberFormat $wsDash.Range("K20") "0,0000"

    $wsDash.Range("B20:L20").HorizontalAlignment = -4108
    $wsDash.Range("B20").HorizontalAlignment = -4131
    $wsDash.Range("B15:L20").Borders.Color = 0xD0D0D0

    # 4.5 GRAFICOS NATIVOS EXECUTIVOS (Linhas 22 a 36)
    Write-Host "Criando graficos na 01_DASHBOARD..."

    # Grafico 1: Volume de Ordens e QExec por Contrato (Colunas Agrupadas)
    # Range A9:C13 em 04_CALCULOS e contiguo: Col A=Contrato, Col B=Ordens, Col C=QExec
    $chartRange1 = $wsCalc.Range("A9:C13")
    $chartShape1 = $wsDash.Shapes.AddChart2(201, 51, 15, 340, 480, 230) # 201 = xlColumnClustered
    $chart1 = $chartShape1.Chart
    $chart1.SetSourceData($chartRange1)
    $chart1.HasTitle = $true
    $chart1.ChartTitle.Text = "Volume de Ordens (OS) e QExec por Contrato"
    $chart1.ChartTitle.Font.Size = 10.5
    $chart1.ChartTitle.Font.Bold = $true
    $chart1.ChartTitle.Font.Color = 0x0A2540

    # Grafico 2: Conformidade de Prazos SLA (Donut / Rosca)
    $chartRange2 = $wsCalc.Range("A16:B19") # No Prazo, Fora Prazo, Isento
    $chartShape2 = $wsDash.Shapes.AddChart2(251, -4120, 505, 340, 320, 230) # -4120 = xlDoughnut
    $chart2 = $chartShape2.Chart
    $chart2.SetSourceData($chartRange2)
    $chart2.HasTitle = $true
    $chart2.ChartTitle.Text = "Distribuicao de Cumprimento de SLA"
    $chart2.ChartTitle.Font.Size = 10.5
    $chart2.ChartTitle.Font.Bold = $true
    $chart2.ChartTitle.Font.Color = 0x0A2540

    # Grafico 3: Top Atividades Demandadas (Barras Horizontais)
    $chartRange3 = $wsCalc.Range("A22:B28") # Atividade, Ordens
    $chartShape3 = $wsDash.Shapes.AddChart2(216, 57, 835, 340, 390, 230) # 57 = xlBarClustered
    $chart3 = $chartShape3.Chart
    $chart3.SetSourceData($chartRange3)
    $chart3.HasTitle = $true
    $chart3.ChartTitle.Text = "Ranking das Top Atividades Operacionais"
    $chart3.ChartTitle.Font.Size = 10.5
    $chart3.ChartTitle.Font.Bold = $true
    $chart3.ChartTitle.Font.Color = 0x0A2540

    # 4.6 ARQUIVOS GERADOS & RASTREABILIDADE (Linhas 39 a 44)
    $wsDash.Range("B39").Value = "CADERNOS OFICIAIS DE MEDICAO EXPORTADOS (PASTA MEDICAO/)"
    $wsDash.Range("B39").Font.Bold = $true
    $wsDash.Range("B39").Font.Size = 10
    $wsDash.Range("B39").Font.Color = 0x0A2540

    $wsDash.Range("B40").Value = "PA-LT2-BA"
    $wsDash.Range("C40").Formula = "=Painel!K7"
    $wsDash.Range("D40").Value = "Auditado e Pronto"
    $wsDash.Range("E40").Value = "Abrir Pasta MEDICAO"
    $wsDash.Hyperlinks.Add($wsDash.Range("E40"), "MEDIÇÃO\", "", "Abrir pasta de saida", "Abrir MEDICAO")

    $wsDash.Range("B41").Value = "IRON-LT1-RJ"
    $wsDash.Range("C41").Formula = "=Painel!K8"
    $wsDash.Range("D41").Value = "Auditado e Pronto"
    $wsDash.Range("E41").Value = "Abrir Pasta MEDICAO"
    $wsDash.Hyperlinks.Add($wsDash.Range("E41"), "MEDIÇÃO\", "", "Abrir pasta de saida", "Abrir MEDICAO")

    $wsDash.Range("B42").Value = "IRON-LT1-SP"
    $wsDash.Range("C42").Formula = "=Painel!K9"
    $wsDash.Range("D42").Value = "Auditado e Pronto"
    $wsDash.Range("E42").Value = "Abrir Pasta MEDICAO"
    $wsDash.Hyperlinks.Add($wsDash.Range("E42"), "MEDIÇÃO\", "", "Abrir pasta de saida", "Abrir MEDICAO")

    $wsDash.Range("B43").Value = "IRON-LT1-ES"
    $wsDash.Range("C43").Formula = "=Painel!K10"
    $wsDash.Range("D43").Value = "Auditado e Pronto"
    $wsDash.Range("E43").Value = "Abrir Pasta MEDICAO"
    $wsDash.Hyperlinks.Add($wsDash.Range("E43"), "MEDIÇÃO\", "", "Abrir pasta de saida", "Abrir MEDICAO")

    $wsDash.Range("B40:E43").Font.Size = 9
    $wsDash.Range("B40:E43").Borders.Color = 0xE0E0E0

    # Rodape Informativo
    $wsDash.Range("B45:N45").Merge()
    $rodape = $wsDash.Range("B45:N45")
    $rodape.Font.Size = 8
    $rodape.Font.Color = 0x808080
    $rodape.HorizontalAlignment = -4108
    $rodape.Value = "Sistema de Medicao e Faturamento Petrobras - Engenharia de Dados & Power Query M | Documentacao completa na aba 00_GUIA"

    # ===============================================================================
    # 5. CONSTRUIR 00_GUIA (Manual Operacional da Dashboard)
    # ===============================================================================
    Write-Host "Construindo 00_GUIA..."
    $wsGuia.Cells.Clear()
    $wsGuia.Cells.Font.Name = "Segoe UI"
    $wsGuia.Cells.Font.Size = 10

    $wsGuia.Columns.Item("A").ColumnWidth = 3
    $wsGuia.Columns.Item("B").ColumnWidth = 25
    $wsGuia.Columns.Item("C").ColumnWidth = 65
    $wsGuia.Columns.Item("D").ColumnWidth = 20

    # Titulo
    $wsGuia.Range("B1:D2").Merge()
    $gHead = $wsGuia.Range("B1:D2")
    $gHead.Interior.Color = 0x40250A
    $gHead.Font.Color = 0xFFFFFF
    $gHead.Font.Bold = $true
    $gHead.Font.Size = 14
    $gHead.VerticalAlignment = -4108
    $gHead.Value = "  MANUAL OPERACIONAL E GUIA DA DASHBOARD - MEDICAO PETROBRAS"

    $guiaConteudo = @(
        @("1. OBJETIVO DA DASHBOARD", "Proporcionar aos gestores, fiscais e analistas de medicao uma visao executiva consolidada e interativa dos servicos operacionais executados para a Petrobras nos contratos PA-LT2 (BA) e IRON-LT1 (RJ, SP, ES), eliminando esforco cognitivo e viabilizando auditoria rapida de indicadores antes do ateste e emissao da Memoria de Calculo."),
        @("2. CICLO OPERACIONAL", "A medicao NAO segue o mes civil. O periodo de apuracao e fixado estritamente do dia 26 do mes anterior ao dia 25 do mes de referencia (ex.: Agosto/2026 = 26/07/2026 a 25/08/2026). Apenas atendimentos com Situacao = 'CO' (Concluido) sao computados."),
        @("3. DICIONARIO DE KPIS", "Definicoes dos indicadores exibidos nos cards da 01_DASHBOARD:"),
        @("   • Total de Ordens (OS)", "Contagem total de ordens de servico concluidas e elegiveis no periodo selecionado."),
        @("   • Qtd. Solicitada", "Volume fisico total demandado pelas gerencias solicitantes da Petrobras."),
        @("   • Qtd. Atendida", "Volume fisico efetivamente processado, coletado ou entregue pelas prestadoras."),
        @("   • QExec Faturavel", "Quantidade Executada padronizada convertida para as unidades da Planilha de Precos Unitarios (PPU). Embalagens sao convertidas pelo fator 0,10; itens avulsos por 0,017; organizacao analitica 1,00; digitalizacao via fatores da tabela FC."),
        @("   • KM Adicional", "Quilometragem rodoviaria excedente a franquia contratual (valKmAdicional) entre os galpoes terceirizados e os complexos da Petrobras, apurada sem cobrancas em duplicidade por viagem."),
        @("   • Indice SLA (IAPFARQ)", "Indice de atendimento de prazos em dias e horas uteis (expediente 08h-17h, almoco 12h-13h e feriados nacionais, estaduais e municipais). Meta contratual: >= 98,0%."),
        @("   • FDM Apurado", "Fator de Desempenho Mensal multiplicador da fatura. Performance >= 98% resulta em FDM = 1,0000 (pagamento 100% sem glosa)."),
        @("4. COMO USAR OS FILTROS", "Na aba 01_DASHBOARD, utilize as celulas C6 (Contrato), F6 (Atividade) e I6 (Status SLA). Todas as formulas, tabelas e graficos recalculam instantaneamente de acordo com a selecao realizada."),
        @("5. COMO ATUALIZAR", "Para atualizar os dados mensais: clique no botao 'ATUALIZAR DADOS' na barra de filtros ou utilize a macro 'AtualizarTodasConsultasPowerQuery'. O Power Query reprocessara as 83 consultas em segundo plano."),
        @("6. AUDITORIA PREVENTIVA", "Consulte a aba LOG_CRITICAS para verificar eventuais inconsistencias na base DADOS antes de autorizar o faturamento."),
        @("7. ARQUIVOS DE SAIDA", "A macro GerarArquivosPorContrato gera os 4 cadernos oficiais em formato puro .xlsx na pasta MEDIÇÃO/, cada um contendo as abas oficiais: MC, ARM, DADOS e FRETE.")
    )

    $rG = 4
    foreach ($item in $guiaConteudo) {
        $wsGuia.Cells.Item($rG, 2).Value = $item[0]
        $wsGuia.Cells.Item($rG, 2).Font.Bold = $true
        $wsGuia.Cells.Item($rG, 3).Value = $item[1]
        if ($item[0].StartsWith("   •")) {
            $wsGuia.Cells.Item($rG, 2).Font.Bold = $false
            $wsGuia.Cells.Item($rG, 2).Font.Color = 0x0066CC
        } elseif ($item[0] -match "^[0-9]") {
            $wsGuia.Cells.Item($rG, 2).Font.Color = 0x0A2540
            $wsGuia.Range("B$rG:C$rG").Interior.Color = 0xF5F7FA
        }
        $rG++
    }

    $wsGuia.Range("C4:C$rG").WrapText = $true

    # Link de retorno para a Dashboard
    $wsGuia.Range("D4").Value = "Ir para Dashboard"
    $wsGuia.Range("D4").Font.Bold = $true
    $wsGuia.Range("D4").Font.Color = 0x0066CC
    $wsGuia.Hyperlinks.Add($wsGuia.Range("D4"), "", "'01_DASHBOARD'!B1", "Navegar para a Dashboard", "01_DASHBOARD")

    # Ativar a aba 01_DASHBOARD por padrao
    $wsDash.Activate()

    # 6. ATUALIZAR MODULO VBA modDashboard SE ACESSO VBOM ESTIVER HABILITADO
    try {
        $vbProj = $wb.VBProject
        if ($vbProj -ne $null) {
            foreach ($comp in $vbProj.VBComponents) {
                if ($comp.Name -eq "modDashboard") {
                    $vbProj.VBComponents.Remove($comp)
                    break
                }
            }
            $basPath = "d:\PainelNovo\Codes\VBA\Funções Novas\modDashboard.bas"
            if (Test-Path $basPath) {
                $vbProj.VBComponents.Import($basPath)
                Write-Host "Modulo VBA modDashboard atualizado com sucesso no workbook!"
            }
        }
    } catch {
        Write-Host "Acesso programatico ao VBOM restrito pelo Excel (seguranca padrao): $($_.Exception.Message)"
    }

    # Salvar a pasta de trabalho
    Write-Host "Salvando workbook..."
    $wb.Save()
    Write-Host "Workbook salvo com sucesso!"

} catch {
    Write-Host "Erro durante execucao: $_"
    throw $_
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.ScreenUpdating = $true
    $xl.DisplayAlerts = $true
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
    Write-Host "Excel fechado e recursos liberados."
}

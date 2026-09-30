<#
.SYNOPSIS
    limpar_planilha_monitoramento.ps1
    Automação nativa em PowerShell / Excel COM para limpar e redimensionar a aba 'Monitoramento'
    para exatamente 100 linhas, preservando validações de dados e formatações da tabela.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File scripts/limpar_planilha_monitoramento.ps1 -Caminho "Monitoramento/Monitoramento individual_A4UU_v3.0.xlsm"
#>

param (
    [string]$Caminho = ""
)

$ErrorActionPreference = "Stop"

function Limpar-ArquivoExcel {
    param ([string]$Arquivo)

    $ArquivoAbs = (Resolve-Path $Arquivo).Path
    $NomeArquivo = Split-Path -Leaf $ArquivoAbs
    Write-Host "`n[PROCESSANDO] $NomeArquivo..." -ForegroundColor Cyan

    $xl = $null
    $wb = $null

    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $xl.ScreenUpdating = $false
        $xl.EnableEvents = $false

        $wb = $xl.Workbooks.Open($ArquivoAbs)

        # 1. Localizar aba Monitoramento
        $ws = $null
        foreach ($s in $wb.Sheets) {
            if ($s.Name -eq "Monitoramento") {
                $ws = $s
                break
            }
        }

        if ($ws -eq $null) {
            Write-Host "  [AVISO] Aba 'Monitoramento' não encontrada em $NomeArquivo." -ForegroundColor Yellow
            $wb.Close($false)
            return
        }

        # 2. Localizar Tabela Estruturada (ListObject)
        $tbl = $null
        if ($ws.ListObjects.Count -gt 0) {
            try {
                $tbl = $ws.ListObjects.Item("Monitoramento")
            } catch {
                $tbl = $ws.ListObjects.Item(1)
            }
        }

        if ($tbl -eq $null) {
            Write-Host "  [AVISO] Nenhuma tabela estruturada encontrada na aba 'Monitoramento'." -ForegroundColor Yellow
            $wb.Close($false)
            return
        }

        $linhasAtuais = $tbl.ListRows.Count
        Write-Host "  - Tabela: '$($tbl.Name)' | Linhas atuais: $linhasAtuais | Intervalo: $($tbl.Range.Address())"

        # 3. Limpar Filtros Ativos (evita Erro 1004)
        if ($ws.FilterMode) { $ws.ShowAllData() }
        if ($tbl.AutoFilter -ne $null -and $tbl.AutoFilter.FilterMode) { $tbl.AutoFilter.ShowAllData() }

        # 4. Mapear Validações de Dados Atuais
        $validacoes = @{}
        for ($c = 1; $c -le $tbl.ListColumns.Count; $c++) {
            $colRng = $tbl.ListColumns.Item($c).DataBodyRange
            if ($colRng -ne $null) {
                $primeiraCelula = $colRng.Cells.Item(1, 1)
                try {
                    $v = $primeiraCelula.Validation
                    if ($v.Type -ne $null -and $v.Type -ne 0) {
                        $validacoes[$c] = @{
                            Type = $v.Type
                            AlertStyle = $v.AlertStyle
                            Operator = $v.Operator
                            Formula1 = $v.Formula1
                            Formula2 = $v.Formula2
                            IgnoreBlank = $v.IgnoreBlank
                            InCellDropdown = $v.InCellDropdown
                            ShowInput = $v.ShowInput
                            ShowError = $v.ShowError
                        }
                    }
                } catch {}
            }
        }

        # Fallbacks conhecidos da TabelaA
        $fallback = @{
            3 = '=TabelaA!$B$2:$B$5'
            4 = '=TabelaA!$E$2:$E$6'
            5 = '=TabelaA!$AD$2:$AD$57'
            6 = '=TabelaA!$AO$2:$AO$35'
            7 = '=TabelaA!$AQ$2:$AQ$53'
            8 = '=TabelaA!$AH$2:$AH$3'
            25 = '=TabelaA!$AW$2:$AW$3'
        }

        # 5. Redimensionamento Seguro com Resize (1 Header + 100 Dados = 101 Linhas)
        $LINHAS_PADRAO = 100
        $rngNovo = $ws.Range($tbl.Range.Cells.Item(1, 1), $tbl.Range.Cells.Item($LINHAS_PADRAO + 1, $tbl.Range.Columns.Count))
        $tbl.Resize($rngNovo)
        Write-Host "  - Tabela redimensionada com sucesso para 100 linhas úteis ($($tbl.Range.Address()))."

        # 6. Limpeza de Conteúdo e Formatações
        if ($tbl.DataBodyRange -ne $null) {
            $tbl.DataBodyRange.ClearContents()
            $tbl.DataBodyRange.FormatConditions.Delete()
            $tbl.DataBodyRange.ClearFormats()
            $tbl.DataBodyRange.RowHeight = 20
            $tbl.DataBodyRange.VerticalAlignment = -4108 # xlCenter
        }

        # 7. Eliminar Sujeiras Residuais Abaixo da Tabela
        $fimTabela = $tbl.Range.Row + $tbl.Range.Rows.Count - 1
        $lastRow = $ws.Cells.SpecialCells(11).Row # xlCellTypeLastCell
        if ($lastRow -gt $fimTabela) {
            $ws.Rows.Item(($fimTabela + 1).ToString() + ":" + $lastRow.ToString()).Delete() | Out-Null
            Write-Host "  - Linhas residuais da planilha ($($fimTabela + 1) até $lastRow) removidas."
        }

        # 8. Limpar Células à Direita
        $colFimTabela = $tbl.Range.Column + $tbl.Range.Columns.Count - 1
        $lastCol = $ws.Cells.SpecialCells(11).Column
        if ($lastCol -gt $colFimTabela) {
            $ws.Range($ws.Cells.Item(1, $colFimTabela + 1), $ws.Cells.Item($ws.Rows.Count, $lastCol)).Clear() | Out-Null
        }

        # 9. Reaplicar e Garantir Validações de Dados nas 100 Linhas
        for ($c = 1; $c -le $tbl.ListColumns.Count; $c++) {
            $colRng = $tbl.ListColumns.Item($c).DataBodyRange
            if ($colRng -ne $null) {
                $info = $validacoes[$c]
                if ($info -eq $null -and $fallback.ContainsKey($c)) {
                    $info = @{
                        Type = 3
                        AlertStyle = 1
                        Operator = 1
                        Formula1 = $fallback[$c]
                        Formula2 = ""
                        IgnoreBlank = $true
                        InCellDropdown = $true
                        ShowInput = $true
                        ShowError = $true
                    }
                }
                if ($info -ne $null) {
                    try {
                        $colRng.Validation.Delete()
                        if ($info.Formula2 -ne $null -and $info.Formula2 -ne "") {
                            $colRng.Validation.Add($info.Type, $info.AlertStyle, $info.Operator, $info.Formula1, $info.Formula2)
                        } else {
                            $colRng.Validation.Add($info.Type, $info.AlertStyle, $info.Operator, $info.Formula1)
                        }
                        $colRng.Validation.IgnoreBlank = $info.IgnoreBlank
                        $colRng.Validation.InCellDropdown = $info.InCellDropdown
                        $colRng.Validation.ShowInput = $info.ShowInput
                        $colRng.Validation.ShowError = $info.ShowError
                    } catch {}
                }
            }
        }

        # 10. Formatações de Alinhamento e Datas
        try {
            if ($tbl.ListColumns.Count -ge 2) {
                $tbl.ListColumns.Item(2).DataBodyRange.NumberFormat = "dd/mm/aaaa (ddd)"
                $tbl.ListColumns.Item(2).DataBodyRange.HorizontalAlignment = -4108
            }
            foreach ($cAlign in @(3, 4, 5, 8, 15, 20, 25)) {
                if ($tbl.ListColumns.Count -ge $cAlign) {
                    $tbl.ListColumns.Item($cAlign).DataBodyRange.HorizontalAlignment = -4108
                }
            }
            foreach ($cDate in @(12, 17, 18)) {
                if ($tbl.ListColumns.Count -ge $cDate) {
                    $tbl.ListColumns.Item($cDate).DataBodyRange.NumberFormat = "dd/mm/aaaa hh:mm"
                }
            }
        } catch {}

        # Ativar estilos e posicionar em A3
        try {
            $tbl.ShowTableStyleRowStripes = $true
            $tbl.ShowAutoFilter = $true
            $ws.Activate()
            $ws.Range("A3").Select() | Out-Null
        } catch {}

        $wb.Save()
        $wb.Close($false)
        Write-Host "  [OK] Sucesso! Tabela ajustada para 100 linhas (A3:AK102) com validações preservadas." -ForegroundColor Green

    } catch {
        Write-Host "  [ERRO] Falha ao processar $NomeArquivo : $($_.Exception.Message)" -ForegroundColor Red
        if ($wb -ne $null) {
            $wb.Close($false)
        }
    } finally {
        if ($xl -ne $null) {
            $xl.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
        }
    }
}

# Início da Execução
if ([string]::IsNullOrWhiteSpace($Caminho)) {
    $dirMonit = Join-Path (Split-Path -Parent $PSScriptRoot) "Monitoramento"
    $arquivos = Get-ChildItem -Path $dirMonit -Filter "Monitoramento individual_*.xlsm" | Where-Object { $_.Name -notlike "~$*" }
    foreach ($arq in $arquivos) {
        Limpar-ArquivoExcel -Arquivo $arq.FullName
    }
} elseif (Test-Path $Caminho -PathType Leaf) {
    Limpar-ArquivoExcel -Arquivo $Caminho
} elseif (Test-Path $Caminho -PathType Container) {
    $arquivos = Get-ChildItem -Path $Caminho -Filter "Monitoramento individual_*.xlsm" | Where-Object { $_.Name -notlike "~$*" }
    foreach ($arq in $arquivos) {
        Limpar-ArquivoExcel -Arquivo $arq.FullName
    }
} else {
    Write-Host "[ERRO] Caminho inválido: $Caminho" -ForegroundColor Red
}

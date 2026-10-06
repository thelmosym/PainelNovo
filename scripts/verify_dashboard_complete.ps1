$ErrorActionPreference = "Stop"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
try {
    $wb = $xl.Workbooks.Open("d:\PainelNovo\Painel de controle MemoriaPetrobras V6.xlsm", $false, $true)
    $wsDash = $wb.Worksheets.Item("01_DASHBOARD")
    $wsCalc = $wb.Worksheets.Item("04_CALCULOS")
    $wsGuia = $wb.Worksheets.Item("00_GUIA")

    Write-Host "=== TESTE 1: VALORES GERAIS (TODOS) ==="
    Write-Host ("C6 (Contrato):  " + $wsDash.Range("C6").Text)
    Write-Host ("F6 (Atividade): " + $wsDash.Range("F6").Text)
    Write-Host ("I6 (SLA):       " + $wsDash.Range("I6").Text)
    Write-Host ("Card OS:        " + $wsDash.Range("B10").Text + " (Esperado: 582)")
    Write-Host ("Card Sol:       " + $wsDash.Range("D10").Text + " (Esperado: 93.083)")
    Write-Host ("Card Atend:     " + $wsDash.Range("F10").Text + " (Esperado: 93.082)")
    Write-Host ("Card QExec:     " + $wsDash.Range("H10").Text + " (Esperado: 91.026,65)")
    Write-Host ("Card KM:        " + $wsDash.Range("J10").Text + " (Esperado: 2.852)")
    Write-Host ("Card SLA:       " + $wsDash.Range("L10").Text + " (Esperado: ~51,8%)")

    Write-Host "`n=== TESTE 2: TABELA POR CONTRATO ==="
    for ($r = 16; $r -le 20; $r++) {
        $cName = $wsDash.Range("B$r").Text
        $cOS   = $wsDash.Range("C$r").Text
        $cQEx  = $wsDash.Range("D$r").Text
        $cSol  = $wsDash.Range("E$r").Text
        $cAtn  = $wsDash.Range("F$r").Text
        $cKM   = $wsDash.Range("G$r").Text
        $cSLA  = $wsDash.Range("J$r").Text
        Write-Host ("Linha $r : $cName | OS=$cOS | QExec=$cQEx | Sol=$cSol | Atend=$cAtn | KM=$cKM | SLA=$cSLA")
    }

    Write-Host "`n=== TESTE 3: GRAFICOS NA DASHBOARD ==="
    Write-Host ("Total Shapes: " + $wsDash.Shapes.Count)
    for ($s = 1; $s -le $wsDash.Shapes.Count; $s++) {
        $shp = $wsDash.Shapes.Item($s)
        Write-Host ("Shape $s : " + $shp.Name + " | Tipo=" + $shp.Type)
    }

    Write-Host "`n=== TESTE 4: TESTE DE REATIVIDADE DO FILTRO DE CONTRATO ==="
    $wsDash.Range("C6").Value = "PA-LT2-BA"
    $wb.Application.Calculate()
    Write-Host ("Apos filtro PA-LT2-BA -> Card OS: " + $wsDash.Range("B10").Text + " (Esperado: 162) | QExec: " + $wsDash.Range("H10").Text + " (Esperado: 16.781,70)")

    $wsDash.Range("C6").Value = "IRON-LT1-RJ"
    $wb.Application.Calculate()
    Write-Host ("Apos filtro IRON-LT1-RJ -> Card OS: " + $wsDash.Range("B10").Text + " (Esperado: 377) | QExec: " + $wsDash.Range("H10").Text + " (Esperado: 72.880,75)")

    $wsDash.Range("C6").Value = "Todos os Contratos"
    $wb.Application.Calculate()
    Write-Host ("Apos reset -> Card OS: " + $wsDash.Range("B10").Text + " (Esperado: 582)")

    Write-Host "`n=== TESTE 5: VERIFICACAO DE ERROS (#VALOR, #REF, #N/D) ==="
    $erros = 0
    foreach ($ws in @($wsDash, $wsCalc, $wsGuia)) {
        $usedRng = $ws.UsedRange
        for ($row = 1; $row -le [Math]::Min(50, $usedRng.Rows.Count); $row++) {
            for ($col = 1; $col -le [Math]::Min(20, $usedRng.Columns.Count); $col++) {
                $txt = $ws.Cells.Item($row, $col).Text
                if ($txt -match "#VALOR!|#REF!|#N/D|#NOME\?|#DIV/0!") {
                    Write-Host ("ALERTA DE ERRO em " + $ws.Name + "!" + $ws.Cells.Item($row, $col).Address + ": " + $txt)
                    $erros++
                }
            }
        }
    }
    if ($erros -eq 0) {
        Write-Host "ZERO erros de formula encontrados! Dashboard 100% integra!"
    } else {
        Write-Host ("Total de celulas com erro: " + $erros)
    }

} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}

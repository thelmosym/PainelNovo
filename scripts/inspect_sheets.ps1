$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
try {
    $wb = $xl.Workbooks.Open("d:\PainelNovo\Painel de controle MemoriaPetrobras V6.xlsm", $false, $true)
    Write-Host ("Total sheets: " + $wb.Worksheets.Count)
    for ($i = 1; $i -le [Math]::Min(12, $wb.Worksheets.Count); $i++) {
        Write-Host ("Sheet $i : " + $wb.Worksheets.Item($i).Name)
    }

    # Inspect 04_CALCULOS if it exists
    foreach ($ws in $wb.Worksheets) {
        if ($ws.Name -eq "04_CALCULOS") {
            Write-Host "--- 04_CALCULOS ---"
            Write-Host ("A1: " + $ws.Range("A1").Text + " | B1: " + $ws.Range("B1").Text)
            Write-Host ("A2: " + $ws.Range("A2").Text + " | B2: " + $ws.Range("B2").Text)
            Write-Host ("A3: " + $ws.Range("A3").Text + " | B3: " + $ws.Range("B3").Text)
            Write-Host ("B6 (OS): " + $ws.Range("B6").Text + " | C6 (Sol): " + $ws.Range("C6").Text + " | D6 (Atend): " + $ws.Range("D6").Text)
            Write-Host ("E6 (QExec): " + $ws.Range("E6").Text + " | F6 (KM): " + $ws.Range("F6").Text + " | I6 (SLA): " + $ws.Range("I6").Text)
            Write-Host "Contratos:"
            for ($r = 10; $r -le 14; $r++) {
                Write-Host ("Row $r : " + $ws.Cells.Item($r, 1).Text + " | OS=" + $ws.Cells.Item($r, 2).Text + " | QExec=" + $ws.Cells.Item($r, 3).Text)
            }
        }
        if ($ws.Name -eq "01_DASHBOARD") {
            Write-Host "--- 01_DASHBOARD ---"
            Write-Host ("C6 (Contrato): " + $ws.Range("C6").Text)
            Write-Host ("F6 (Atividade): " + $ws.Range("F6").Text)
            Write-Host ("I6 (SLA): " + $ws.Range("I6").Text)
            Write-Host ("Card OS (B10): " + $ws.Range("B10").Text)
            Write-Host ("Card Sol (D10): " + $ws.Range("D10").Text)
            Write-Host ("Card Atend (F10): " + $ws.Range("F10").Text)
            Write-Host ("Card QExec (H10): " + $ws.Range("H10").Text)
            Write-Host ("Card KM (J10): " + $ws.Range("J10").Text)
            Write-Host ("Card SLA (L10): " + $ws.Range("L10").Text)
            Write-Host "Shapes count: " $ws.Shapes.Count
            for ($s = 1; $s -le $ws.Shapes.Count; $s++) {
                Write-Host ("Shape $s : " + $ws.Shapes.Item($s).Name + " Type=" + $ws.Shapes.Item($s).Type)
            }
        }
    }
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}

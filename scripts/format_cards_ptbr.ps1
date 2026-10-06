$ErrorActionPreference = "Stop"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
try {
    $wb = $xl.Workbooks.Open("d:\PainelNovo\Painel de controle MemoriaPetrobras V6.xlsm", $false, $false)
    $wsDash = $wb.Worksheets.Item("01_DASHBOARD")

    Write-Host "Antes:"
    Write-Host ("B10 Text: " + $wsDash.Range("B10").Text + " | FormatLocal: " + $wsDash.Range("B10").NumberFormatLocal)
    Write-Host ("H10 Text: " + $wsDash.Range("H10").Text + " | FormatLocal: " + $wsDash.Range("H10").NumberFormatLocal)

    # Definir formato local exato para PT-BR
    $wsDash.Range("B10:C11").NumberFormatLocal = "#.##0"
    $wsDash.Range("D10:E11").NumberFormatLocal = "#.##0"
    $wsDash.Range("F10:G11").NumberFormatLocal = "#.##0"
    $wsDash.Range("H10:I11").NumberFormatLocal = "#.##0,00"
    $wsDash.Range("J10:K11").NumberFormatLocal = "#.##0"
    $wsDash.Range("L10:N11").NumberFormatLocal = "0,0%"

    # Tabela linhas 16 a 20
    $wsDash.Range("C16:C20").NumberFormatLocal = "#.##0"
    $wsDash.Range("D16:D20").NumberFormatLocal = "#.##0,00"
    $wsDash.Range("E16:I20").NumberFormatLocal = "#.##0"
    $wsDash.Range("J16:J20").NumberFormatLocal = "0,0%"

    $wb.Save()
    Write-Host "`nDepois:"
    Write-Host ("B10 (OS):    " + $wsDash.Range("B10").Text)
    Write-Host ("D10 (Sol):   " + $wsDash.Range("D10").Text)
    Write-Host ("F10 (Atend): " + $wsDash.Range("F10").Text)
    Write-Host ("H10 (QExec): " + $wsDash.Range("H10").Text)
    Write-Host ("J10 (KM):    " + $wsDash.Range("J10").Text)
    Write-Host ("L10 (SLA):   " + $wsDash.Range("L10").Text)
    Write-Host "`nTabela Linha 20 (Total):"
    Write-Host ("OS:    " + $wsDash.Range("C20").Text)
    Write-Host ("QExec: " + $wsDash.Range("D20").Text)
    Write-Host ("Sol:   " + $wsDash.Range("E20").Text)
    Write-Host ("Atend: " + $wsDash.Range("F20").Text)
    Write-Host ("KM:    " + $wsDash.Range("G20").Text)
    Write-Host ("SLA:   " + $wsDash.Range("J20").Text)
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}

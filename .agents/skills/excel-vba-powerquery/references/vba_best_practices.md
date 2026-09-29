# Guia de Boas Práticas em VBA (Visual Basic for Applications)

Este documento complementa a skill `excel-vba-powerquery` com padrões e técnicas consolidadas para macros robustas e de alta performance.

---

## 1. Otimização de Performance: Matrizes em Memória (Arrays)

A leitura e escrita célula por célula na planilha é o principal gargalo em VBA. Sempre prefira carregar o intervalo em uma matriz na memória:

```vba
Sub ProcessarIntervaloEmMemoria(ws As Worksheet)
    Dim rng As Range
    Dim dados As Variant
    Dim r As Long, c As Long
    
    Set rng = ws.Range("A2:D" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
    If rng.Row < 2 Or rng.Cells.Count = 0 Then Exit Sub
    
    ' Carrega todos os valores de uma vez em uma matriz 2D (1 To n, 1 To m)
    dados = rng.Value
    
    For r = 1 To UBound(dados, 1)
        ' Exemplo: transformar texto em maiúsculas na coluna 2
        If Not IsEmpty(dados(r, 2)) Then
            dados(r, 2) = UCase(Trim(dados(r, 2)))
        End If
    Next r
    
    ' Grava os dados processados de volta na planilha de uma só vez
    rng.Value = dados
End Sub
```

---

## 2. Manipulação de Tabelas Estruturadas (`ListObject`)

Tabelas nativas do Excel (`ListObject`) facilitam referências sem depender de números fixos de linhas:

```vba
Sub ManipularTabela(ws As Worksheet, nomeTabela As String)
    Dim tbl As ListObject
    Dim novaLinha As ListRow
    
    On Error Resume Next
    Set tbl = ws.ListObjects(nomeTabela)
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "Tabela " & nomeTabela & " não encontrada.", vbExclamation
        Exit Sub
    End If
    
    ' Inserir nova linha
    Set novaLinha = tbl.ListRows.Add
    novaLinha.Range(1, tbl.ListColumns("Data").Index).Value = Date
    novaLinha.Range(1, tbl.ListColumns("Status").Index).Value = "Concluído"
    
    ' Limpar dados mantendo cabeçalho
    ' If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete
End Sub
```

---

## 3. Atualização Síncrona de Conexões do Power Query

Por padrão, conexões OLEDB do Power Query podem atualizar em segundo plano, fazendo com que o código VBA seguinte execute antes do término da atualização dos dados. Para forçar a execução síncrona:

```vba
Sub AtualizarConsultasSincrono()
    Dim conn As WorkbookConnection
    Dim oledbConn As OLEDBConnection
    
    For Each conn In ThisWorkbook.Connections
        If conn.Type = xlConnectionTypeOLEDB Then
            Set oledbConn = conn.OLEDBConnection
            oledbConn.BackgroundQuery = False
        ElseIf conn.Type = xlConnectionTypeMODEL Then
            ' Conexões de Modelo de Dados PowerPivot
        End If
    Next conn
    
    ThisWorkbook.RefreshAll
    DoEvents
End Sub
```

---

## 4. Exportação Segura de Abas para Novos Arquivos

Ao exportar abas para arquivos de entrega (como memórias de cálculo ou relatórios consolidados):

```vba
Sub ExportarAbaParaArquivo(nomeAba As String, caminhoDestino As String)
    Dim wsOrigem As Worksheet
    Dim wbNovo As Workbook
    
    Set wsOrigem = ThisWorkbook.Sheets(nomeAba)
    
    ' Copia a aba para um novo workbook isolado
    wsOrigem.Copy
    Set wbNovo = ActiveWorkbook
    
    ' Salva sem macros (.xlsx)
    Application.DisplayAlerts = False
    wbNovo.SaveAs Filename:=caminhoDestino, _
                  FileFormat:=xlOpenXMLWorkbook, _
                  CreateBackup:=False
    wbNovo.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Sub
```

---

## 5. Dicionários e Coleções para Desduplicação e Busca Rápida

Utilize `Scripting.Dictionary` para buscas em tempo $O(1)$:

```vba
Sub ExemploDicionario()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare ' Ignora maiúsculas/minúsculas
    
    Dim chave As String
    chave = "CONTRATO-001"
    
    If Not dict.Exists(chave) Then
        dict.Add chave, 10500.50
    Else
        dict(chave) = dict(chave) + 10500.50
    End If
End Sub
```

---

## 6. Codificação de Arquivos e Compatibilidade com o Editor do VBA (Português/VBE)

O editor do VBA (*Visual Basic Editor - VBE*) do Microsoft Excel é uma aplicação legada baseada em Windows ANSI. Ele **NÃO oferece suporte nativo a UTF-8**.

### Consequências de Salvar em UTF-8:
* Textos e comentários em português com acentuação (`á`, `é`, `í`, `ó`, `ú`, `ã`, `õ`, `ç`) transformam-se em sequências corrompidas de bytes (*mojibake*, como `Ã¡`, `Ã§Ã£o`).
* Emojis de 4 bytes (como `🟡`, `🟢`, `✅`, `❌`) transformam-se em caracteres ilegíveis (`ðŸŸ¡`) ou provocam erros de sintaxe em tempo de compilação.
* Ao importar arquivos `.bas` exportados sem o cabeçalho `Attribute VB_Name`, o Excel renomeia o módulo arbitrariamente para `Módulo1`.

### Padrão Técnico Obrigatório para Códigos VBA:
1. **Codificação Windows-1252 (CP1252 / ANSI):**
   Todos os arquivos `.bas`, `.cls`, `.frm` e `.vba` do repositório devem ser gerados e mantidos estritamente na codificação **Windows-1252**.
2. **Terminações de Linha CRLF (`\r\n`):**
   Sempre utilizar quebras de linha padrão Windows (`CRLF`). Quebras Unix puras (`LF`) podem causar exibição em linha única no VBE.
3. **Substituição de Emojis por Tags Textuais em Português:**
   * `🟡` $\rightarrow$ `[AMARELO]`
   * `🟢` $\rightarrow$ `[VERDE]`
   * `🔴` $\rightarrow$ `[VERMELHO]`
   * `✅` $\rightarrow$ `[OK]`
   * `⚠️` $\rightarrow$ `[AVISO]`
   * `❌` $\rightarrow$ `[ERRO]`
4. **Cabeçalho de Módulo Mandatório:**
   A primeira linha de qualquer módulo `.bas` deve conter:
   ```vba
   Attribute VB_Name = "NomeDoModulo"
   ```
5. **Automação com o Utilitário `fix_vba_encoding.py`:**
   Execute periodicamente ou após modificações:
   ```bash
   python .agents/skills/excel-vba-powerquery/scripts/fix_vba_encoding.py "Codes/VBA"
   ```


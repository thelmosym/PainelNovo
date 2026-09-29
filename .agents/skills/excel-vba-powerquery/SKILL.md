---
name: excel-vba-powerquery
description: >-
  Especialista em automação de Excel com foco em VBA (Macros), Power Query (Linguagem M)
  e engenharia de planilhas (.xlsm / .xlsx). Use sempre que o usuário solicitar desenvolvimento,
  correção, refatoração ou otimização de rotinas VBA, criação/ajuste de consultas e funções Power Query M,
  leitura e inspeção estrutural de pastas de trabalho e orquestração de fluxos de dados no Excel.
  Inclui a garantia estrita de codificação Windows-1252 (ANSI) e quebras CRLF em códigos VBA para que textos e acentos em português apareçam corretamente no editor do VBA (VBE).
---

# Skill: Automação Avançada de Excel (VBA & Power Query M)

Esta skill orienta o desenvolvimento, depuração e manutenção de soluções corporativas no Microsoft Excel, com foco na divisão correta de responsabilidades entre **Power Query (M)** e **VBA (Macros)**.

---

## 1. Princípios Arquiteturais e Divisão de Papéis

1. **Power Query (Linguagem M) — ETL e Modelagem:**
   * Utilizar para extração de múltiplas fontes (pastas, arquivos `.xlsx`, `.csv`, bancos de dados).
   * Transformação, limpeza, tipagem de colunas e mesclagem de tabelas relacionais.
   * Não delegar ao VBA rotinas de loop linha a linha para tratar dados se o Power Query puder fazer isso nativamente com melhor desempenho.

2. **VBA (Macros) — Orquestração e Ações de Sistema:**
   * Controle de interface e interação com o usuário (UserForms, botões, caixas de diálogo).
   * Orquestração de atualizações (ex.: disparar atualização de conexões e consultas do Power Query em sequência síncrona).
   * Operações de sistema de arquivos (criar pastas, salvar cópias `.xlsx`/`.pdf`, exportar abas, enviar e-mails).
   * Regras operacionais imperativas que dependem de eventos de planilha (`Worksheet_Change`, `Workbook_Open`).

---

## 2. Boas Práticas e Padrões para VBA

Consulte o guia detalhado em [vba_best_practices.md](./references/vba_best_practices.md).

### Regras Mandatórias:
* **Codificação Estrita em Windows-1252 (CP1252 / ANSI) & Quebras CRLF:**
  O editor do VBA (*Visual Basic Editor - VBE*) é uma aplicação ANSI legada e **NÃO** oferece suporte nativo a UTF-8. Arquivos `.bas`, `.cls` e `.frm` salvos em UTF-8 corrompem caracteres acentuados em português (`ç`, `ã`, `é`, `ó`, etc.), gerando *mojibake* ao serem importados no Excel. Todos os códigos VBA devem ser estritamente salvos na codificação **Windows-1252 (CP1252)** com terminações de linha **CRLF (`\r\n`)**.
* **Proibição de Emojis Unicode de 4 bytes:**
  Emojis como `🟡`, `🟢`, `✅`, `❌` não existem na tabela de caracteres CP1252 e causam corrupção (`ðŸŸ¡`) ou erro no VBE. Devem ser sempre substituídos por tags textuais claras em português: `[AMARELO]`, `[VERDE]`, `[OK]`, `[AVISO]`, `[ERRO]`.
* **Identificador de Módulo Obrigatório (`Attribute VB_Name`):**
  A linha 1 de todo arquivo `.bas` exportado deve conter `Attribute VB_Name = "<NomeDoModulo>"` para que o Excel preserve o nome do módulo ao ser importado, em vez de atribuir genericamente `Módulo1`.
* **Sempre declarar variáveis explicitamente:** Todo módulo deve iniciar com `Option Explicit`.
* **Nunca usar `.Select` ou `.Activate`:** Trabalhe diretamente com objetos qualificados (`ws.Range("A1")`, `tbl.DataBodyRange`).
* **Padrão de Alta Performance para Rotinas:**
  Sempre desabilitar atualização visual, alertas e cálculo automático no início e garantir a restauração no bloco de tratamento de erro:
  ```vba
  Sub ExecutarRotinaOtimizada()
      On Error GoTo TratarErro
      
      Dim appCalc As XlCalculation
      appCalc = Application.Calculation
      
      Application.ScreenUpdating = False
      Application.DisplayAlerts = False
      Application.EnableEvents = False
      Application.Calculation = xlCalculationManual
      
      ' --- Lógica aqui ---
      
  SairRotina:
      Application.Calculation = appCalc
      Application.EnableEvents = True
      Application.DisplayAlerts = True
      Application.ScreenUpdating = True
      Exit Sub
      
  TratarErro:
      MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Erro na Execução"
      Resume SairRotina
  End Sub
  ```
* **Compatibilidade 64-bit:** Funções de API do Windows devem usar `PtrSafe` e `LongPtr` onde aplicável.
* **Processamento de Volumes em Memória:** Ao manipular dados de células via VBA, carregue o intervalo em uma matriz (`Variant Array`), processe na memória e devolva ao range de uma só vez.
* **Atualização Síncrona do Power Query via VBA:** Desabilitar `BackgroundQuery` para evitar concorrência antes de processar etapas dependentes:
  ```vba
  Dim conn As WorkbookConnection
  For Each conn In ThisWorkbook.Connections
      If conn.Type = xlConnectionTypeOLEDB Or conn.Type = xlConnectionTypeMODEL Then
          conn.OLEDBConnection.BackgroundQuery = False
      End If
  Next conn
  ThisWorkbook.RefreshAll
  ```

---

## 3. Boas Práticas e Padrões para Power Query (Linguagem M)

Consulte o guia de funções e sintaxe em [power_query_m_guide.md](./references/power_query_m_guide.md).

### Regras Mandatórias:
* **Estrutura Padrão `let ... in`:** Nomes de etapas claros e descritivos em português ou inglês (ex.: `#"Linhas Filtradas"`, `#"Tipo Alterado"`).
* **Tipagem Explícita e Precoce:** Sempre garanta que `Table.TransformColumnTypes` defina tipos primitivos corretos (`type text`, `type number`, `type date`, etc.) logo após a carga da fonte.
* **Tratamento de Nulos e Inconsistências:** Usar construções defensivas com `try ... otherwise ...` ou `Table.ReplaceErrorValues`.
* **Funções Reutilizáveis (`fn...`):** Funções customizadas devem ter parâmetros explicitamente tipados e documentação de entrada e saída:
  ```powerquery
  (dataInicial as date, prazoDias as number) as date =>
  let
      // Cálculo de data de SLA
      DataFinal = Date.AddDays(dataInicial, prazoDias)
  in
      DataFinal
  ```
* **Performance e Query Folding:**
  * Filtre e remova colunas o mais cedo possível no fluxo de etapas.
  * Tenha cautela com `Table.Buffer`: use apenas quando uma tabela intermediária menor for consultada repetidamente em merges locais.

---

## 4. Estrutura do Repositório e Versionamento

Ao editar ou propor alterações no repositório:
* **Códigos M:** Manter arquivos `.m` sincronizados na pasta [Codes/CodigosM/](../../Codes/CodigosM/).
* **Códigos VBA:** Manter módulos exportados (`.bas`, `.vba`, `.cls`) sincronizados na pasta [Codes/VBA/](../../Codes/VBA/), obrigatoriamente codificados em **Windows-1252 (ANSI)** com quebras **CRLF**.
* **Correção e Sanitização de Encoding:** Utilize o utilitário [scripts/fix_vba_encoding.py](./scripts/fix_vba_encoding.py) para auditar e converter automaticamente arquivos VBA para CP1252/CRLF, eliminando mojibake e protegendo acentuações em português.
* **Inspeção de Planilhas:** Utilize o script utilitário [scripts/inspect_workbook.py](./scripts/inspect_workbook.py) para auditar abas, tabelas e objetos de pastas de trabalho `.xlsm`/`.xlsx` sem precisar de dependências externas.

# Apontamentos e Proposta de Melhorias
## Projeto: Painel de Controle MemoriaPetrobras V6.xlsm

**Data do relatório:** 23/09/2026  
**Ambiente:** Excel com suporte a Macros (`.xlsm`), Power Query (Linguagem M) e VBA  
**Contratos monitorados:** `IRON-LT1` (Lote 1 - RJ, SP, ES) e `PA-LT2` (Lote 2 - BA)  
**Objetivo do projeto:** Consolidação de atendimentos operacionais, apuração de SLA/prazos, cálculo de FDM (Fator de Desempenho Mensal), cálculo de volumes executados (QExec), armazenagem e frete, com geração automatizada de arquivos de Memória de Cálculo (`.xlsx`) na pasta `MEDIÇÃO/`.

---

## 1. Visão Geral da Arquitetura e Estado Atual

O projeto opera atualmente com uma **arquitetura híbrida** resultante de uma migração gradual de rotinas manuais/célula a célula para consultas estruturadas:

1. **Motor Legado em VBA (`Codes/VBA/FuncoesAntigas/`):**
   - Contém rotinas históricas como `Planilha7.cls`, `01 funcoes principais.vba`, `A_Calc_Medicao.bas` e `fncGeral.bas`.
   - Realizava leitura e validação linha por linha em planilhas do Excel, com checagem de feriados e cruzamentos manuais.
2. **Motor Moderno em Power Query M (`Codes/CodigosM/`):**
   - Composto por 83 arquivos de consultas e funções M.
   - Lê os arquivos de monitoramento individual da pasta `Monitoramento/`, aplica filtros (`Situação = "CO"`, período de faturamento), enriquece dados via merges (`TabelaB-V2.xlsx`, tabelas de PPU, frete, fatores de conversão) e carrega os resultados em tabelas estruturadas do Excel.
3. **Orquestrador Atual em VBA (`Codes/VBA/Funções Novas/`):**
   - `Modulo2.vba` (`modAtualizarConsultas`): macro `AtualizarTodasConsultasPowerQuery` para disparar a atualização de conexões e exibir status na célula `Painel!B12`.
   - `modulo1.vba` (`modGerarArquivosPorContrato`): macro `GerarArquivosPorContrato` para gerar os 4 arquivos finais (`PA-LT2`, `IRON-LT1-RJ`, `IRON-LT1-SP`, `IRON-LT1-ES`) contendo as abas `MC`, `ARM`, `DADOS` e `FRETE`.
   - `modDashboard.bas`: módulo experimental para desenhar um painel operacional com cards de indicadores e resumos.

---

## 2. Diagnóstico Detalhado do Código VBA

### 2.1. Módulo de Atualização das Consultas (`Modulo2.vba`)

| Ponto de Atenção | Causa Raiz Identificada | Impacto no Sistema | Recomendação de Melhoria |
| :--- | :--- | :--- | :--- |
| **Atualização indiscriminada de 91 conexões** | O loop `For Each cn In wb.Connections` atualiza todas as conexões sequencialmente. | Reavaliação desnecessária de consultas intermediárias, parâmetros individuais (`localBA`, `localES`, etc.) e funções M, causando lentidão severa e múltiplos acessos a disco. | Atualizar exclusivamente as tabelas de destino (ou usar `ThisWorkbook.RefreshAll` síncrono controlado) respeitando o grafo de dependências do Power Query. |
| **Conexões duplicadas e órfãs** | Existem conexões ativas como `FRETE_IRON-LT1-ES1`, `FRETE_IRON-LT1-RJ(1)`, `TabelaFDMPorContrato1` e conexões sem consulta vinculada. | Mensagens de alerta frequentes (*"Query does not exist"*), atrasos e inconsistências na carga de dados. | Remover conexões duplicadas e limpar as conexões órfãs da pasta de trabalho. |
| **Falso positivo de status verde (`COR_VERDE`)** | No encerramento da rotina (linhas 124-133), a célula `B12` é pintada de verde mesmo quando `iComErro > 0`. | O usuário visualiza status verde e assume que toda a medição foi calculada com sucesso, quando na verdade conexões vitais podem ter falhado. | Implementar regra de cores estrita: **Verde** (100% OK), **Âmbar/Laranja** (sucesso parcial com falhas registradas) e **Vermelho** (erro impeditivo/fatal). |
| **Gestão insegura do estado do Excel** | `Application.EnableEvents = False` é forçado sem gravar o estado anterior da aplicação. | Se o processo for interrompido pelo usuário ou por erro não tratado, o Excel permanece com eventos desabilitados na sessão. | Salvar o estado em variável local (`bPrevEvents = Application.EnableEvents`) e restaurar no bloco `Finally`/`Exit Sub`. |
| **Forçamento de OLEDB** | `cn.OLEDBConnection.BackgroundQuery = False` | Gera erro em conexões que não são do tipo OLEDB (ignorado apenas por `On Error Resume Next`). | Verificar se `cn.Type = xlConnectionTypeOLEDB` antes de acessar a propriedade `OLEDBConnection`. |

---

### 2.2. Módulo de Geração de Arquivos por Contrato (`modulo1.vba`)

| Ponto de Atenção | Causa Raiz Identificada | Impacto no Sistema | Recomendação de Melhoria |
| :--- | :--- | :--- | :--- |
| **Risco de corte de dados em `ObterAreaRealDados`** | Usa `ws.Cells(ws.Rows.Count, 1).End(xlUp).Row` e `ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column`. | Se a coluna A tiver células vazias no final, ou se os cabeçalhos não começarem estritamente em A1, linhas e colunas reais de medição são truncadas no arquivo exportado. | Como as abas são geradas por Power Query, obter o intervalo seguro nativo via `ws.ListObjects(1).Range` (ou `ws.UsedRange`). |
| **Gargalo de performance em `ReplicarAlturasLinhas`** | O código itera linha por linha (`For iLin = 1 To rngOrigem.Rows.Count: wsDestino.Rows(iLin).RowHeight = ...: Next iLin`). | Em planilhas com milhares de registros de medição, o Excel trava por minutos ajustando alturas célula a célula. | Substituir a cópia manual de range pela cópia nativa da aba (`wsOrigem.Copy After:=...`) seguida de desconexão de fórmulas (`wsDestino.UsedRange.Value = wsDestino.UsedRange.Value`). Processamento instantâneo. |
| **Configuração de contratos "Hardcoded"** | Funções `ListaContratos`, `ListaNumerosContrato` e `ListaSufixosEstado` contêm arrays estáticos no código (`PA-LT2`, `4600687006`, `BA`, etc.). | Qualquer alteração contratual, novo lote ou renovação de número SAP exige intervenção manual no código VBA. | Mover o mapeamento de contratos para uma tabela estruturada na aba `Validações` e carregar dinamicamente na macro. |
| **Hyperlinks fixos nas células K7:K10** | A escrita de links dos arquivos gerados aponta rigidamente para a coluna K a partir da linha 7. | Se a ordem dos contratos mudar ou forem incluídos novos itens, os links são sobrescritos incorretamente. | Escrever os links com base na correspondência do nome do contrato na tabela de controle. |
| **Sensibilidade Petrobras (MIP/AIP)** | Chamada `wb.SensitivityLabel.SetLabel lbl, lbl` com repetição de parâmetros. | Pode falhar dependendo da versão do Office 365 e do status de vinculação ao tenant corporativo. | Validar disponibilidade da API de rotulagem com tratamento de exceção seguro sem interromper o fluxo principal. |

---

### 2.3. Módulo do Dashboard Experimental (`modDashboard.bas`)

* **Bug de Sobreposição de Linha:**
  A rotina `ConstruirFiltros` escreve os títulos de filtros nas células `B9:L9` ("Periodo", "Contrato", "Situacao"...). Logo em seguida, `ConstruirCards` sobrescreve a mesma linha `B9:L9` com os nomes dos cards de KPI ("REGISTROS", "QTD. SOLICITADA"...), apagando a seção anterior.
* **Cálculo Imperativo em VBA:**
  As contagens e somas dos KPIs são executadas via iteração sequencial em VBA sobre a planilha `DADOS`. Isso impede que o painel seja dinâmico ao aplicar filtros manuais ou segmentações de dados. O ideal é estruturar os cards com fórmulas nativas do Excel (`SOMASE`, `CONT.SES`, `SUBTOTAL`) ou Tabela Dinâmica conectada ao modelo de dados.

---

## 3. Divergência Crítica de Negócio: SLA e Feriados (VBA vs Power Query)

> [!WARNING]
> ### Risco de Glosa e Penalidades por Divergência de Regra de Prazo
> 
> * **No VBA Legado (`01 funcoes principais.vba`):**
>   A rotina `calcular_dataSLA` pesquisava a tabela de feriados por município e estado (`pesquisar_EFeriado(dDataInicial, sMunicipio, sUF)`). Dias com feriados nacionais, estaduais ou municipais eram suspensos da contagem do prazo de atendimento.
> * **No Power Query M (`fnCalcularDataSLA.m`):**
>   A função valida o expediente exclusivamente por:
>   ```powerquery
>   TemExpediente = (data as date) as logical =>
>       Date.DayOfWeek(data, Day.Monday) <= 4
>   ```
>   A regra considera apenas **segunda a sexta-feira**, ignorando completamente feriados nacionais e locais.
> * **Efeito Operacional:** Solicitações abertas na véspera de feriados (Carnaval, Corpus Christi, 7 de Setembro, feriados municipais de bases como Macaé, Santos, Salvador, etc.) são classificadas no Power Query como **"FP" (Fora do Prazo)** de forma indevida, gerando apurações incorretas de SLA e cálculo incorreto de FDM.
> * **Regra de 0,5 Dia (Meio Período):** A função M `fnCalcularPrazoSLA.m` utiliza faixas fixas de horário (8h-11h e 11h-16h) que devem ser auditadas com a fiscalização do contrato para confirmar se refletem exatamente os termos aditivos vigentes.

---

## 4. Diagnóstico das Fórmulas e Configurações das Planilhas

Na aba [`Validações`](file:///d:/PainelNovo/Painel%20de%20controle%20MemoriaPetrobras%20V6.xlsm):

1. **Célula D6 com erro `#NAME?`:**
   - Fórmula gravada: `=LEFT(CÉLULA("filename",A4),MAX(IFERROR(SEARCH("\",CÉLULA("filename",A4),...))))`
   - **Causa:** O identificador foi gravado em português (`CÉLULA`) dentro do XML da fórmula, onde a especificação do Excel requer a função canônica em inglês (`CELL`).
2. **Célula D5 sensível à localização:**
   - Fórmula: `=LEFT(CELL("nome.arquivo"),FIND("[",CELL("nome.arquivo"))-1)`
   - O parâmetro `"nome.arquivo"` só funciona em versões de idioma português brasileiro. Se executado em estações com Excel em inglês, retorna erro. O parâmetro universal é `"filename"`.
3. **Caminhos Locais Fixos na Célula D10:**
   - Registra `C:\Thelmo\Tabelas Especificas`. Se outro técnico executar o painel em sua estação, qualquer rotina que dependa dessa célula falhará por caminho inexistente.

---

## 5. Diagnóstico do Pipeline Power Query (Linguagem M)

1. **Consumo Excessivo de Memória por `Table.Buffer`:**
   - Na consulta principal `Painel T2M.m`, o comando `Table.Buffer` é aplicado em cascata em várias etapas intermediárias (`ExpandirTabelaBuffer`, `FiltrarContratosBuffer`, `FiltrarPeriodoValidoBuffer`, `OrdenarContratoBuffer`).
   - O uso de buffer deve ser restrito a **tabelas de dimensão pequenas** (ex.: lista de localidades ou fatores de conversão). Aplicar buffers sucessivos na tabela fato principal força a materialização de centenas de milhares de células em memória RAM, desativando o streaming do mecanismo Mashup e sobrecarregando o Excel.
2. **Tratamento Silencioso de Erros (`try ... otherwise`):**
   - Na etapa de classificação de FDM e cálculo de prazos, erros de transformação são mascarados com retornos `null` ou `""`.
   - **Risco:** Registros com novos tipos de atividade, divergências de grafia em municípios ou falhas de preenchimento nos arquivos de monitoramento deixam de ser processados sem que a equipe receba nenhum alerta de exceção.

---

## 6. Plano de Ação e Recomendações Estruturadas

### Fase 1: Correções Imediatas (Estabilidade e Segurança Operacional)

- [ ] **Saneamento de Conexões:** Excluir da pasta de trabalho as conexões duplicadas (`FRETE_IRON-LT1-ES1`, `TabelaFDMPorContrato1`, etc.) e as conexões órfãs.
- [ ] **Refatoração do `Modulo2.vba`:**
  - Implementar semáforo visual correto em `Painel!B12` (Verde apenas para 100% OK, Âmbar para falha parcial, Vermelho para erro geral).
  - Preservar o estado prévio de `Application.EnableEvents`.
  - Atualizar prioritariamente as tabelas finais da pasta de trabalho.
- [ ] **Otimização do `modulo1.vba`:**
  - Substituir o loop de cópia linha a linha por cópia nativa de planilha (`wsOrigem.Copy`) com desconexão de fórmulas para valores.
  - Obter áreas tabulares via `ws.ListObjects(1).Range` para prevenir truncamento de colunas.
- [ ] **Correção da Aba `Validações`:**
  - Corrigir as células D5 e D6 com a fórmula canônica neutra:
    `=LEFT(CELL("filename",A1),FIND("[",CELL("filename",A1))-1)`
  - Eliminar referências fixas ao disco `C:\Thelmo\`.

### Fase 2: Governança do Cálculo e Regras de Negócio

- [ ] **Integração de Feriados no Power Query M:**
  - Criar consulta lendo a tabela de feriados a partir da `TabelaB-V2.xlsx`.
  - Atualizar `fnCalcularDataSLA.m` para checar se a data analisada consta na lista de feriados do estado/município correspondente.
- [ ] **Criação de Tabela de Auditoria e Inconsistências:**
  - Adicionar consulta no Power Query que consolida registros com pendências (ex.: situação "CO" sem fechamento, atividade sem PPU correspondente, localidade sem galpão associado).
  - Exibir o contador de pendências no Painel para conferência antes da emissão da memória.
- [ ] **Parametrização Centralizada de Contratos:**
  - Criar tabela na aba `Validações` com colunas: `Contrato`, `Número SAP`, `Sufixo`, `Status`.
  - Alterar a macro `GerarArquivosPorContrato` para ler essa tabela dinamicamente.

### Fase 3: Modernização Visual e Experiência do Usuário (Dashboard)

- [ ] **Ajuste do Protótipo de Dashboard:**
  - Corrigir a sobreposição de linhas e layouts em `modDashboard.bas`.
  - Implementar cards dinâmicos conectados via fórmulas (`SOMASE`, `CONT.SES`, `SUBTOTAL`), dispensando reprocessamento via macro para filtros visuais.
  - Disponibilizar tabela de arquivos gerados com hiperlinks amigáveis e botão direto para "Abrir Pasta MEDIÇÃO".

### Fase 4: Limpeza e Redução de Complexidade

- [ ] **Depreciação dos Códigos Legados:**
  - Manter os módulos legados arquivados externamente em `Codes/VBA/FuncoesAntigas/` e removê-los do arquivo `.xlsm` em produção, reduzindo o tamanho do arquivo, o tempo de abertura e alertas de segurança.

---

*Documento gerado para servir de base técnica para a modernização e auditoria do projeto.*

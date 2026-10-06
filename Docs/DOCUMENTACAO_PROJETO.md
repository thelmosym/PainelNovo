# Documentação do projeto

## Painel de Controle MemoriaPetrobras V6.xlsm

**Data da documentação:** 22/09/2026 (Revisado e Atualizado em 29/09/2026)  
**Escopo:** análise estática do workbook, dos códigos exportados e da estrutura de pastas em `D:\PainelNovo`.

> Esta documentação descreve o que foi identificado nos arquivos disponíveis. A execução real das macros, a atualização das consultas e a conferência dos valores calculados ainda precisam ser validadas no Excel.

## 1. Visão geral

O projeto é uma solução de controle de medição, desempenho e faturamento de serviços associados a contratos Petrobras. O workbook consolida arquivos de monitoramento operacional, aplica regras de período, prazo/SLA, FDM, quantidade executada, armazenamento e frete, e gera arquivos de Memória de Cálculo por contrato.

O arquivo principal é:

- `Painel de controle MemoriaPetrobras V6.xlsm`

O sistema possui dois fluxos que coexistem:

1. **Fluxo Power Query M:** lê arquivos e tabelas de referência, transforma os dados, calcula campos derivados e carrega resultados em tabelas/abas do workbook.
2. **Fluxo VBA legado e novo:** executa validações, cálculos, atualização de conexões, rotinas operacionais e exportação dos resultados para arquivos `.xlsx`.

A existência dos dois fluxos indica uma migração gradual da lógica originalmente implementada em VBA para Power Query M. Por isso, algumas regras aparecem em versões equivalentes nas duas linguagens.

## 2. Estrutura de pastas

### Arquivos principais

- `Painel de controle MemoriaPetrobras V6.xlsm`: workbook principal, com macros, consultas, tabelas e abas de resultado.
- `DOCUMENTACAO_PROJETO.md`: este documento.
- `resumo_projeto_medicao_13-08-2026.txt`: histórico técnico e registro de decisões do projeto.

### `Codes/`

Contém os códigos exportados para versionamento e análise.

- `Codes/VBA/FuncoesAntigas/`: módulos VBA originais/legados.
- `Codes/VBA/Funções Novas/`: rotinas novas de atualização Power Query e geração de arquivos.
- `Codes/CodigosM/`: consultas e funções Power Query M exportadas.
- `Codes/TXT/`: cópias textuais de módulos e procedimentos antigos.
- `Codes/NovoPainelMonitoramento.md`: anotações sobre o mapeamento de colunas do painel.
- `Codes/readme.md`: atualmente vazio.

### `Monitoramento/`

Contém os arquivos de monitoramento individual que alimentam a consolidação. Foram observados arquivos como:

- `Monitoramento individual_A4UU_v4.xlsm` (e versões anteriores v3.0)
- `Monitoramento individual_DPBR_v4.xlsm`
- `Monitoramento individual_GPZ1_v4.xlsm`
- `Monitoramento individual_GQ6S_v4.xlsm`
- `Monitoramento individual_S2IJ_v4.xlsm`

A consulta `Transformar Arquivo.m` abre os arquivos e procura a tabela `Monitoramento`.

### `Tabelas Especificas/`

Contém tabelas de referência utilizadas nos cruzamentos e cálculos, especialmente a `TabelaB-V2.xlsx`. Essas referências fornecem localidades, centros, contratos, grupos, prazos, PPU, FDM, fatores de conversão e informações de origem/destino.

### `Medição - Contratada/`

Contém arquivos de medição contratada separados por contrato/região:

- `Medição - Contratada/IRON/`
- `Medição - Contratada/PA/`

As consultas `ConsolidadoES`, `ConsolidadoRJ`, `ConsolidadoSP` e `ConsolidadoPA` carregam esses arquivos.

### `MEDIÇÃO/`

Pasta de saída dos arquivos gerados pela macro `GerarArquivosPorContrato`. Os arquivos de saída são separados por contrato e normalmente contêm as abas `MC`, `ARM`, `DADOS` e `FRETE`.

## 3. Objetivo de negócio

O processo transforma registros operacionais de solicitações e atendimentos em dados de medição contratual. Em termos de negócio, a solução precisa responder:

- quais solicitações foram concluídas no período de faturamento;
- qual contrato e região são responsáveis pelo atendimento;
- qual quantidade foi solicitada e atendida;
- qual quantidade padronizada deve ser considerada para execução;
- se o prazo de atendimento foi cumprido;
- qual classificação FDM se aplica;
- quais valores e quantidades devem compor a Memória de Cálculo;
- quais arquivos finais devem ser entregues por contrato.

O resultado não é apenas uma listagem de monitoramento. Ele funciona como base operacional para medição, conferência e exportação contratual.

## 4. Entradas de dados

### 4.1 Monitoramento individual

Os arquivos da pasta `Monitoramento/` são a fonte operacional principal. Eles registram solicitações, contratos, atividades, itens, datas, quantidades, códigos OS, localidades, observações e situação.

A consulta `localMonitIndividualT2M.m` lê o caminho definido no workbook pelo nome `localMonitIndividualT2M`. A consulta `Transformar Arquivo.m` aplica a mesma estrutura aos arquivos encontrados na pasta.

Campos importantes observados na transformação:

- `Situação`;
- `Contrato`;
- `Descrição da atividade`;
- `Item`;
- `Aplicação`;
- `Código da solicitação`;
- `Chave solicitante`;
- `Data solicitação`;
- `D. abertura`;
- `D. fechamento`;
- `Prazo combinado`;
- `Qtd. Solicitada`;
- `Qtd. Atendida`;
- `Código OS`;
- `Localidade`;
- `Comentários`;
- `Observações`.

### 4.2 Tabelas de referência

A pasta `Tabelas Especificas/` fornece dados auxiliares para enriquecer os registros. Entre as consultas relacionadas estão:

- `Localidade`;
- `Localidade Com Endereço`;
- `Localidade Sem endereço`;
- `Centro`;
- `Contrato`;
- `Grupo`;
- `Galpão`;
- `validUFGalpao`;
- `valEnderecoGalpao`;
- `tabelaPrazoContrato`;
- `tabRefPrazoFDM`;
- `tabelaAtividadeItemPPU`;
- `tabela_FC`;
- `tabela_FC_arm`;
- `tabela_FC_digitalizacao`;
- `tabela_FC_migracao`;
- `PPU`;
- `PPU-IRON-LT1`;
- `PPU-PA-LT2`;
- `Origem_Destino`;
- `RefFDMHistorico`.

### 4.3 Parâmetros do período

O painel possui nomes definidos `PeriodoMes` e `PeriodoAno`. As consultas `PeriodoMes.m` e `PeriodoAno.m` leem esses valores do workbook.

A função `fnGerarPeriodo.m` converte o mês e o ano em um período de negócio que começa no dia 26 do mês anterior e termina no dia 25 do mês selecionado.

Exemplo identificado:

```text
Mês: Agosto
Ano: 2026
Período: 26/07/2026 à 25/08/2026
```

## 5. Fluxo Power Query M

### 5.1 Carregamento e expansão

A consulta principal `Painel T2M.m` executa, em linhas gerais:

1. lê os arquivos da pasta de monitoramento;
2. ignora arquivos ocultos;
3. invoca `Transformar Arquivo` para cada arquivo;
4. expande a tabela `Monitoramento`;
5. remove colunas não necessárias;
6. define tipos de texto, datas, horários e quantidades.

A consulta utiliza `Table.Buffer` em referências e subconjuntos para reduzir releituras durante merges e cálculos.

### 5.2 Filtros de negócio

A versão analisada filtra:

- `Situação = "CO"`, indicando registros concluídos;
- período de faturamento preenchido;
- contratos `IRON-LT1` ou `PA-LT2`;
- período de faturamento dentro do intervalo produzido por `fnGerarPeriodo`.

Esse filtro define o conjunto de dados que será considerado para a medição do período selecionado.

### 5.3 Limpeza e padronização

A consulta remove colunas operacionais que não seguem para o resultado final, renomeia `D. solicitação` para `Data solicitação`, converte tipos e normaliza textos.

Para localidades, são aplicadas operações como:

- conversão para maiúsculas;
- `Text.Clean`;
- `Text.Trim`.

Essa padronização é importante porque vários merges dependem de correspondência textual.

### 5.4 Enriquecimento por merges

Os registros são enriquecidos por junções com consultas de referência. O resultado pode incluir:

- centro;
- município;
- UF;
- contrato regional;
- galpão;
- endereço do galpão;
- grupo;
- prazo contratual;
- fator de conversão;
- linha de PPU;
- origem e destino;
- quilômetros e quilômetros adicionais.

As consultas `Painel otimizado.m` e `Painel T2M.m` representam versões importantes do pipeline. É necessário confirmar no Excel qual delas está efetivamente carregada em cada tabela do workbook.

### 5.5 Identificação de duplicidades e FDM

A lógica de FDM utiliza uma chave concatenada formada por:

- `Descrição da atividade`;
- `Código da solicitação`;
- `Código OS`.

Essa chave é usada para identificar ocorrências repetidas. A consulta e as funções de FDM distinguem a ocorrência que deve compor o cálculo das demais ocorrências relacionadas.

A função `fnClassificarFDM.m` classifica registros em categorias como:

- `Normal`;
- `Expresso`;
- `Isento`;
- `FDM definido`.

Também são produzidos indicadores como:

- `NP`: dentro do prazo;
- `FP`: fora do prazo;
- `LOG SLA`: explicação ou registro do resultado da regra.

### 5.6 Cálculo de SLA

As funções envolvidas são:

- `fnCalcularDataSLA.m`;
- `fnCalcularPrazoSLA.m`;
- `fnCalcularDataSLA`/rotinas equivalentes no VBA legado.

A lógica documentada considera:

- expediente das 08:00 às 17:00;
- intervalo de almoço das 12:00 às 13:00;
- dias úteis de segunda a sexta-feira;
- avanço para o próximo dia útil quando o tempo restante ultrapassa o expediente;
- prazos de 0, 0,5, 1, 2, 3, 4 e 5 dias, conforme a função de prazo.

A regra de `0,5` dia possui tratamento específico e deve ser validada com casos reais. A integração de feriados foi concluída na função `fnCalcularDataSLA.m`, que agora consome a base unificada `Tabelas Especificas/Feriados_2026_2030.csv`, cruzando feriados nacionais, estaduais e municipais por Município e UF.

### 5.7 Cálculo de QExec

A função `fnCalcularQExec.m` calcula uma quantidade padronizada de execução. As regras observadas incluem fatores específicos para atividades e itens, tais como:

- embalagens: fator `0,1`;
- item avulso: fator `0,017`;
- organização analítica: fator `1`;
- organização simples: fator `0,5`;
- atividades de passagem direta: quantidade sem conversão;
- migração e digitalização: fator vindo das tabelas de conversão;
- `MATERIAL PARA ARQUIVAMENTO` e `BAIXA PERMANENTE`: sem QExec;
- `DEVOLUCAO DE EMPRESTIMO`: regra dependente do item.

A função `fnCalcularQExecAgrupado.m` trata situações em que o cálculo precisa considerar um agrupamento de transporte/frete.

#### 5.7.1 Cálculo de QExec Agrupado para Frete (`fnCalcularQExecAgrupado.m`)

Para atividades de transporte/frete, o faturamento não é apurado individualmente por ordem de serviço quando múltiplas ordens compartilham o mesmo deslocamento. A função `fnCalcularQExecAgrupado.m` implementa o agrupamento e a deduplicação de frete:

- **Parâmetros da função:**
  - `tabela`: tabela de entrada já enriquecida com o `QExec` individual;
  - `colDataFechamento`: coluna de data de fechamento do atendimento;
  - `colLocalPetrobras`: coluna da localidade Petrobras (origem/destino);
  - `colLocalGuarda`: coluna da cidade do galpão da contratada;
  - `colLinhaServico`: coluna contendo a Linha de Serviço PPU (`Linha de serviço PPU`);
  - `colQExec`: coluna com o `QExec` individual da linha.

- **Composição da Chave de Agrupamento (`_ChaveAgrupamento`):**
  A chave é montada de forma segura contra nulos:
  `Data Fechamento | Localidade Petrobras | Galpão_Cidade | Linha de serviço PPU`.
  > **Importante:** A inclusão explícita da `Linha de serviço PPU` garante a separação estrita entre Frete Normal (`FRE-NRM`) e Frete Expresso (`FRE-EXP`). Dessa forma, atendimentos com SLAs e tabelas de preço distintas realizados na mesma rota e no mesmo dia não colidem nem são anulados mutuamente (conforme regra legada em `A_Calc_Medicao.bas:492`).

- **Regra de Cálculo ($\text{soma}/10$, piso 1):**
  1. Identifica a **primeira ocorrência** de cada chave (menor índice sequencial estável);
  2. Soma o `QExec` de todas as linhas que compartilham a chave;
  3. Se a $\text{soma} / 10 \le 1$, o resultado do grupo é `1`; caso contrário, é $\text{soma} / 10$;
  4. Atribui o resultado calculado **apenas na linha da primeira ocorrência** na coluna `QExecAgrupado`;
  5. Todas as demais linhas subsequentes do mesmo grupo recebem `null`.

### 5.8 FDM mensal por contrato

A consulta `TabelaFDMPorContrato.m` calcula indicadores mensais por contrato. O fluxo identificado:

1. obtém contratos distintos da consulta principal;
2. calcula `TSPE1`, `TSPR1`, `TSPE2` e `TSPR2`;
3. calcula `IAPFARQ1` e `IAPFARQ2`;
4. calcula valores brutos de FDM;
5. busca períodos anteriores;
6. valida se os períodos anteriores existem;
7. recalcula o FDM quando necessário.

Funções relacionadas:

- `fnCalcularTSPE_TSPR.m`;
- `fnCalcularIAPFARQ.m`;
- `fnCalcularFDMMensal.m`;
- `fnCalcularPeriodosAnteriores.m`;
- `fnValidarPeriodoAnterior.m`;
- `fnBuscarFDMAnterior.m`;
- `fnRecalcularFDM.m`.

## 6. Organização dos resultados no workbook

As abas identificadas podem ser agrupadas por responsabilidade:

### Controle e parâmetros

- `Painel`: seleção de mês/ano, comandos de atualização, status e links dos arquivos gerados.
- `Validações`: caminhos, parâmetros e verificações de configuração.

### Dados consolidados

- `DADOS`: registros detalhados da medição, com campos operacionais e derivados.
- `ARM`: dados e saldos de armazenamento.
- `Frete`/`FRETE`: informações de frete, incluindo atividades expressas e normais.
- `MC`: memória de cálculo e consolidação contratual.

### Abas por contrato/região

Foram observadas variações para:

- `PA-LT2`;
- `IRON-LT1-RJ`;
- `IRON-LT1-SP`;
- `IRON-LT1-ES`.

Os conjuntos seguem a ideia de:

- `MC_*`;
- `ARM_*`;
- `DADOS_*`;
- `FRETE_*`.

Também existem abas consolidadas como `ARM_IRON-LT1`, `DADOS_IRON-LT1`, `Frete_IRON-LT1` e `MC_IRON-LT1`.

### Referências e apoio

As consultas e abas de referência sustentam os cruzamentos de localidade, contratos, prazos, PPU, FDM, frete, armazenamento e fatores de conversão.

## 7. Fluxo VBA legado

Os principais módulos antigos estão em `Codes/VBA/FuncoesAntigas/`.

### `Planilha7.cls` / `fncPainelDeControle.bas`

Concentram procedimentos históricos do painel, incluindo validação, análise de dados, concatenação para FDM e preenchimento de resultados.

A rotina `analiseDeDados` valida, entre outros pontos:

- situação;
- atividade;
- item;
- datas;
- quantidades;
- localidade;
- município;
- UF;
- contrato;
- grupo;
- prazo;
- compatibilidade entre atividade e item.

### `fncGeral.bas`

Reúne funções gerais, mensagens e regras auxiliares. A função de cálculo de data SLA considera expediente, intervalo de almoço, fins de semana e, no fluxo legado, feriados pesquisados por município e UF.

### `A_Calc_Medicao.bas`

Responsável por partes do cálculo de medição e pagamento dos itens. Entre os procedimentos relacionados estão `Inicial_Importar_e_calcular_dados_Atendimentos`, `Organizar_dados` e `CalcularPagamentoItens`.

### `B_Calc_ARM.bas`

Relaciona-se ao cálculo de armazenamento. O procedimento `importar_e_calcular_ARM` é um dos pontos principais.

### `frete.bas`

Concentra regras e cálculos de frete.

### `fncTabelaA.bas` e `fncTabelaB.bas`

Atuam sobre tabelas de referência e validações usadas para determinar regras de atividade, contrato, localidade, prazo, grupo e outros parâmetros.

### `C_Export_dados.bas` e módulos de exportação

Executam exportações do fluxo legado. O procedimento `ExportarDadosParaNovoArquivoFinal` copia abas de resultado para um novo workbook, remove referências externas e salva o arquivo final como `.xlsx`.

## 8. Fluxo VBA novo

### 8.1 Atualização das consultas

O módulo [Codes/VBA/Funções Novas/modAtualizarConsultas.bas](Codes/VBA/Fun%C3%A7%C3%B5es%20Novas/modAtualizarConsultas.bas) contém `AtualizarTodasConsultasPowerQuery`.

A rotina:

1. usa `ThisWorkbook` como workbook de origem;
2. conta as conexões;
3. localiza a aba `Painel` e a célula de status `F9`;
4. solicita confirmação ao usuário;
5. desabilita eventos;
6. percorre `wb.Connections`;
7. tenta forçar `BackgroundQuery = False`;
8. atualiza cada conexão individualmente;
9. aguarda consultas assíncronas;
10. escreve o resultado em `F9` e na barra de status;
11. exibe um resumo em `MsgBox`.

Pontos de atenção:

- todas as conexões são processadas, inclusive auxiliares ou duplicadas;
- falhas parciais podem ser registradas, mas o status visual precisa diferenciar sucesso total de sucesso parcial;
- o estado anterior de `Application.EnableEvents` deveria ser preservado e restaurado;
- a execução real deve confirmar se todas as conexões listadas são necessárias.

### 8.2 Geração de arquivos por contrato

O módulo [Codes/VBA/Funções Novas/modGerarArquivos.bas](Codes/VBA/Fun%C3%A7%C3%B5es%20Novas/modGerarArquivos.bas) contém `GerarArquivosPorContrato`.

Contratos configurados no código:

| Contrato | Número SAP | Sufixo/estado |
| --- | ---: | --- |
| `PA-LT2` | `4600687006` | `BA` |
| `IRON-LT1-RJ` | `4600686987` | vazio |
| `IRON-LT1-SP` | `4600686987` | vazio |
| `IRON-LT1-ES` | `4600686987` | vazio |

A rotina:

1. verifica se o workbook foi salvo;
2. cria ou localiza a pasta `MEDIÇÃO` ao lado do workbook;
3. confirma os contratos e tipos de aba;
4. gera um arquivo por contrato;
5. copia as abas `MC`, `ARM`, `DADOS` e `FRETE`:
   - **Aba `MC`:** Preserva fórmulas dinâmicas do Excel (`SOMASE`, subtotais, multiplicadores e totais), redirecionando referências externas do workbook principal para as abas locais (`ARM!`, `DADOS!`, etc.);
   - **Abas `ARM` e `DADOS`:** Copiadas com dados e formatação completa, recriando as tabelas estruturadas (`ListObjects`) com a nomenclatura original (`ARM_<CONTRATO>`, `DADOS_<CONTRATO>`) para resolver referências estruturadas nas fórmulas da `MC`;
   - **Desconexão de Vínculos Externos (`BreakLink`):** Links para tabelas auxiliares não exportadas (`PPU` e `TabelaFDMPorContrato`) são convertidos em valores via `BreakLink`, eliminando mensagens de erro de vínculos quebrados ao abrir o arquivo;
   - **Aba `FRETE`:** Copiada com valores e formatação completa;
6. grava links em `Painel!F14:F17` (com sinalização por cores: Amarelo durante o processo, Verde Claro para sucesso e Vermelho para erros);
7. informa sucessos e falhas.

A função interna `GerarArquivoDoContrato` seleciona as abas de origem, orquestra a cópia com fórmulas na `MC`, recria as tabelas estruturadas e monta o nome do arquivo com contrato, número, data e hora.

### 8.3 Auditoria de Fontes Externas Power Query (`modAuditoriaFontesPQ.bas`)

O módulo [Codes/VBA/Funções Novas/modAuditoriaFontesPQ.bas](Codes/VBA/Fun%C3%A7%C3%B5es%20Novas/modAuditoriaFontesPQ.bas) implementa a rotina `AuditarFontesPowerQuery`.

Finalidade e funcionamento:
1. Percorre todas as consultas em `ThisWorkbook.Queries` e todas as conexões em `ThisWorkbook.Connections`;
2. Inspeciona as fórmulas em Linguagem M em busca de conectores de dados externos (`Folder.Files`, `File.Contents`, `Excel.Workbook`, `Csv.Document`, `Odbc.Query`, `Web.Contents`, `Sql.Database`, etc.);
3. Extrai os caminhos de pastas, arquivos e parâmetros utilizados por cada consulta;
4. Identifica o tipo de fonte, status da consulta (ativa, intermediária ou órfã) e relacionamentos com tabelas do workbook;
5. Cria/atualiza a aba `AUDITORIA_FONTES_PQ` com formatação profissional, tabela estruturada, resumo executivo e métricas de fontes locais vs. fontes parametrizadas.

### 8.4 Módulo de Auditoria Preventiva e Qualidade de Dados (`modAuditoriaLog.bas`)

O módulo [Codes/VBA/Funções Novas/modAuditoriaLog.bas](Codes/VBA/Fun%C3%A7%C3%B5es%20Novas/modAuditoriaLog.bas) implementa a rotina `MotorFiscalizadorDADOS`.

Destaques da implementação:
1. **Mapeamento Flexível de Cabeçalhos:** Localiza colunas independentemente de acentuação, maiúsculas/minúsculas ou quebras de linha (`_x000a_`);
2. **Processamento em Memória (RAM):** Carrega a base inteira em arrays na memória, fiscalizando milhares de registros em segundos;
3. **Classificação de Inconsistências:** Diferencia apontamentos entre `CRÍTICO` (datas invertidas, campos obrigatórios nulos) e `ALERTA` (campos incompletos sem impedimento imediato);
4. **Relatório Visual Interativo:** Gera ou atualiza a aba `LOG_CRITICAS` com contadores, cartões de métricas e hiperlinks diretos que navegam o usuário para a célula exata da inconsistência;
5. **Ajuste de Negócio Operacional:** A verificação de duplicidade de chaves `Atividade + Solicitação + OS` foi suprimida para permitir ordens legítimas com múltiplos itens/etapas complementares.

## 9. Fluxo resumido ponta a ponta

```text
Arquivos Monitoramento/*.xlsm
        |
        v
Transformar Arquivo + localMonitIndividualT2M
        |
        v
Painel T2M / Painel otimizado
        |
        +--> filtros de situação, contrato e período
        +--> padronização de textos e tipos
        +--> merges com TabelaB e referências
        +--> FDM, SLA, QExec e classificação
        |
        v
DADOS / ARM / FRETE / MC / abas por contrato
        |
        +--> conferência na aba Painel
        +--> GerarArquivosPorContrato
        |
        v
MEDIÇÃO/*.xlsx
```

## 10. Regras de negócio consolidadas

### Período de faturamento

O período não segue necessariamente o mês civil. Para um mês selecionado, começa no dia 26 do mês anterior e termina no dia 25 do mês selecionado.

### Situação

A consulta principal considera registros com `Situação = "CO"`. Isso significa que registros não concluídos ficam fora do conjunto principal de medição, salvo quando tratados por consultas auxiliares.

### Contrato

O contrato operacional pode ser enriquecido por UF, localidade, galpão e tabelas de referência. O código também diferencia o contrato-base `IRON-LT1` das regiões `RJ`, `SP` e `ES` em etapas de consolidação e exportação.

### SLA

O prazo depende da referência contratual e é convertido em data/hora limite respeitando expediente e dias úteis. A divergência entre a regra M e a regra VBA, especialmente para feriados e meio período, precisa ser tratada como ponto de governança do cálculo.

### FDM

O FDM utiliza classificação de serviço e desempenho no prazo. Duplicidades são controladas por uma chave de atividade, solicitação e OS para evitar contagem indevida.

### QExec

A quantidade executada pode ser direta ou convertida por fatores definidos para itens, atividades, migração, digitalização, embalagem e organização.

### Exportação

A saída operacional é um conjunto de arquivos por contrato, cada um com abas de memória de cálculo, armazenamento, dados e frete.

## 11. Riscos e pontos de atenção

### Caminhos absolutos

Algumas consultas exportadas usam caminhos absolutos, como `D:\PainelNovo\Monitoramento\` e caminhos específicos de usuário. Isso reduz a portabilidade do workbook.

**Melhoria recomendada:** manter os caminhos em células nomeadas ou em uma tabela de configuração e fazer as consultas M lerem esses parâmetros por `Excel.CurrentWorkbook()`.

### Conexões duplicadas, órfãs ou dessincronizadas em tabelas de saída

A macro de atualização percorre todas as conexões. Conexões duplicadas podem aumentar o tempo, causar mensagens de erro ou atualizar resultados que não são necessários para a operação do painel.

Além disso, em pastas de trabalho com dezenas de consultas, tabelas do Excel (`ListObjects`) podem perder a amarração correta com sua consulta ativa do Power Query:
- **Caso Real Diagnosticado:** Na aba `DADOS_IRON-LT1-RJ`, a tabela `DADOS_IRON_LT1_RJ` estava atrelada a uma conexão interna órfã (`connectionId="18"`, marcada como deletada), enquanto a consulta Power Query ativa estava na conexão `connectionId="22"`. Isso fazia com que forçar a atualização na tabela não refletisse os dados novos gerados pelo Power Query.
- **Solução / Mitigação:** Usar a macro de auditoria `AuditarFontesPowerQuery` (`modAuditoriaFontesPQ.bas`) para verificar a correspondência entre tabelas e conexões ativas.

### Divergência entre VBA e M (Feriados)

Originalmente, o VBA legado consultava feriados por município/UF, enquanto as funções M iniciais consideravam apenas dias úteis de segunda a sexta.
- **Status:** **Resolvido**. A função `fnCalcularDataSLA.m` foi modernizada e integrada com a base unificada `Feriados_2026_2030.csv`, considerando calendário nacional, estadual e municipal com filtros de Município e UF.

### Tratamento silencioso de erros em M

O uso de `try ... otherwise` evita que a consulta inteira falhe, mas pode gerar valores nulos ou logs que passam despercebidos.

**Melhoria recomendada:** criar uma tabela de erros de transformação com origem, chave do registro, etapa e mensagem.

### Correspondência textual

Merges dependem de nomes exatos de localidades, atividades, itens, UF, contratos e grupos. Espaços, acentos, valores nulos e grafias alternativas podem impedir correspondências.

### Estado global do Excel

As macros alteram `EnableEvents`, `ScreenUpdating` e `StatusBar`. Em caso de erro, esses estados devem ser restaurados ao valor original, não simplesmente forçados para `True` ou `False`.

### Arquivos gerados

A gravação de arquivos com data e hora é útil para histórico, mas pode acumular versões e deixar o usuário sem clareza sobre qual é a versão oficial.

**Melhoria recomendada:** registrar contrato, status, data de geração e caminho em uma tabela de controle no painel, além de separar arquivos atuais de históricos.

## 12. Melhorias recomendadas para o painel

A aba `Painel` hoje concentra parâmetros, botões, status e links. Ela pode evoluir para um dashboard operacional sem duplicar a lógica de cálculo.

### Indicadores sugeridos

- total de registros no período;
- quantidade solicitada;
- quantidade atendida;
- quantidade executada/QExec;
- registros por contrato;
- registros dentro e fora do prazo;
- classificação FDM;
- solicitações sem código OS;
- localidades sem correspondência;
- arquivos gerados com sucesso;
- erros de atualização.

### Componentes de UX

- cabeçalho com período calculado e horário da última atualização;
- botões `Atualizar painel` e `Gerar arquivos` com feedback de progresso;
- status textual combinado com cor, sem depender apenas de ícones;
- tabela de arquivos com contrato, status, data e ação `Abrir`;
- área de alertas para exceções de dados;
- filtros por contrato, atividade, localidade, FDM e situação;
- links curtos e amigáveis em vez de nomes técnicos longos;
- destaque para informações que exigem ação do usuário.

A implementação deve priorizar indicadores derivados das consultas existentes. O VBA deve orquestrar a interface e a exportação, enquanto a transformação e a regra de negócio permanecem preferencialmente no Power Query.

## 13. Validação recomendada no Excel

1. Abrir o workbook com macros habilitadas.
2. Confirmar os nomes definidos `PeriodoMes`, `PeriodoAno`, `localMonitIndividualT2M` e `localTabelaB`.
3. Confirmar qual consulta alimenta `DADOS`, `ARM`, `FRETE`, `MC` e as abas por contrato.
4. Atualizar as consultas pela interface do Excel.
5. Conferir se o período exibido corresponde ao filtro aplicado nos registros.
6. Comparar uma amostra com o resultado do fluxo VBA legado.
7. Testar um registro dentro do prazo e outro fora do prazo.
8. Testar prazo `0,5` dia e um prazo que atravesse almoço ou fim de semana.
9. Testar uma localidade sem correspondência na tabela de referência.
10. Interromper ou provocar erro em uma conexão e confirmar a restauração dos estados do Excel.
11. Executar `GerarArquivosPorContrato` e conferir as quatro abas de cada arquivo.
12. Confirmar se os links em `Painel!F14:F17` apontam para a última geração e exibem fundo verde claro.

## 14. Status e Resoluções de Itens do Projeto

- **Feriados no cálculo M de SLA:** [RESOLVIDO] Implementado via `fnCalcularDataSLA.m` integrando a base `Feriados_2026_2030.csv` com filtros de Município e UF.
- **Frete Expresso vs. Frete Normal:** [RESOLVIDO] Inclusão de `Linha de serviço PPU` na chave de `fnCalcularQExecAgrupado.m`, corrigindo agrupamento indevido no contrato `IRON-LT1-RJ` (restaurando as 2 viagens expressas).
- **Auditoria de fontes externas e conexões:** [IMPLEMENTADO] Criado o módulo `modAuditoriaFontesPQ.bas` para rastreamento de caminhos e conectores.
- **Regra de duplicidade em OS:** [RESOLVIDO] Removida a checagem rígida de duplicidade em `modAuditoriaLog.bas`, aceitando múltiplos itens por OS conforme a realidade operacional.
- **Sincronização de fórmulas no workbook:** [PADRONIZADO] Consultas `fnCalcularQExecAgrupado`, `Painel T2M` e `Painel otimizado` sincronizadas no workbook principal `Painel de controle MemoriaPetrobras V6.xlsm`.
- **Qual fluxo de exportação é oficial:** [RESOLVIDO] O fluxo novo por contrato (`modGerarArquivos.bas`) gera arquivos limpos em `MEDIÇÃO/` com as 4 abas contratuais (`MC`, `ARM`, `DADOS`, `FRETE`), mantendo fórmulas dinâmicas ativas na aba `MC` vinculadas a `ARM` e `DADOS`, recriando tabelas estruturadas e convertendo vínculos externos residuais via `BreakLink`.
- **Se todos os arquivos de monitoramento mantêm o mesmo esquema de colunas:** Validado via `Transformar Arquivo.m`.

## 15. Conclusão

O workbook é uma aplicação de medição contratual com três responsabilidades principais:

1. **consolidar dados operacionais**;
2. **aplicar regras de negócio de medição, SLA, FDM, QExec, armazenamento e frete**;
3. **produzir arquivos de Memória de Cálculo por contrato**.

A arquitetura atual é funcional, mas possui duplicação entre VBA e Power Query, dependência de caminhos locais, múltiplas versões de consultas e atualização ampla de conexões. A evolução mais segura é consolidar gradualmente as regras no Power Query M, manter o VBA como camada de orquestração e exportação, parametrizar caminhos e transformar a aba `Painel` em uma camada de acompanhamento operacional com indicadores e alertas.

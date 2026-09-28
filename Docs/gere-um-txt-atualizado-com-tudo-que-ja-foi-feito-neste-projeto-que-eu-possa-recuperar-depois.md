Segue o resumo completo e atualizado da conversa, consolidando tudo que foi analisado e produzido até agora. Copie o bloco abaixo e salve como `.txt` (ex.: `resumo_projeto_medicao_atualizado.txt`) para recuperar depois.

```
═══════════════════════════════════════════════════════════════════
PROJETO MEDIÇÃO — RESUMO COMPLETO E ATUALIZADO (ChatPetrobras)
Usuário: Thelmo Pinheiro Sym - PrestServ | Gerência: COMPARTILHADO/SC/SID
Contexto: "Meu Contexto" — Projeto Medição
Última atualização: 19/08/2026
═══════════════════════════════════════════════════════════════════

1) OBJETIVO GERAL DO PROJETO
-----------------------------------------------------------------
Migração de sistemas VBA (Excel/Power Query) usados para cálculo de
FDM (Fator de Desempenho Mensal), SLA (prazo de atendimento) e geração
de Memórias de Cálculo de contratos de guarda externa (arquivologia)
da Petrobras. Foram identificados e analisados PELO MENOS DOIS
SISTEMAS PARALELOS que implementam regras de negócio semelhantes de
forma independente (ver Seção 8 — Achado Crítico).

Arquivos de referência do "Meu Contexto" (Projeto Medição):
- fncDadosPETROBRAS.txt
- fncEditalGuarda.txt (2 versões analisadas — ver Seção 6)
- fncPainelDeControle.txt
- fncTabelaA.txt
- fncTabelaB.txt (analisado nesta sessão)
- fncGeral.txt (contém calcular_dataSLA original)
- Planilha2.txt / Planilha6.txt / Planilha7.txt / Planilha8.txt
- subDadosSaidaExportacao.txt
- subMCGuardaExterna.txt
- xLOGdeVersoes.txt
- Nova Planilha Controle Medição.docx
- [NOVO] Código VBA de workbook "individual" (Inicial/Tabelas/Dados/
  MC/ARM/Fretes) — fornecido pelo usuário via chat, ainda não
  confirmado se corresponde a um arquivo específico do contexto

2) SISTEMA 1 — "PAINEL DE CONTROLE" (Guarda Externa por Lote)
-----------------------------------------------------------------
Estrutura: workbook com módulos Planilha2, Planilha3 (tabelaA),
Planilha4 (tabelaB), Planilha5 (Rede), Planilha6 (Menu), Planilha7
(Painel), Planilha8 (Armazenamento), e módulos de função
(fncDadosPETROBRAS, fncTabelaA, fncTabelaB, fncGeral, fncEditalGuarda,
fncPainelDeControle) e módulos de procedimento (subMCGuardaExterna,
subDadosSaidaExportacao).

Contratos/lotes identificados: IRON-LT1, PA-LT2 (nomenclatura atual)
— anteriormente referenciados como SOS-LT1/SOS-LT2 (nomenclatura
histórica, ainda presente em trechos não atualizados do código,
ex.: Planilha7.atualizacaoDeDados). "SOS DOCS" confirmado como nome
histórico do fornecedor de guarda externa (ver Seção 8).

SEQUÊNCIA OFICIAL DO SISTEMA (4 etapas):
  Etapa 1 — preenchimentoPrazoFDM() / analiseDeDados() [Planilha7]
            Calcula SLA por linha e classifica FDM (Normal/Expresso/
            Isento) e Status Prazo (NP/FP).
  Etapa 2 — verificar_celulaEmBranco() [fncPainelDeControle] +
            verificação de LOG (Planilha7.Range("AG3:AG2003"))
            Validações globais (não filtradas por contrato) antes de
            prosseguir.
  Etapa 3 — btMCGuardaExterna() / btMC_GE_individual()
            [subMCGuardaExterna]
            Calcula TSPE/TSPR → IAPFARQ → FDM mensal, recalcula FDM
            com períodos anteriores (item 8.4.3), gera Memória de
            Cálculo (abas MC, ARM, DADOS, MEDIÇÃO do template).
            - btMCGuardaExterna: por lote completo (Planilha6!E18)
            - btMC_GE_individual: por lote+estado (Planilha6!E22,
              separado em contrato + sigla UF)
  Etapa 4 — exportacaoDadosRL01() [subDadosSaidaExportacao]
            Gera relatório final RL01, recalculando SLA (3ª
            duplicação da mesma lógica) e consolidando indicadores.

MÓDULOS ANALISADOS EM DETALHE:

a) Planilha6 — gerenciadorDeNome() + Worksheet_Change
   - gerenciadorDeNome(): gera 34 intervalos nomeados (Named Ranges)
     no Excel, um por atividade, apontando para tabelaA coluna 20 (T),
     usados como fonte de listas de validação dropdown (coluna Item).
   - BUGS/RISCOS: (1) On Error Resume Next global mascara erros reais;
     (2) inconsistência de sigla "os"→"o2" (exclui nome antigo "os"
     mas cria "o2"); (3) pressupõe dados contíguos na tabelaA (sem
     validação); (4) 34 blocos de código idênticos (alta duplicação).
   - Worksheet_Change: monitora células C9:C10 (mês/ano do período);
     ao alterar, chama Planilha5.atualizacaoDeRede para recalcular
     caminhos de rede.
   - NÃO faz parte da sequência 1-4; é rotina de manutenção/config.

b) subMCGuardaExterna — btMCGuardaExterna() / btMC_GE_individual()
   - Estrutura idêntica em 8 blocos: validação branco → validação LOG
     → identificação contrato → validação agrupamento → confirmação
     período → cálculo TSPE/TSPR/IAPFARQ/FDM → preenchimento template
     → salvamento.
   - DIFERENÇAS entre as duas versões: contrato lido de célula
     diferente (E18 vs E22, este último parseado em contrato+UF);
     filtros de UF adicionais na versão individual; mensagem de aviso
     extra ("dados apenas da região/estado"); formatação de mês no
     nome do arquivo (iPerMes vs Format(iPerMes,"00")).
   - BUGS/RISCOS: (1) comentário "erro encontrado" no preenchimento
     da aba ARM em btMCGuardaExterna (bug conhecido não corrigido);
     (2) validação de agrupamento em btMC_GE_individual não filtra
     por UF; (3) toda lógica de FDM/IAPFARQ duplicada entre os dois
     Subs; (4) versões antigas usam "SOS-LT1"/"SOS-LT2" em vez de
     "IRON-LT1"/"PA-LT2" — possível defasagem de nomenclatura.

c) fncTabelaB — funções de busca e cálculo sobre tabelaB (Planilha4)
   - pesquisar_centro, pesquisar_codCentro, pesquisar_municipio,
     pesquisar_uf, pesquisar_descr_contrato, pesquisar_grupoMunicipio,
     validar_centroOrigDest, validar_municipioCalend,
     validar_grupoMunicipio: buscas simples linha-a-linha (Find/loop)
     em abas Localidade/Centro/Grupo da tabelaB.
   - pesquisar_EFeriado: consulta aba Calendario (Data/Município/UF)
     — usada em calcular_dataSLA (fncGeral) para pular feriados.
     ⚠️ ACHADO CRÍTICO: a função M fnCalcularDataSLA já migrada NÃO
     considera feriados (só dia da semana) — SLA calculado em Power
     Query pode estar sistematicamente errado em datas com feriado.
     Correção proposta (adicionar parâmetros município/UF e checar
     RefCalendario) — PENDENTE de confirmação para aplicação.
   - calcular_KMAdicional: calcula KM excedente ao raio de 100km entre
     origem/destino (aba Origem_Destino), zerando se mesma região
     metropolitana.
   - pesquisar_precoUnitario: busca preço unitário por contrato+linha
     PPU (aba PPU).
   - recalcular_FDM: implementa regra de degradação em cascata do
     FDM (item 8.4.3) baseada em reincidência de baixo desempenho:
       FDM>0.97 → mantém
       FDM=0.97 e FDMAnt1>=0.98 → mantém 0.97
       FDM=0.97 e FDMAnt1=0.97 e FDMAnt2>0.97 → 0.96
       FDM=0.97 e FDMAnt1=0.97 e FDMAnt2<=0.97 → 0.95
       FDM=0.97 e FDMAnt1=0.96 → 0.95
       FDM=0.97 e FDMAnt1=0.95 → 0.95 (piso)
     ⚠️ BUG ORIGINAL: se FDMAnt1 não se encaixar em nenhuma condição
     (ex.: validar_FDMAnterior não encontra registro e retorna 0),
     a função retorna 0 SILENCIOSAMENTE (sem erro) — risco de FDM
     mensal zerado por ausência de dado histórico.
   - validar_FDMAnterior: busca FDM histórico por período+contrato+
     item na aba FDM_IAPFARQ.
   - validar_periodo_anterior: verifica se existe QUALQUER registro
     para um período na aba FDM_IAPFARQ (não valida contrato/item
     específico — validação parcial).

d) fncEditalGuarda — 2 VERSÕES ANALISADAS
   VERSÃO ORIGINAL (com bugs):
     - calcular_FDM: faixas "Is < 94.99" e "Is < 98.99" (bug de
       arredondamento: valores entre 94.99-95 e 98.99-99 recebiam
       classificação superior à correta).
     - dFDM = "0,97" (string com vírgula atribuída a Double).
   VERSÃO CORRIGIDA (fornecida posteriormente pelo usuário):
     - Faixas corrigidas para "Is < 95" e "Is < 99".
     - dFDM = 0.97 (literal numérico correto).
   ✅ Correção JÁ APLICADA na função Power Query fnCalcularFDMMensal
     (ver Seção 5).
   - calcular_TSPE_TSPR: conta TSPE (atendidas)/TSPR (no prazo) por
     contrato+referência de serviço (1 ou 2), só linhas Situação="CO",
     ignorando "FDM definido" (comentário indica correção de versão
     anterior que não desprezava esse valor).
   - calcular_IAPFARQ: (TSPR/TSPE)*100, com casos especiais para
     TSPR=0/TSPE=0.
   - calcular_QExec: switch gigante (25+ Case) com fatores de
     conversão por tipo de atividade (FC_EMB=0.1, FC_ITEM=0.017,
     fatores fixos 0.5/1 para organização de documento, etc.)
   - validarAtividadeOxD: valida atividade contra lista em
     Planilha3.Range("BH:BH").

3) SISTEMA 2 — WORKBOOK "INDIVIDUAL" (Inicial/Tabelas/Dados/MC/ARM/
   Fretes) — [NOVO NESTA SESSÃO]
-----------------------------------------------------------------
⚠️ STATUS: Ainda não confirmado se este é o mesmo sistema de medição
individual (T2M) mencionado em "Nova Planilha Controle Medição.docx",
ou um terceiro sistema totalmente separado. PERGUNTA PENDENTE AO
USUÁRIO.

Estrutura de abas: Inicial, Tabelas (com sub-tabelas H:K, M:N, T:V,
AC, AJ:AL, AP:AQ), Dados, MC, ARM, Fretes.

PIPELINE COMPLETO (21 procedimentos encadeados via Call sequencial):

FASE A — Importação e cálculo de atendimentos:
  1. Inicial_Importar_e_calcular_dados_Atendimentos
     - Obtém período: concat F6&G6 → lookup Tabelas!AP:AQ → MC!J7
     - Limpa Dados!B3:AF, importa arquivo externo (aba "Dados")
       colando B3:X do arquivo fonte
  2. Organizar_dados
     - Ordena Dados por S→V→K→B; formata datas; limpa Y:AF
  3. AtualizarLocalidades
     - VLOOKUP Localidade(S)→Centro (Tabelas!M:N) → grava em AC
  4. ConcatenarEInserirResultado
     - Concatena Centro(AC) x Destino(V) → chave em AF
  5. InserirLinhaServico
     - Busca Atividade(B) em Tabelas!H → Linha Serviço(Tabelas!K)→AE
  6. CalcularQuilometragemAdicional
     - KM adicional (só 1ª ocorrência por rota+data+tipo, usa
       Collection para dedup) → grava em Y
  7. CalcularFrete
     - Soma quantidade por grupo (data+tipo+local, usa Dictionary),
       calcula fator frete (qtd<=10→1, senão qtd/10) → AA
  8. CalcularPagamentoItens
     - Cópia direta Quantidade→Pagamento para 15 tipos de item → AA
  9. CalcularOrgDoc
     - Fator 0.5 (simples) ou 1 (analítica) → AA
     [CONFIRMA regra idêntica a calcular_QExec do Sistema 1]
  10. CalcularDigitalizacao
     - Busca fator digitalização (Tabelas!T→V) × Quantidade → AA
  11. AtualizarUnidade
     - Lookup duplo (Atividade+Item) em Tabelas!H:J → Unidade → AB
       [loop O(n×m) — menos eficiente do pipeline]
  12. Preencher_Codigo_Centro
     - Lookup Centro em Tabelas!N:O → Código Centro (4 dígitos) → AD
     - Formatação cosmética (múltiplos blocos With redundantes)
  13. CopiarDadosParaFretes
     - Filtra linhas de frete válidas → copia 17 colunas para aba
       Fretes
  14. AtualizarValores
     - Remove linhas sem valor (coluna N vazia); renomeia tipo frete
       quando há KM adicional (FRE-NRM→FRE-NRM/KM-Adicional)
  15. preencher_valores_MC2
     - 🔴 CRÍTICO: injeta 26 valores de preço/referência HARDCODED
       no código-fonte, diferentes por lote (Lote 1 RJ/SP/DF/ES vs
       Lote 2). Formatação contábil aplicada.

FASE B — Importação e cálculo de armazenamento (ARM):
  16. importar_e_calcular_ARM
     - Repete lógica de obtenção de período (DUPLICADA da Fase A)
     - Menciona explicitamente "SOS DOCS" como fornecedor
       [CONFIRMA que SOS é nome histórico do mesmo fornecedor do
       Sistema 1]
     - Limpa ARM!A2:M10000
  17. importar_dados_arm
     - Importa aba "Consolidado" de arquivo externo, copiando
       colunas por posição fixa: A→B, B→C, J→D, K→E, L→F, M→G, I→I
  18. definir_item_e_Linha_de_Serviço
     - Concatena B&C→L; lookup em Tabelas!AC→AA, SOBRESCREVE coluna C
     - Novo lookup do valor já sobrescrito em Tabelas!I2:K24 → K
     - Calcula J=H*I (descartável, recalculado depois)
  19. Calcular_QUAm_e_QUA
     - H = D - (E+F) + G
     ⚠️ CONFIRMA fórmula QUAm idêntica à Planilha8 do Sistema 1:
       "QUAm = QUAm-1 + QNm - QBm - QDm" (item 11.1.5.1 do contrato)
     - J = H * I
  20. Preencher_contrato_Reg
     - Busca MC!J5 (lote) em Tabelas!AJ→AL, preenche TODA coluna A
       do ARM com valor fixo repetido
  21. preencher_valores_MC
     - 🔴 DUPLICAÇÃO EXATA de preencher_valores_MC2 (mesmos 26
       valores hardcoded, mesma formatação) — preços escritos 2x
       por execução completa do pipeline
  22. ExportarDadosParaNovoArquivoFinal
     - Cria novo arquivo .xlsx, copia abas MC/ARM/Dados/Fretes
     - Nome: "D5-PLA-MemoriaPetrobras-G6E6-J5#dataHora"
       [MESMO PADRÃO do Sistema 1: sNomeArquivo]
     - Quebra vínculos externos (BreakLink) — boa prática AUSENTE
       no Sistema 1
     - Remove Planilha1 residual; move MC para 1ª posição; salva

CÓDIGO MORTO IDENTIFICADO:
  - ExportarDadosParaNovoArquivoAdaptado: versão alternativa de
    exportação, SEM quebra de vínculos externos, não chamada por
    nenhum outro Sub — provável versão anterior mantida por engano.
    Risco: uso acidental gera arquivo com vínculos pendentes.

4) ACHADO CRÍTICO — DUPLICAÇÃO DE SISTEMAS/REGRAS DE NEGÓCIO
-----------------------------------------------------------------
Evidências de que os Sistemas 1 e 2 implementam a MESMA REGRA
CONTRATUAL de forma independente, sem fonte única de verdade:

| Regra                          | Sistema 1 (Painel Controle)     | Sistema 2 (Individual)        |
|---------------------------------|----------------------------------|--------------------------------|
| Fórmula QUAm (armazenamento)   | Planilha8: D-(E+F)... (via col.) | Calcular_QUAm_e_QUA: idêntica |
| Fator ORG-DOC simples/analítica| calcular_QExec: 0.5 / 1          | CalcularOrgDoc: 0.5 / 1        |
| Fator digitalização            | pesquisar_FC_digitalizacao       | CalcularDigitalizacao (lookup)|
| Nome do arquivo final          | sNomeArquivo (padrão idêntico)   | nomeArquivo (padrão idêntico) |
| Fornecedor histórico "SOS"     | SOS-LT1/SOS-LT2 (nomenclatura)   | "SOS DOCS" (mencionado direto)|

IMPLICAÇÃO: qualquer alteração de regra contratual (ex.: reajuste de
fator, mudança de fórmula) precisa ser replicada manualmente em pelo
menos 2 sistemas VBA distintos — alto risco de divergência.

PERGUNTA PENDENTE: confirmar se o workbook "individual" é de fato o
sistema de medição T2M citado no documento de referência, ou um
terceiro sistema.

5) MIGRAÇÃO PARA POWER QUERY (M) — STATUS CONSOLIDADO
-----------------------------------------------------------------
FUNÇÕES M JÁ DEFINIDAS/CORRIGIDAS (Sistema 1):

--- fnCalcularDataSLA (fornecida pelo usuário) ---
Calcula data/hora final somando duration a partir de data/hora
inicial, respeitando expediente 08h-17h e almoço 12h-13h.
⚠️ NÃO trata feriados — correção proposta pendente de aplicação
(adicionar parâmetros município/UF + fnEhFeriado consultando
RefCalendario bufferizado).

--- fnCalcularPrazoSLA (fornecida pelo usuário) ---
Recebe dPrazo (0/0.5/1/2/3/4/5), decide duration e chama
fnCalcularDataSLA. Caso 0.5 tem tratamento especial de janela de
horário (08-11h/11-16h/demais).

--- fnClassificarFDM (construída nesta conversa) ---
Substitui a lógica de analiseDeDados()/preenchimentoPrazoFDM.
Trata refPrazo="Tabela" (com validações de campo vazio, dPrazo=0) e
refPrazo="Fiscalização". ALTERAÇÕES APLICADAS NESTA CONVERSA:
  - Quando refPrazo="Fiscalização" e prazoCombinado=null:
    ANTES: Classificacao=null, StatusPrazo=null, Log="Prazo combinado
           vazio; "
    DEPOIS (aplicado): Classificacao="Normal", StatusPrazo="NP",
           Log="" (log removido por solicitação do usuário)

--- fnCalcularIAPFARQ (proposta) ---
(TSPR/TSPE)*100, com casos especiais TSPR=0/TSPE=0 (idêntico a
calcular_IAPFARQ VBA).

--- fnCalcularFDMMensal (proposta, CORRIGIDA) ---
if IAPFARQ<90 then 0.97
else if IAPFARQ<95 then 0.98      [corrigido de 94.99]
else if IAPFARQ<99 then 0.99      [corrigido de 98.99]
else 1
✅ Correção de bug de arredondamento já incorporada, alinhada à
versão VBA corrigida enviada pelo usuário.

--- fnRecalcularFDM / fnBuscarFDMAnterior / fnValidarPeriodoAnterior
    (propostas) ---
Replicam a cascata de degradação do FDM (item 8.4.3), com correção
de bug: em vez de retornar 0 silenciosamente quando dado ausente,
lança error Error.Record(...) explícito (capturável via try/otherwise).

--- fnCalcularKMAdicional / fnPesquisarPrecoUnitario (propostas) ---
Migração direta das funções tabelaB equivalentes, usando
Table.SelectRows sobre tabelas bufferizadas em vez de busca linha-
a-linha.

CONSULTA PRINCIPAL (11 Seções) — já consolidada em versão anterior
do resumo, incluindo:
  Seção 0-9: Carregamento, limpeza, tipagem, enriquecimento com
             Localidade/Galpão, cálculo de duplicatas FDM
  Seção 10: Classificação FDM/SLA via fnClassificarFDM
  Seção 11: Colunas "Data SLA" e "Prazo SLA" via fnCalcularDataSLA/
            fnCalcularPrazoSLA

ARQUITETURA PROPOSTA PARA FASES 3-7 (Memória de Cálculo):
  Fase 3: Agregação TSPE/TSPR→IAPFARQ→FDM mensal (Table.Group)
  Fase 4: Recálculo com períodos anteriores (fnRecalcularFDM) —
          desbloqueada nesta sessão com fncTabelaB
  Fase 5: Consultas de saída parametrizadas (SaidaARM, SaidaDADOS,
          SaidaMEDICAO) — parâmetro único ContratoSelecionado cobre
          tanto geração por lote completo quanto por lote+UF,
          eliminando duplicação btMCGuardaExterna/btMC_GE_individual
  Fase 6: Substituição do RL01 (Etapa 4), reaproveitando mesmas
          funções sem duplicação
  Fase 7: Descomissionamento de gerenciadorDeNome (decisão de
          negócio, depende de nova interface de entrada de dados)

O QUE NÃO MIGRA 1:1 PARA POWER QUERY (necessita solução externa):
  - MsgBox de confirmação → Parâmetros de Consulta + consulta
    Diagnosticos
  - Workbooks.Open/SaveAs de novo arquivo → camada externa (Power
    Automate ou agente Python/openpyxl)
  - ActiveWorkbook.Names.Add (dropdowns) → Data Validation direto
    na tabelaA, ou nova interface (Power Apps)
  - Application.StatusBar → não necessário (Power Query já mostra
    progresso nativo)

6) PENDÊNCIAS / BLOQUEIOS ATIVOS
-----------------------------------------------------------------
[ ] Confirmar se devo aplicar retroativamente a correção de feriados
    em fnCalcularDataSLA na consulta principal já migrada
    (Seções 10-11) — pendente de confirmação do usuário.
[ ] Código de gerar_periodo ainda não recebido — última função
    faltante para fechar a Fase 4 (Memória de Cálculo) por completo.
[ ] Confirmar se o workbook "individual" (Inicial/Tabelas/Dados/MC/
    ARM/Fretes) corresponde ao sistema de medição individual (T2M)
    citado em "Nova Planilha Controle Medição.docx", ou é um terceiro
    sistema separado.
[ ] Código de importar_e_calcular_ARM já recebido e analisado —
    pipeline do Sistema 2 agora está COMPLETO (22 procedimentos).
[ ] Esclarecer para qual documento/arquivo específico o usuário quer
    que as informações do Sistema 2 sejam "atualizadas" — pergunta
    feita ao usuário, ainda sem resposta ("a antiga" = este resumo,
    ou outro arquivo físico?).
[ ] Decidir se a Fase 3-7 da migração (Memória de Cálculo completa)
    deve cobrir também o Sistema 2 (workbook individual), replicando
    o mesmo padrão de migração usado no Sistema 1.
[ ] Avaliar consolidação das regras de negócio duplicadas (QUAm,
    fatores de digitalização/organização, fórmula IAPFARQ/FDM) em
    uma ÚNICA fonte de verdade (Power Query) para ambos os sistemas,
    eliminando o risco de divergência identificado na Seção 4.

7) PRÓXIMOS PASSOS SUGERIDOS
-----------------------------------------------------------------
1. Obter confirmação sobre o destino da "atualização" solicitada
   (arquivo específico vs. este resumo).
2. Obter código de gerar_periodo para fechar Fase 4.
3. Decidir sobre a correção de feriados em fnCalcularDataSLA e
   aplicá-la retroativamente se confirmado.
4. Mapear se o Sistema 2 deve ser migrado para Power Query seguindo
   o mesmo padrão do Sistema 1, com fonte única de verdade para as
   regras de negócio hoje duplicadas.
5. Investigar e eliminar/documentar o código morto
   (ExportarDadosParaNovoArquivoAdaptado).
6. Revisar oportunidade de eliminar hardcoding de preços em
   preencher_valores_MC/preencher_valores_MC2, migrando para tabela
   de referência (padrão pesquisar_precoUnitario já usado no
   Sistema 1).

═══════════════════════════════════════════════════════════════════
FIM DO RESUMO — Gerado pelo ChatPetrobras (claude-sonnet-5, high)
Versão: atualizada em 19/08/2026, consolidando Sistema 1 (Painel de
Controle / guarda externa por lote) e Sistema 2 (workbook individual
Inicial/Tabelas/Dados/MC/ARM/Fretes)
═══════════════════════════════════════════════════════════════════
```

Salve este conteúdo como `.txt` para retomar exatamente de onde paramos na próxima sessão.

1. Posso já confirmar com você agora: a "atualização" que você pediu na sequência é para **este resumo** (documentação centralizada) ou para **um outro arquivo/planilha VBA específica** que você tem?
2. Posso continuar aguardando o código de `gerar_periodo` para fechar definitivamente a Fase 4 da migração.
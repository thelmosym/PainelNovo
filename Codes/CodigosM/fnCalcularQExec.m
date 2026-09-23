let
    // -------------------------------------------------------------------
    // fnCalcularQExec
    // -----------------------------------------------------------------
    // Equivalente M da função VBA calcular_QExec (módulo
    // fncEditalGuarda) — usada para medição/faturamento das atividades
    // de guarda externa (ver resultado no template de Memória de
    // Cálculo, aba DADOS, coluna AA).
    //
    // O QUE FAZ:
    //   Converte a quantidade física solicitada (qtd) na quantidade
    //   padronizada de medição (QExec), de acordo com a atividade e o
    //   item, seguindo diferentes regras de conversão por padrão de
    //   atividade:
    //
    //   PADRÃO A — Atividades tipo "Embalagem" (COLETA/ENTREGA de
    //   embalagem): se qtd <= 10 e sem agrupamento, aplica piso mínimo
    //   de 1 unidade; caso contrário, qtd * 0,1 (fator FC_EMB).
    //
    //   PADRÃO B — Atividades tipo "Item avulso" (COLETA/ENTREGA de
    //   item avulso): se qtd <= 60 e sem agrupamento, piso mínimo de 1
    //   unidade; caso contrário, qtd * 0,017 (fator FC_ITEM).
    //
    //   PADRÃO C — Passagem direta (sem conversão): QExec = qtd, para
    //   atividades como TRANSFERENCIA DE ACERVO, DESTRUICAO SEGURA DE
    //   DOCUMENTO, HIGIENIZACAO DE DOCUMENTO, entre outras.
    //
    //   PADRÃO D — Organização de documento, com fator fixo:
    //     ANALITICA ? fator 1 | SIMPLES ? fator 0,5
    //
    //   PADRÃO E — Migração/Digitalização: o fator de conversão (FC)
    //   é lido diretamente do parâmetro "fc" (valor já pré-calculado
    //   em uma coluna "FC" na planilha principal, via lookup/merge
    //   externo — decisão tomada nesta conversa para simplificar a
    //   função, em vez de recriar internamente as funções VBA
    //   pesquisar_FC_migracao/pesquisar_FC_digitalizacao).
    //
    //   PADRÃO F — INDEXACAO DE DOCUMENTO: regra especial que NÃO
    //   verifica agrupamento (diferente dos Padrões A/B) — replica
    //   fielmente o comportamento do VBA original, ainda pendente de
    //   confirmação se é regra intencional ou lacuna do código legado.
    //
    //   ATIVIDADES SEM CÁLCULO — "MATERIAL PARA ARQUIVAMENTO" e
    //   "BAIXA PERMANENTE" retornam NULL diretamente: ambas entram na
    //   Memória de Cálculo como registro informativo, mas sem nenhum
    //   valor de QExec associado (replica o "GoTo proximaLinha" do
    //   VBA original, que pulava todo o bloco de cálculo para essas
    //   atividades específicas).
    //
    // ? CORREÇÕES APLICADAS EM RELAÇÃO AO VBA ORIGINAL:
    //   1. Removidos branches ElseIf/Else redundantes presentes em 6
    //      Case do VBA original (calculavam o mesmo valor duas vezes).
    //   2. "ENTREGA DE ITEM NORMAL" agora arredonda a 3 casas
    //      decimais, igual a "ENTREGA DE ITEM EXPRESSO" (o VBA
    //      original tinha essa inconsistência de precisão entre as
    //      duas atividades análogas).
    //   3. Fallback explícito com "error" para qualquer atividade não
    //      mapeada nesta função — em vez de retornar 0 silenciosamente
    //      (bug real identificado em produção: "BAIXA PERMANENTE" e
    //      "MATERIAL PARA ARQUIVAMENTO" geravam esse erro até serem
    //      tratadas explicitamente como exceções sem cálculo).
    //   4. "DEVOLUCAO DE EMPRESTIMO" com valor de Item não reconhecido
    //      (diferente de "Embalagem"/"Item avulso") também lança erro
    //      explícito, em vez de zerar silenciosamente.
    //   5. Uso de Number.Round em vez de Format(...) — evita
    //      dependência do separador decimal do locale do sistema, e
    //      mantém o tipo numérico do início ao fim do cálculo.
    //
    // PARÂMETROS:
    //   descricaoAtividade - texto com o nome da atividade (ex.:
    //                         "COLETA DE EMBALAGEM")
    //   item                - texto com o item (ex.: "Embalagem",
    //                         "Item avulso") — usado apenas por
    //                         DEVOLUCAO DE EMPRESTIMO
    //   qtd                 - quantidade física solicitada/atendida
    //   agrupamento         - texto do código de agrupamento da linha
    //                         (vazio/null se a linha não pertence a
    //                         um agrupamento)
    //   fc                  - fator de conversão pré-calculado (coluna
    //                         "FC" da planilha principal), usado
    //                         apenas pelas atividades de
    //                         Migração/Digitalização
    //
    // RETORNO:
    //   nullable number — a quantidade padronizada de medição (QExec),
    //   ou null para atividades sem cálculo aplicável
    // -------------------------------------------------------------------

    fnCalcularQExec = (
        descricaoAtividade as text,
        item as nullable text,
        qtd as nullable number,
        agrupamento as nullable text,
        fc as nullable number
    ) as nullable number =>
        let
            FC_EMB = 0.1,
            FC_ITEM = 0.017,

            SemAgrupamento = agrupamento = null or agrupamento = "",

            // ? CORREÇÃO DEFINITIVA: qtd nula é tratada ANTES de
            // qualquer comparação "qtd <= X" ou multiplicação, evitando
            // que null se propague para dentro de um "if" e cause o
            // erro "não conseguimos converter null em Logical/Number".
            // Decisão: qtd nula é tratada como 0 (equivalente a "nenhum
            // item atendido" naquela linha).
            qtdValidada = if qtd = null then 0 else qtd,

            AtividadesEmbalagem = {
                "COLETA DE EMBALAGEM",
                "ENTREGA DE EMBALAGEM EXPRESSO",
                "ENTREGA DE EMBALAGEM NORMAL"
            },

            AtividadesItemAvulso = {
                "COLETA DE ITEM AVULSO",
                "ENTREGA DE ITEM EXPRESSO",
                "ENTREGA DE ITEM NORMAL"
            },

            AtividadesPassagemDireta = {
                "INSERCAO DE ITEM AVULSO",
                "PESQUISA DE ITEM AVULSO EXPRESSO",
                "PESQUISA DE ITEM AVULSO NORMAL",
                "DESTRUICAO SEGURA DE DOCUMENTO",
                "TRANSFERENCIA DE ACERVO",
                "CATALOGACAO DE EMBALAGEM",
                "CATALOGACAO DE ITEM AVULSO",
                "HIGIENIZACAO DE DOCUMENTO",
                "CONVERSAO DE MIDIA",
                "GRAVACAO DE MIDIA",
                "COPIA DE MIDIA",
                "COPIA DE VIDEO"
            },

            AtividadesFCExterno = {
                "MIGRACAO DE ACERVO DOCUMENTAL",
                "DIGITALIZACAO DE MICROFILME",
                "DIGITALIZACAO DE MICROFICHA",
                "DIGITALIZACAO DE DOCUMENTO CONTRATUAL",
                "DIGITALIZACAO DE DOCUMENTO TECNICO",
                "DIGITALIZACAO DE DOCUMENTO BIBLIOGRAFICO",
                "DIGITALIZACAO DE DOCUMENTO CONTABIL E FINANCEIRO",
                "DIGITALIZACAO DE DOCUMENTO ADMINISTRATIVO"
            },

            AtividadesSemCalculo = {
                "MATERIAL PARA ARQUIVAMENTO",
                "BAIXA PERMANENTE"
            },

            Resultado =

                if List.Contains(AtividadesSemCalculo, descricaoAtividade) then
                    null

                else if descricaoAtividade = "ORGANIZACAO DE DOCUMENTO ANALITICA" then
                    Number.Round(1 * qtdValidada, 3)

                else if descricaoAtividade = "ORGANIZACAO DE DOCUMENTO SIMPLES" then
                    Number.Round(0.5 * qtdValidada, 3)

                else if List.Contains(AtividadesFCExterno, descricaoAtividade) then
                    if fc = null then
                        null   // Opção C — sem FC cadastrado, sem interromper a consulta
                    else
                        Number.Round(fc * qtdValidada, 3)

                else if descricaoAtividade = "INDEXACAO DE DOCUMENTO" then
                    if qtdValidada <= 10 then 1
                    else Number.Round(FC_EMB * qtdValidada, 3)

                else if List.Contains(AtividadesPassagemDireta, descricaoAtividade) then
                    Number.Round(qtdValidada, 3)

                else if List.Contains(AtividadesEmbalagem, descricaoAtividade) then
                    if qtdValidada <= 10 and SemAgrupamento then 1
                    else Number.Round(FC_EMB * qtdValidada, 3)

                else if List.Contains(AtividadesItemAvulso, descricaoAtividade) then
                    if qtdValidada <= 60 and SemAgrupamento then 1
                    else Number.Round(FC_ITEM * qtdValidada, 3)

                else if descricaoAtividade = "DEVOLUCAO DE EMPRESTIMO" then
                    if item = "Embalagem" then
                        if qtdValidada <= 10 and SemAgrupamento then 1
                        else Number.Round(FC_EMB * qtdValidada, 3)
                    else if item = "Item avulso" then
                        if qtdValidada <= 60 and SemAgrupamento then 1
                        else Number.Round(FC_ITEM * qtdValidada, 3)
                    else
                        error Error.Record(
                            "ItemNaoReconhecido",
                            "DEVOLUCAO DE EMPRESTIMO com Item não reconhecido: '" & Text.From(item) & "'. Esperado 'Embalagem' ou 'Item avulso'."
                        )

                else
                    error Error.Record(
                        "AtividadeNaoMapeada",
                        "Atividade '" & descricaoAtividade & "' não possui regra de cálculo de QExec definida em fnCalcularQExec."
                    )
        in
            Resultado
in
    fnCalcularQExec

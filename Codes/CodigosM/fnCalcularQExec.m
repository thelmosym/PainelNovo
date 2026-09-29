let
    // -------------------------------------------------------------------
    // fnCalcularQExec
    // -------------------------------------------------------------------
    // Equivalente M da funcao VBA calcular_QExec (modulo
    // fncEditalGuarda) - usada para medicao/faturamento das atividades
    // de guarda externa (ver resultado no template de Memoria de
    // Calculo, aba DADOS, coluna AA).
    //
    // O QUE FAZ:
    //   Converte a quantidade fisica solicitada (qtd) na quantidade
    //   padronizada de medicao (QExec), de acordo com a atividade e o
    //   item, seguindo diferentes regras de conversao por padrao de
    //   atividade:
    //
    //   PADRAO A - Atividades tipo "Embalagem" (COLETA/ENTREGA de
    //   embalagem): se qtd <= 10 e sem agrupamento, aplica piso minimo
    //   de 1 unidade; caso contrario, qtd * 0,1 (fator FC_EMB).
    //
    //   PADRAO B - Atividades tipo "Item avulso" (COLETA/ENTREGA de
    //   item avulso): se qtd <= 60 e sem agrupamento, piso minimo de 1
    //   unidade; caso contrario, qtd * 0,017 (fator FC_ITEM).
    //
    //   PADRAO C - Passagem direta (sem conversao): QExec = qtd, para
    //   atividades como TRANSFERENCIA DE ACERVO, DESTRUICAO SEGURA DE
    //   DOCUMENTO, HIGIENIZACAO DE DOCUMENTO, entre outras.
    //
    //   PADRAO D - Organizacao de documento, com fator fixo:
    //     ANALITICA ? fator 1 | SIMPLES ? fator 0,5
    //
    //   PADRAO E - Migracao/Digitalizacao: o fator de conversao (FC)
    //   e lido diretamente do parametro "fc" (valor ja pre-calculado
    //   em uma coluna "FC" na planilha principal, via lookup/merge
    //   externo - decisao tomada nesta conversa para simplificar a
    //   funcao, em vez de recriar internamente as funcoes VBA
    //   pesquisar_FC_migracao/pesquisar_FC_digitalizacao).
    //
    //   PADRAO F - INDEXACAO DE DOCUMENTO: regra especial que NAO
    //   verifica agrupamento (diferente dos Padroes A/B) - replica
    //   fielmente o comportamento do VBA original, ainda pendente de
    //   confirmacao se e regra intencional ou lacuna do codigo legado.
    //
    //   ATIVIDADES SEM CALCULO - "MATERIAL PARA ARQUIVAMENTO" e
    //   "BAIXA PERMANENTE" retornam NULL diretamente: ambas entram na
    //   Memoria de Calculo como registro informativo, mas sem nenhum
    //   valor de QExec associado (replica o "GoTo proximaLinha" do
    //   VBA original, que pulava todo o bloco de calculo para essas
    //   atividades especificas).
    //
    // ? CORRECOES APLICADAS EM RELACAO AO VBA ORIGINAL:
    //   1. Removidos branches ElseIf/Else redundantes presentes em 6
    //      Case do VBA original (calculavam o mesmo valor duas vezes).
    //   2. "ENTREGA DE ITEM NORMAL" agora arredonda a 3 casas
    //      decimais, igual a "ENTREGA DE ITEM EXPRESSO" (o VBA
    //      original tinha essa inconsistencia de precisao entre as
    //      duas atividades analogas).
    //   3. Fallback explicito com "error" para qualquer atividade nao
    //      mapeada nesta funcao - em vez de retornar 0 silenciosamente
    //      (bug real identificado em producao: "BAIXA PERMANENTE" e
    //      "MATERIAL PARA ARQUIVAMENTO" geravam esse erro ate serem
    //      tratadas explicitamente como excecoes sem calculo).
    //   4. "DEVOLUCAO DE EMPRESTIMO" com valor de Item nao reconhecido
    //      (diferente de "Embalagem"/"Item avulso") tambem lanca erro
    //      explicito, em vez de zerar silenciosamente.
    //   5. Uso de Number.Round em vez de Format(...) - evita
    //      dependencia do separador decimal do locale do sistema, e
    //      mantem o tipo numerico do inicio ao fim do calculo.
    //
    // PARAMETROS:
    //   descricaoAtividade - texto com o nome da atividade (ex.:
    //                         "COLETA DE EMBALAGEM")
    //   item                - texto com o item (ex.: "Embalagem",
    //                         "Item avulso") - usado apenas por
    //                         DEVOLUCAO DE EMPRESTIMO
    //   qtd                 - quantidade fisica solicitada/atendida
    //   agrupamento         - texto do codigo de agrupamento da linha
    //                         (vazio/null se a linha nao pertence a
    //                         um agrupamento)
    //   fc                  - fator de conversao pre-calculado (coluna
    //                         "FC" da planilha principal), usado
    //                         apenas pelas atividades de
    //                         Migracao/Digitalizacao
    //
    // RETORNO:
    //   nullable number - a quantidade padronizada de medicao (QExec),
    //   ou null para atividades sem calculo aplicavel
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
            // ? CORRECAO DEFINITIVA: qtd nula e tratada ANTES de
            // qualquer comparacao "qtd <= X" ou multiplicacao, evitando
            // que null se propague para dentro de um "if" e cause o
            // erro "nao conseguimos converter null em Logical/Number".
            // Decisao: qtd nula e tratada como 0 (equivalente a "nenhum
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
                        null   // Opcao C - sem FC cadastrado, sem interromper a consulta
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

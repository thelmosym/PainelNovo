let
    // -------------------------------------------------------------------
    // fnValidarPeriodoAnterior
    // -----------------------------------------------------------------
    // Equivalente M da função VBA validar_periodo_anterior (módulo
    // fncTabelaB).
    //
    // O QUE FAZ:
    //   Verifica se existe QUALQUER registro para um determinado
    //   período (chave "AAAA#MM") na tabela histórica de FDM/IAPFARQ
    //   (aba FDM_IAPFARQ da tabelaB, referenciada aqui como
    //   RefFDMHistorico).
    //
    //   É usada como uma validação PRÉVIA: antes de tentar buscar o
    //   FDM histórico de um contrato/item específico (função
    //   fnBuscarFDMAnterior), primeiro se confirma que aquele período
    //   já foi cadastrado na base — se não estiver, o sistema deve
    //   bloquear o recálculo e orientar o usuário a atualizar a
    //   tabelaB antes de prosseguir (mesma mensagem de erro do VBA
    //   original: "Atualize o arquivo TabelaB na rede...").
    //
    // ?? OBSERVAÇÃO IMPORTANTE (herdada da análise do VBA original):
    //   Esta validação verifica apenas se o PERÍODO existe em ALGUM
    //   registro da tabela — ela NÃO confirma que existe um registro
    //   para o CONTRATO/ITEM específico que será buscado depois. Ou
    //   seja, é uma validação parcial: um período pode "passar" nesta
    //   verificação mesmo que falte o registro exato de um contrato
    //   específico (esse caso residual é tratado com "error" explícito
    //   dentro de fnRecalcularFDM, que usa fnBuscarFDMAnterior e trata
    //   o retorno null como falha).
    //
    // PARÂMETROS:
    //   periodo - texto no formato "AAAA#MM" a ser verificado
    //
    // RETORNO:
    //   logical (true/false) — indica se o período existe na base
    //
    // DEPENDÊNCIA:
    //   RefFDMHistorico - consulta bufferizada da tabela FDM_IAPFARQ
    //                     (ver definição da consulta RefFDMHistorico)
    // -------------------------------------------------------------------
    fnValidarPeriodoAnterior = (periodo as text) as logical =>
        List.Contains(RefFDMHistorico[Medição], periodo)
in
    fnValidarPeriodoAnterior

let
    // -------------------------------------------------------------------
    // fnValidarPeriodoAnterior
    // -------------------------------------------------------------------
    // Equivalente M da funcao VBA validar_periodo_anterior (modulo
    // fncTabelaB).
    //
    // O QUE FAZ:
    //   Verifica se existe QUALQUER registro para um determinado
    //   periodo (chave "AAAA#MM") na tabela historica de FDM/IAPFARQ
    //   (aba FDM_IAPFARQ da tabelaB, referenciada aqui como
    //   RefFDMHistorico).
    //
    //   E usada como uma validacao PREVIA: antes de tentar buscar o
    //   FDM historico de um contrato/item especifico (funcao
    //   fnBuscarFDMAnterior), primeiro se confirma que aquele periodo
    //   ja foi cadastrado na base - se nao estiver, o sistema deve
    //   bloquear o recalculo e orientar o usuario a atualizar a
    //   tabelaB antes de prosseguir (mesma mensagem de erro do VBA
    //   original: "Atualize o arquivo TabelaB na rede...").
    //
    // ?? OBSERVACAO IMPORTANTE (herdada da analise do VBA original):
    //   Esta validacao verifica apenas se o PERIODO existe em ALGUM
    //   registro da tabela - ela NAO confirma que existe um registro
    //   para o CONTRATO/ITEM especifico que sera buscado depois. Ou
    //   seja, e uma validacao parcial: um periodo pode "passar" nesta
    //   verificacao mesmo que falte o registro exato de um contrato
    //   especifico (esse caso residual e tratado com "error" explicito
    //   dentro de fnRecalcularFDM, que usa fnBuscarFDMAnterior e trata
    //   o retorno null como falha).
    //
    // PARAMETROS:
    //   periodo - texto no formato "AAAA#MM" a ser verificado
    //
    // RETORNO:
    //   logical (true/false) - indica se o periodo existe na base
    //
    // DEPENDENCIA:
    //   RefFDMHistorico - consulta bufferizada da tabela FDM_IAPFARQ
    //                     (ver definicao da consulta RefFDMHistorico)
    // -------------------------------------------------------------------
    fnValidarPeriodoAnterior = (periodo as text) as logical =>
        List.Contains(RefFDMHistorico[Medição], periodo)
in
    fnValidarPeriodoAnterior

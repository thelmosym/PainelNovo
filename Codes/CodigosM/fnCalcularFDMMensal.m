let
    // -------------------------------------------------------------------
    // fnCalcularFDMMensal
    // -------------------------------------------------------------------
    // Equivalente M da funcao VBA calcular_FDM (modulo fncEditalGuarda)
    // - item no 8.4.3 do contrato de guarda externa.
    //
    // O QUE FAZ:
    //   Classifica o indice IAPFARQ em uma das 4 faixas de FDM (Fator
    //   de Desempenho Mensal), conforme regra contratual:
    //
    //       IAPFARQ < 90        ?  FDM = 0,97
    //       90 <= IAPFARQ < 95  ?  FDM = 0,98
    //       95 <= IAPFARQ < 99  ?  FDM = 0,99
    //       IAPFARQ >= 99       ?  FDM = 1,00
    //
    // ?? CORRECAO DE BUG APLICADA (em relacao a versao VBA original):
    //   O codigo VBA analisado inicialmente usava os limites
    //   "Is < 94.99" e "Is < 98.99" em vez de "< 95" e "< 99". Isso
    //   causava classificacao incorreta para valores entre 94,99-95 e
    //   98,99-99 (ex.: IAPFARQ=94,995 recebia FDM=0,99 em vez do
    //   correto 0,98). O codigo VBA foi posteriormente corrigido pelo
    //   proprio usuario (com o comentario "corrigido"), e esta funcao M
    //   ja reflete a versao CORRIGIDA (95/99), nao a versao com bug.
    //
    // PARAMETROS:
    //   IAPFARQ - o indice calculado por fnCalcularIAPFARQ
    //
    // RETORNO:
    //   number - o Fator de Desempenho Mensal (0,97 / 0,98 / 0,99 / 1)
    // -------------------------------------------------------------------
    fnCalcularFDMMensal = (IAPFARQ as number) as number =>
        if IAPFARQ < 90 then 0.97
        else if IAPFARQ < 95 then 0.98
        else if IAPFARQ < 99 then 0.99
        else 1
in
    fnCalcularFDMMensal

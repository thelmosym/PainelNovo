let
    // -------------------------------------------------------------------
    // fnCalcularFDMMensal
    // -------------------------------------------------------------------
    // Equivalente M da função VBA calcular_FDM (módulo fncEditalGuarda)
    // - item nº 8.4.3 do contrato de guarda externa Petrobras.
    //
    // O QUE FAZ:
    //   Classifica o índice IAPFARQ em uma das 4 faixas de FDM (Fator
    //   de Desempenho Mensal), conforme regra contratual:
    //
    //       IAPFARQ < 90        ->  FDM = 0,97
    //       90 <= IAPFARQ < 95  ->  FDM = 0,98
    //       95 <= IAPFARQ < 99  ->  FDM = 0,99
    //       IAPFARQ >= 99       ->  FDM = 1,00
    //
    // MELHORIAS DE ROBUSTEZ:
    //   - Aceita parâmetros nulos (retorna null defensivamente).
    //   - Estabiliza o valor de entrada com arredondamento a 2 casas
    //     decimais antes da comparação, garantindo fidelidade às faixas
    //     e eliminando distorções de ponto flutuante.
    // -------------------------------------------------------------------
    fnCalcularFDMMensal = (IAPFARQ as nullable number) as nullable number =>
        if IAPFARQ = null then
            null
        else
            let
                vIAPFARQ = Number.Round(IAPFARQ, 2)
            in
                if vIAPFARQ < 90.0 then 0.97
                else if vIAPFARQ < 95.0 then 0.98
                else if vIAPFARQ < 99.0 then 0.99
                else 1.00
in
    fnCalcularFDMMensal

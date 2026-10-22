let
    // -------------------------------------------------------------------
    // fnCalcularIAPFARQ
    // -------------------------------------------------------------------
    // Equivalente M da função VBA calcular_IAPFARQ (módulo fncEditalGuarda)
    // - item nº 8.4.2 do contrato de guarda externa Petrobras.
    //
    // O QUE FAZ:
    //   Calcula o índice IAPFARQ (percentual de solicitações atendidas
    //   DENTRO DO PRAZO em relação ao total de solicitações elegíveis
    //   atendidas no período), pela fórmula:
    //
    //       IAPFARQ = (TSPR / TSPE) * 100
    //
    // CASOS ESPECIAIS & REGRAS DE ROBUSTEZ:
    //   - Se TSPE = 0: retorna 100.0 (desempenho pleno por ausência de demanda).
    //   - Se TSPR = 0 e TSPE > 0: retorna 0.0 (conformidade nula real).
    //   - Clamping defensivo: garante que TSPR não exceda TSPE.
    //   - Tratamento seguro de nulos com condicionais nativas M.
    //   - Arredondamento padrão a 2 casas decimais via Number.Round.
    // -------------------------------------------------------------------
    fnCalcularIAPFARQ = (
        TSPE as nullable number,
        TSPR as nullable number,
        optional casasDecimais as nullable number
    ) as nullable number =>
        let
            vTSPE = if TSPE = null or TSPE < 0 then 0 else TSPE,
            vTSPR = if TSPR = null or TSPR < 0 then 0 else TSPR,
            precisao = if casasDecimais = null then 2 else casasDecimais,

            ResultadoCalculado =
                if vTSPE = 0 then
                    100.0
                else if vTSPR = 0 then
                    0.0
                else
                    let
                        vTSPRValido = if vTSPR > vTSPE then vTSPE else vTSPR,
                        Indice = (vTSPRValido / vTSPE) * 100.0
                    in
                        Indice,

            ResultadoFinal =
                if ResultadoCalculado = null then
                    null
                else
                    Number.Round(ResultadoCalculado, precisao)
        in
            ResultadoFinal
in
    fnCalcularIAPFARQ

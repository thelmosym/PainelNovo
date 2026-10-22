let
    // -------------------------------------------------------------------
    // fnRecalcularFDM
    // -------------------------------------------------------------------
    // Equivalente M da função VBA recalcular_FDM (módulo fncTabelaB)
    // - item nº 8.4.3 do contrato de guarda externa Petrobras.
    //
    // O QUE FAZ:
    //   Aplica a regra de DEGRADAÇÃO EM CASCATA sobre o FDM mensal "bruto"
    //   (calculado por fnCalcularFDMMensal), penalizando a REINCIDÊNCIA
    //   de baixo desempenho ao longo de até 3 meses consecutivos:
    //
    //     FDM > 0,97                                  -> mantém o FDM
    //     FDM = 0,97 e FDM(mês-1) >= 0,98              -> mantém 0,97
    //     FDM = 0,97 e FDM(mês-1) = 0,97 e
    //                  FDM(mês-2) > 0,97               -> degrada a 0,96
    //     FDM = 0,97 e FDM(mês-1) = 0,97 e
    //                  FDM(mês-2) <= 0,97              -> degrada a 0,95
    //     FDM = 0,97 e FDM(mês-1) = 0,97 e
    //                  FDM(mês-2) = null (só 1 anterior)-> degrada a 0,96
    //     FDM = 0,97 e FDM(mês-1) = 0,96               -> degrada a 0,95
    //     FDM = 0,97 e FDM(mês-1) = 0,95               -> mantém 0,95 (piso)
    // -------------------------------------------------------------------
    fnRecalcularFDM = (
        sPerVig as text,
        sPerAnt1 as text,
        sPerAnt2 as text,
        sContrato as text,
        dFDM as nullable number,
        iItem as number
    ) as nullable number =>
        if dFDM = null then
            null
        else if dFDM > 0.97 then
            dFDM
        else if dFDM = 0.97 then
            let
                fdmAnt1 = fnBuscarFDMAnterior(sPerAnt1, sContrato, iItem),
                fdmAnt2 = fnBuscarFDMAnterior(sPerAnt2, sContrato, iItem)
            in
                if fdmAnt1 = null then
                    error Error.Record(
                        "FDMAnteriorNaoEncontrado",
                        "FDM do período anterior (" & sPerAnt1 & ") não localizado para o contrato '" & sContrato & "', item " & Text.From(iItem) & "."
                    )
                else if fdmAnt1 >= 0.98 then
                    dFDM
                else if fdmAnt1 = 0.97 and fdmAnt2 <> null and fdmAnt2 > 0.97 then
                    0.96
                else if fdmAnt1 = 0.97 and fdmAnt2 <> null and fdmAnt2 <= 0.97 then
                    0.95
                else if fdmAnt1 = 0.97 and fdmAnt2 = null then
                    0.96
                else if fdmAnt1 = 0.96 then
                    0.95
                else if fdmAnt1 = 0.95 then
                    0.95
                else
                    error Error.Record(
                        "RegraFDMNaoTratada",
                        "Combinação de FDM anterior não prevista na regra 8.4.3: fdmAnt1=" & Text.From(fdmAnt1) & ", fdmAnt2=" & Text.From(fdmAnt2)
                    )
        else
            dFDM
in
    fnRecalcularFDM

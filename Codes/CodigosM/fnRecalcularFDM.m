let
    // -------------------------------------------------------------------
    // fnRecalcularFDM
    // -----------------------------------------------------------------
    // Equivalente M da função VBA recalcular_FDM (módulo fncTabelaB)
    // — item nº 8.4.3 do contrato de guarda externa.
    //
    // O QUE FAZ:
    //   Aplica uma regra de DEGRADAÇÃO EM CASCATA sobre o FDM mensal
    //   "bruto" (calculado por fnCalcularFDMMensal), penalizando a
    //   REINCIDÊNCIA de baixo desempenho ao longo de 3 meses
    //   consecutivos:
    //
    //     FDM > 0,97                                  ? mantém o FDM
    //     FDM = 0,97 e FDM(mês-1) >= 0,98              ? mantém 0,97
    //     FDM = 0,97 e FDM(mês-1) = 0,97 e
    //                  FDM(mês-2) > 0,97               ? degrada a 0,96
    //     FDM = 0,97 e FDM(mês-1) = 0,97 e
    //                  FDM(mês-2) = 0,97               ? degrada a 0,95
    //     FDM = 0,97 e FDM(mês-1) = 0,97 e
    //                  FDM(mês-2) < 0,97               ? degrada a 0,95
    //     FDM = 0,97 e FDM(mês-1) = 0,96               ? degrada a 0,95
    //     FDM = 0,97 e FDM(mês-1) = 0,95               ? mantém 0,95 (piso)
    //
    //   Em outras palavras: só entra na cascata quando o FDM vigente é
    //   exatamente 0,97 (a pior faixa "normal"); valores acima de 0,97
    //   nunca são penalizados. A penalização fica mais severa conforme
    //   mais meses consecutivos ficam em 0,97 ou pior.
    //
    // ? CORREÇÕES APLICADAS EM RELAÇÃO AO VBA ORIGINAL:
    //   1. Se fnBuscarFDMAnterior não encontrar o FDM do mês anterior
    //      (fdmAnt1 = null), esta função lança um "error" explícito
    //      em vez de silenciosamente deixar a variável de retorno
    //      zerada (bug identificado no VBA original: se nenhuma
    //      condição do Select Case fosse satisfeita, a função
    //      retornava 0 sem nenhum aviso, o que distorceria fortemente
    //      o valor final do FDM contratual).
    //   2. Se a combinação de fdmAnt1/fdmAnt2 não se encaixar em
    //      NENHUMA das regras previstas (situação teoricamente não
    //      esperada, mas não impossível se a base histórica tiver
    //      valores fora do padrão 0,95/0,96/0,97/0,98/0,99/1), também
    //      lança "error" explícito, permitindo diagnóstico claro em
    //      vez de mascarar a inconsistência.
    //
    // PARÂMETROS:
    //   sPerVig   - período vigente, no formato "AAAA#MM" (usado
    //               apenas para compor mensagens de erro/diagnóstico)
    //   sPerAnt1  - período do mês anterior, formato "AAAA#MM"
    //   sPerAnt2  - período de 2 meses atrás, formato "AAAA#MM"
    //   sContrato - contrato ao qual o FDM se refere
    //   dFDM      - FDM mensal "bruto" (resultado de
    //               fnCalcularFDMMensal), antes do recálculo
    //   iItem     - número do serviço/item contrato (1 ou 2)
    //
    // RETORNO:
    //   number — o FDM final, já ajustado pela regra de reincidência
    //
    // DEPENDÊNCIA:
    //   fnBuscarFDMAnterior - usada para obter o FDM histórico dos
    //                         2 meses anteriores
    // -------------------------------------------------------------------
    fnRecalcularFDM = (
        sPerVig as text,
        sPerAnt1 as text,
        sPerAnt2 as text,
        sContrato as text,
        dFDM as number,
        iItem as number
    ) as number =>
        // Caso 1: FDM acima da pior faixa (0,98/0,99/1) — nunca é
        // penalizado, mantém o valor bruto calculado
        if dFDM > 0.97 then
            dFDM

        // Caso 2: FDM = 0,97 — precisa verificar o histórico dos
        // 2 meses anteriores para decidir se há degradação
        else if dFDM = 0.97 then
            let
                fdmAnt1 = fnBuscarFDMAnterior(sPerAnt1, sContrato, iItem),
                fdmAnt2 = fnBuscarFDMAnterior(sPerAnt2, sContrato, iItem)
            in
                // Se o mês anterior não tiver registro histórico,
                // não há como aplicar a regra de reincidência —
                // interrompe com erro explícito em vez de assumir 0
                if fdmAnt1 = null then
                    error Error.Record(
                        "FDMAnteriorNaoEncontrado",
                        "FDM do período anterior (" & sPerAnt1 & ") não localizado para o contrato '" & sContrato & "', item " & Text.From(iItem) & "."
                    )

                // Mês anterior teve bom desempenho (>=0,98) — sem
                // reincidência, mantém o FDM vigente em 0,97
                else if fdmAnt1 >= 0.98 then
                    dFDM

                // 2º mês consecutivo em 0,97, mas o mês retrasado
                // (2 meses atrás) foi melhor que 0,97 — 1ª penalização
                else if fdmAnt1 = 0.97 and fdmAnt2 <> null and fdmAnt2 > 0.97 then
                    0.96

                // 3 meses consecutivos exatamente em 0,97 —
                // penalização mais severa (piso mínimo)
                else if fdmAnt1 = 0.97 and fdmAnt2 <> null and fdmAnt2 = 0.97 then
                    0.95

                // 2º mês em 0,97, e o mês retrasado já estava pior
                // que 0,97 — mesma penalização severa (piso mínimo)
                else if fdmAnt1 = 0.97 and fdmAnt2 <> null and fdmAnt2 < 0.97 then
                    0.95

                // Mês anterior já estava em 0,96 (penalização em
                // andamento) — degrada mais um nível, para 0,95
                else if fdmAnt1 = 0.96 then
                    0.95

                // Mês anterior já estava no piso (0,95) — permanece
                // no piso, não degrada mais
                else if fdmAnt1 = 0.95 then
                    0.95

                // Combinação de valores históricos fora do padrão
                // esperado — sinaliza inconsistência em vez de
                // mascarar com um resultado arbitrário
                else
                    error Error.Record(
                        "RegraFDMNaoTratada",
                        "Combinação de FDM anterior não prevista na regra 8.4.3: fdmAnt1=" & Text.From(fdmAnt1) & ", fdmAnt2=" & Text.From(fdmAnt2)
                    )

        // Caso teórico não alcançável na prática (fnCalcularFDMMensal
        // nunca retorna valor menor que 0,97) — mantido apenas por
        // segurança/fidelidade estrutural ao Select Case original
        else
            dFDM
in
    fnRecalcularFDM

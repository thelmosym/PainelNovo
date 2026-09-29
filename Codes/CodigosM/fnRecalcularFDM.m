let
    // -------------------------------------------------------------------
    // fnRecalcularFDM
    // -------------------------------------------------------------------
    // Equivalente M da funcao VBA recalcular_FDM (modulo fncTabelaB)
    // - item no 8.4.3 do contrato de guarda externa.
    //
    // O QUE FAZ:
    //   Aplica uma regra de DEGRADACAO EM CASCATA sobre o FDM mensal
    //   "bruto" (calculado por fnCalcularFDMMensal), penalizando a
    //   REINCIDENCIA de baixo desempenho ao longo de 3 meses
    //   consecutivos:
    //
    //     FDM > 0,97                                  ? mantem o FDM
    //     FDM = 0,97 e FDM(mes-1) >= 0,98              ? mantem 0,97
    //     FDM = 0,97 e FDM(mes-1) = 0,97 e
    //                  FDM(mes-2) > 0,97               ? degrada a 0,96
    //     FDM = 0,97 e FDM(mes-1) = 0,97 e
    //                  FDM(mes-2) = 0,97               ? degrada a 0,95
    //     FDM = 0,97 e FDM(mes-1) = 0,97 e
    //                  FDM(mes-2) < 0,97               ? degrada a 0,95
    //     FDM = 0,97 e FDM(mes-1) = 0,96               ? degrada a 0,95
    //     FDM = 0,97 e FDM(mes-1) = 0,95               ? mantem 0,95 (piso)
    //
    //   Em outras palavras: so entra na cascata quando o FDM vigente e
    //   exatamente 0,97 (a pior faixa "normal"); valores acima de 0,97
    //   nunca sao penalizados. A penalizacao fica mais severa conforme
    //   mais meses consecutivos ficam em 0,97 ou pior.
    //
    // ? CORRECOES APLICADAS EM RELACAO AO VBA ORIGINAL:
    //   1. Se fnBuscarFDMAnterior nao encontrar o FDM do mes anterior
    //      (fdmAnt1 = null), esta funcao lanca um "error" explicito
    //      em vez de silenciosamente deixar a variavel de retorno
    //      zerada (bug identificado no VBA original: se nenhuma
    //      condicao do Select Case fosse satisfeita, a funcao
    //      retornava 0 sem nenhum aviso, o que distorceria fortemente
    //      o valor final do FDM contratual).
    //   2. Se a combinacao de fdmAnt1/fdmAnt2 nao se encaixar em
    //      NENHUMA das regras previstas (situacao teoricamente nao
    //      esperada, mas nao impossivel se a base historica tiver
    //      valores fora do padrao 0,95/0,96/0,97/0,98/0,99/1), tambem
    //      lanca "error" explicito, permitindo diagnostico claro em
    //      vez de mascarar a inconsistencia.
    //
    // PARAMETROS:
    //   sPerVig   - periodo vigente, no formato "AAAA#MM" (usado
    //               apenas para compor mensagens de erro/diagnostico)
    //   sPerAnt1  - periodo do mes anterior, formato "AAAA#MM"
    //   sPerAnt2  - periodo de 2 meses atras, formato "AAAA#MM"
    //   sContrato - contrato ao qual o FDM se refere
    //   dFDM      - FDM mensal "bruto" (resultado de
    //               fnCalcularFDMMensal), antes do recalculo
    //   iItem     - numero do servico/item contrato (1 ou 2)
    //
    // RETORNO:
    //   number - o FDM final, ja ajustado pela regra de reincidencia
    //
    // DEPENDENCIA:
    //   fnBuscarFDMAnterior - usada para obter o FDM historico dos
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
        // Caso 1: FDM acima da pior faixa (0,98/0,99/1) - nunca e
        // penalizado, mantem o valor bruto calculado
        if dFDM > 0.97 then
            dFDM
        // Caso 2: FDM = 0,97 - precisa verificar o historico dos
        // 2 meses anteriores para decidir se ha degradacao
        else if dFDM = 0.97 then
            let
                fdmAnt1 = fnBuscarFDMAnterior(sPerAnt1, sContrato, iItem),
                fdmAnt2 = fnBuscarFDMAnterior(sPerAnt2, sContrato, iItem)
            in
                // Se o mes anterior nao tiver registro historico,
                // nao ha como aplicar a regra de reincidencia -
                // interrompe com erro explicito em vez de assumir 0
                if fdmAnt1 = null then
                    error Error.Record(
                        "FDMAnteriorNaoEncontrado",
                        "FDM do período anterior (" & sPerAnt1 & ") não localizado para o contrato '" & sContrato & "', item " & Text.From(iItem) & "."
                    )
                // Mes anterior teve bom desempenho (>=0,98) - sem
                // reincidencia, mantem o FDM vigente em 0,97
                else if fdmAnt1 >= 0.98 then
                    dFDM
                // 2o mes consecutivo em 0,97, mas o mes retrasado
                // (2 meses atras) foi melhor que 0,97 - 1a penalizacao
                else if fdmAnt1 = 0.97 and fdmAnt2 <> null and fdmAnt2 > 0.97 then
                    0.96
                // 3 meses consecutivos exatamente em 0,97 -
                // penalizacao mais severa (piso minimo)
                else if fdmAnt1 = 0.97 and fdmAnt2 <> null and fdmAnt2 = 0.97 then
                    0.95
                // 2o mes em 0,97, e o mes retrasado ja estava pior
                // que 0,97 - mesma penalizacao severa (piso minimo)
                else if fdmAnt1 = 0.97 and fdmAnt2 <> null and fdmAnt2 < 0.97 then
                    0.95
                // Mes anterior ja estava em 0,96 (penalizacao em
                // andamento) - degrada mais um nivel, para 0,95
                else if fdmAnt1 = 0.96 then
                    0.95
                // Mes anterior ja estava no piso (0,95) - permanece
                // no piso, nao degrada mais
                else if fdmAnt1 = 0.95 then
                    0.95
                // Combinacao de valores historicos fora do padrao
                // esperado - sinaliza inconsistencia em vez de
                // mascarar com um resultado arbitrario
                else
                    error Error.Record(
                        "RegraFDMNaoTratada",
                        "Combinação de FDM anterior não prevista na regra 8.4.3: fdmAnt1=" & Text.From(fdmAnt1) & ", fdmAnt2=" & Text.From(fdmAnt2)
                    )
        // Caso teorico nao alcancavel na pratica (fnCalcularFDMMensal
        // nunca retorna valor menor que 0,97) - mantido apenas por
        // seguranca/fidelidade estrutural ao Select Case original
        else
            dFDM
in
    fnRecalcularFDM

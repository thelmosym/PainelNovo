let
    // -------------------------------------------------------------------
    // fnCalcularTSPE_TSPR
    // -------------------------------------------------------------------
    // Equivalente M da função VBA calcular_TSPE_TSPR (módulo fncEditalGuarda)
    // - item nº 8.4.2 do contrato de guarda externa Petrobras.
    //
    // O QUE FAZ:
    //   Calcula, para um contrato específico, os quantitativos de
    //   solicitações atendidas (TSPE) e atendidas no prazo (TSPR),
    //   segregados por linha de serviço (Serviço 1 e Serviço 2,
    //   determinados pela coluna "Ref. FDM"):
    //     - TSPE1 / TSPE2 = Total de solicitações elegíveis no período
    //       (linhas com Situação = "CO" e Status Prazo = "NP" ou "FP")
    //     - TSPR1 / TSPR2 = Total de solicitações atendidas no prazo
    //       (linhas com Situação = "CO" e Status Prazo = "NP")
    // -------------------------------------------------------------------
    fnCalcularTSPE_TSPR = (tabela as table, contratoA as nullable text) as record =>
        let
            vContrato = if contratoA = null then "" else contratoA,

            // Filtra o contrato informado e apenas atendimentos concluídos ("CO")
            Filtrado = Table.SelectRows(tabela, each
                [Contrato] = vContrato and [Situação] = "CO"
            ),

            // Agrupa por Serviço ("Ref. FDM") com agregação vetorizada via lista
            Agrupado = Table.Group(Filtrado, {"Ref. FDM"}, {
                {"TSPE", each
                    let st = [Status Prazo]
                    in List.Count(List.Select(st, each _ = "NP" or _ = "FP")),
                    Int64.Type
                },
                {"TSPR", each
                    let st = [Status Prazo]
                    in List.Count(List.Select(st, each _ = "NP")),
                    Int64.Type
                }
            }),

            Servico1 = Table.SelectRows(Agrupado, each [#"Ref. FDM"] = 1),
            Servico2 = Table.SelectRows(Agrupado, each [#"Ref. FDM"] = 2),

            rTSPE1 = if not Table.IsEmpty(Servico1) then Servico1{0}[TSPE] else 0,
            rTSPR1 = if not Table.IsEmpty(Servico1) then Servico1{0}[TSPR] else 0,
            rTSPE2 = if not Table.IsEmpty(Servico2) then Servico2{0}[TSPE] else 0,
            rTSPR2 = if not Table.IsEmpty(Servico2) then Servico2{0}[TSPR] else 0
        in
            [ TSPE1 = rTSPE1, TSPR1 = rTSPR1, TSPE2 = rTSPE2, TSPR2 = rTSPR2 ]
in
    fnCalcularTSPE_TSPR

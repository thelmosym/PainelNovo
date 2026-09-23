let
    // -------------------------------------------------------------------
    // fnCalcularTSPE_TSPR
    // -----------------------------------------------------------------
    // Equivalente M da função VBA calcular_TSPE_TSPR (módulo
    // fncEditalGuarda) — item nº 8.4.2 do contrato de guarda externa.
    //
    // O QUE FAZ:
    //   Calcula, para um contrato específico, 4 indicadores separados
    //   por serviço da PPU (Serviço 1 e Serviço 2 — a separação é
    //   definida pela coluna "Ref. FDM", já mesclada previamente na
    //   tabela principal a partir da tabelaA):
    //     - TSPE1 / TSPE2 = Total de Solicitações atendidas no Período
    //       (conta linhas com Status Prazo = "NP" OU "FP")
    //     - TSPR1 / TSPR2 = Total de Solicitações atendidas dentro do
    //       Prazo (conta apenas linhas com Status Prazo = "NP")
    //
    // FILTROS APLICADOS (equivalentes ao VBA original):
    //   - Contrato da linha = contrato informado no parâmetro
    //   - Situação da linha = "CO" (concluído)
    //
    // DIFERENÇA EM RELAÇÃO AO VBA:
    //   O VBA percorre linha a linha com um loop manual (For i = 3 To
    //   iULin) e um Select Case por linha. Em M, a mesma lógica é
    //   resolvida de forma declarativa com Table.Group, agrupando por
    //   "Ref. FDM" e contando as linhas de cada subconjunto — muito
    //   mais performático que reprocessar linha a linha.
    //
    // PARÂMETROS:
    //   tabela    - tabela principal já classificada, contendo as
    //               colunas "Contrato", "Situação", "Status Prazo" e
    //               "Ref. FDM"
    //   contratoA - texto com o contrato para o qual calcular os
    //               índices (ex.: "IRON-LT1")
    //
    // RETORNO:
    //   Um record com 4 campos: [TSPE1, TSPR1, TSPE2, TSPR2]
    //   Se não houver nenhuma linha para um dado serviço (1 ou 2), o
    //   valor correspondente é 0 (equivalente ao comportamento padrão
    //   de variáveis Integer não inicializadas no VBA original).
    // -------------------------------------------------------------------
    fnCalcularTSPE_TSPR = (tabela as table, contratoA as text) as record =>
        let
            // Filtra apenas as linhas do contrato desejado com
            // Situação = "CO" (concluído) — equivalente aos dois
            // "If" iniciais do loop VBA (sContratoA = sContratoN e
            // sSituacao = "CO")
            Filtrado = Table.SelectRows(tabela, each
                [Contrato] = contratoA and [Situação] = "CO"
            ),

            // Agrupa as linhas filtradas por "Ref. FDM" (1 ou 2) e,
            // dentro de cada grupo, conta quantas linhas têm Status
            // Prazo = NP/FP (TSPE) e quantas têm Status Prazo = NP (TSPR)
            Agrupado = Table.Group(Filtrado, {"Ref. FDM"}, {
                {"TSPE", each Table.RowCount(
                    Table.SelectRows(_, each [Status Prazo] = "NP" or [Status Prazo] = "FP")
                ), Int64.Type},
                {"TSPR", each Table.RowCount(
                    Table.SelectRows(_, each [Status Prazo] = "NP")
                ), Int64.Type}
            }),

            // Extrai a linha do grupo correspondente ao Serviço 1
            // (Ref. FDM = 1) e ao Serviço 2 (Ref. FDM = 2)
            Servico1 = Table.SelectRows(Agrupado, each [#"Ref. FDM"] = 1),
            Servico2 = Table.SelectRows(Agrupado, each [#"Ref. FDM"] = 2),

            // Se o grupo não existir (nenhuma linha com aquela Ref.
            // FDM), assume 0 como valor padrão — evita erro de índice
            // fora do intervalo ao acessar {0}
            TSPE1 = if Table.RowCount(Servico1) > 0 then Servico1{0}[TSPE] else 0,
            TSPR1 = if Table.RowCount(Servico1) > 0 then Servico1{0}[TSPR] else 0,
            TSPE2 = if Table.RowCount(Servico2) > 0 then Servico2{0}[TSPE] else 0,
            TSPR2 = if Table.RowCount(Servico2) > 0 then Servico2{0}[TSPR] else 0
        in
            [ TSPE1 = TSPE1, TSPR1 = TSPR1, TSPE2 = TSPE2, TSPR2 = TSPR2 ]
in
    fnCalcularTSPE_TSPR

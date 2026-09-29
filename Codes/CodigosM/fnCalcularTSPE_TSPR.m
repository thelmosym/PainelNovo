let
    // -------------------------------------------------------------------
    // fnCalcularTSPE_TSPR
    // -------------------------------------------------------------------
    // Equivalente M da funcao VBA calcular_TSPE_TSPR (modulo
    // fncEditalGuarda) - item no 8.4.2 do contrato de guarda externa.
    //
    // O QUE FAZ:
    //   Calcula, para um contrato especifico, 4 indicadores separados
    //   por servico da PPU (Servico 1 e Servico 2 - a separacao e
    //   definida pela coluna "Ref. FDM", ja mesclada previamente na
    //   tabela principal a partir da tabelaA):
    //     - TSPE1 / TSPE2 = Total de Solicitacoes atendidas no Periodo
    //       (conta linhas com Status Prazo = "NP" OU "FP")
    //     - TSPR1 / TSPR2 = Total de Solicitacoes atendidas dentro do
    //       Prazo (conta apenas linhas com Status Prazo = "NP")
    //
    // FILTROS APLICADOS (equivalentes ao VBA original):
    //   - Contrato da linha = contrato informado no parametro
    //   - Situacao da linha = "CO" (concluido)
    //
    // DIFERENCA EM RELACAO AO VBA:
    //   O VBA percorre linha a linha com um loop manual (For i = 3 To
    //   iULin) e um Select Case por linha. Em M, a mesma logica e
    //   resolvida de forma declarativa com Table.Group, agrupando por
    //   "Ref. FDM" e contando as linhas de cada subconjunto - muito
    //   mais performatico que reprocessar linha a linha.
    //
    // PARAMETROS:
    //   tabela    - tabela principal ja classificada, contendo as
    //               colunas "Contrato", "Situacao", "Status Prazo" e
    //               "Ref. FDM"
    //   contratoA - texto com o contrato para o qual calcular os
    //               indices (ex.: "IRON-LT1")
    //
    // RETORNO:
    //   Um record com 4 campos: [TSPE1, TSPR1, TSPE2, TSPR2]
    //   Se nao houver nenhuma linha para um dado servico (1 ou 2), o
    //   valor correspondente e 0 (equivalente ao comportamento padrao
    //   de variaveis Integer nao inicializadas no VBA original).
    // -------------------------------------------------------------------
    fnCalcularTSPE_TSPR = (tabela as table, contratoA as text) as record =>
        let
            // Filtra apenas as linhas do contrato desejado com
            // Situacao = "CO" (concluido) - equivalente aos dois
            // "If" iniciais do loop VBA (sContratoA = sContratoN e
            // sSituacao = "CO")
            Filtrado = Table.SelectRows(tabela, each
                [Contrato] = contratoA and [Situação] = "CO"
            ),
            // Agrupa as linhas filtradas por "Ref. FDM" (1 ou 2) e,
            // dentro de cada grupo, conta quantas linhas tem Status
            // Prazo = NP/FP (TSPE) e quantas tem Status Prazo = NP (TSPR)
            Agrupado = Table.Group(Filtrado, {"Ref. FDM"}, {
                {"TSPE", each Table.RowCount(
                    Table.SelectRows(_, each [Status Prazo] = "NP" or [Status Prazo] = "FP")
                ), Int64.Type},
                {"TSPR", each Table.RowCount(
                    Table.SelectRows(_, each [Status Prazo] = "NP")
                ), Int64.Type}
            }),
            // Extrai a linha do grupo correspondente ao Servico 1
            // (Ref. FDM = 1) e ao Servico 2 (Ref. FDM = 2)
            Servico1 = Table.SelectRows(Agrupado, each [#"Ref. FDM"] = 1),
            Servico2 = Table.SelectRows(Agrupado, each [#"Ref. FDM"] = 2),
            // Se o grupo nao existir (nenhuma linha com aquela Ref.
            // FDM), assume 0 como valor padrao - evita erro de indice
            // fora do intervalo ao acessar {0}
            TSPE1 = if Table.RowCount(Servico1) > 0 then Servico1{0}[TSPE] else 0,
            TSPR1 = if Table.RowCount(Servico1) > 0 then Servico1{0}[TSPR] else 0,
            TSPE2 = if Table.RowCount(Servico2) > 0 then Servico2{0}[TSPE] else 0,
            TSPR2 = if Table.RowCount(Servico2) > 0 then Servico2{0}[TSPR] else 0
        in
            [ TSPE1 = TSPE1, TSPR1 = TSPR1, TSPE2 = TSPE2, TSPR2 = TSPR2 ]
in
    fnCalcularTSPE_TSPR

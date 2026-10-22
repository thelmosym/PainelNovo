let
    // -------------------------------------------------------------------
    // fnBuscarFDMAnterior
    // -------------------------------------------------------------------
    // Equivalente M da função VBA validar_FDMAnterior (módulo fncTabelaB).
    // -------------------------------------------------------------------
    fnBuscarFDMAnterior = (periodo as text, contrato as text, item as number) as nullable number =>
        let
            Filtrado = Table.SelectRows(RefFDMHistorico, each
                [Medição] = periodo and [Contrato] = contrato and [#"Item contrato"] = item
            ),
            Resultado = if not Table.IsEmpty(Filtrado) then Filtrado{0}[FDM] else null
        in
            Resultado
in
    fnBuscarFDMAnterior

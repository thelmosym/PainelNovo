let
    // ? Referencia a consulta principal (nome confirmado: "Painel T2M")
    BaseParaCalculo = #"Painel T2M",
    // 1) Lista de contratos distintos presentes na base
    ContratosDistintos = Table.Distinct(Table.SelectColumns(BaseParaCalculo, {"Contrato"})),
    // 2) Para cada contrato, calcula TSPE1/TSPR1/TSPE2/TSPR2
    ComTSPE_TSPR = Table.AddColumn(ContratosDistintos, "_TSPE_TSPR", each
        fnCalcularTSPE_TSPR(BaseParaCalculo, [Contrato])
    ),
    ExpandirTSPE_TSPR = Table.ExpandRecordColumn(ComTSPE_TSPR, "_TSPE_TSPR",
        {"TSPE1", "TSPR1", "TSPE2", "TSPR2"},
        {"TSPE1", "TSPR1", "TSPE2", "TSPR2"}
    ),
    // 3) Calcula IAPFARQ1 e IAPFARQ2
    ComIAPFARQ1 = Table.AddColumn(ExpandirTSPE_TSPR, "IAPFARQ1", each
        fnCalcularIAPFARQ([TSPE1], [TSPR1])
    ),
    ComIAPFARQ2 = Table.AddColumn(ComIAPFARQ1, "IAPFARQ2", each
        fnCalcularIAPFARQ([TSPE2], [TSPR2])
    ),
    // 4) FDM mensal "bruto" (antes do recalculo com periodos anteriores)
    ComFDM1Bruto = Table.AddColumn(ComIAPFARQ2, "FDM1_Bruto", each
        fnCalcularFDMMensal([IAPFARQ1])
    ),
    ComFDM2Bruto = Table.AddColumn(ComFDM1Bruto, "FDM2_Bruto", each
        fnCalcularFDMMensal([IAPFARQ2])
    ),
    // 5) Calcula os periodos anteriores.
    //    ? CORRECAO APLICADA: PeriodoMes e um parametro em TEXTO
    //    (ex.: "Julho"), entao precisa ser convertido para numero
    //    antes de ser passado para fnCalcularPeriodosAnteriores
    //    (que espera "mes as number").
    ComPeriodos = Table.AddColumn(ComFDM2Bruto, "_Periodos", each
        fnCalcularPeriodosAnteriores(fnConverterMesTextoParaNumero(PeriodoMes), PeriodoAno)
    ),
    ExpandirPeriodos = Table.ExpandRecordColumn(ComPeriodos, "_Periodos",
        {"PerAnt1", "PerAnt2"}, {"PerAnt1", "PerAnt2"}
    ),
    // 6) Valida se os periodos anteriores existem na base historica
    //    (equivalente a validar_periodo_anterior)
    ComValidacaoPeriodos = Table.AddColumn(ExpandirPeriodos, "PeriodosValidos", each
        fnValidarPeriodoAnterior([PerAnt1]) and fnValidarPeriodoAnterior([PerAnt2])
    ),
    // 7) Periodo vigente no formato "AAAA#MM".
    //    ? CORRECAO APLICADA: mesma conversao de texto?numero aqui,
    //    antes de formatar com Text.PadStart.
    ComPeriodoVigente = Table.AddColumn(ComValidacaoPeriodos, "PeriodoVigente", each
        Text.From(PeriodoAno) & "#" & Text.PadStart(Text.From(fnConverterMesTextoParaNumero(PeriodoMes)), 2, "0")
    ),
    // 8) Recalculo final do FDM (item 8.4.3), com diagnostico em caso
    //    de erro (mesmo padrao defensivo ja usado em fnClassificarFDM)
    ComFDM1Final = Table.AddColumn(ComPeriodoVigente, "FDM1", each
        if not [PeriodosValidos] then
            null
        else
            try fnRecalcularFDM([PeriodoVigente], [PerAnt1], [PerAnt2], [Contrato], [FDM1_Bruto], 1)
            otherwise null
    ),
    ComFDM2Final = Table.AddColumn(ComFDM1Final, "FDM2", each
        if not [PeriodosValidos] then
            null
        else
            try fnRecalcularFDM([PeriodoVigente], [PerAnt1], [PerAnt2], [Contrato], [FDM2_Bruto], 2)
            otherwise null
    ),
    // 9) Selecao final das colunas - a tabela por contrato pedida
    TabelaFinal = Table.SelectColumns(ComFDM2Final, {
        "Contrato",
        "TSPE1", "TSPR1", "IAPFARQ1", "FDM1_Bruto", "FDM1",
        "TSPE2", "TSPR2", "IAPFARQ2", "FDM2_Bruto", "FDM2",
        "PeriodosValidos"
    })
in
    TabelaFinal

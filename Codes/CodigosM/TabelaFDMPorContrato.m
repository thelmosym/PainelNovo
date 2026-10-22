let
    // -------------------------------------------------------------------
    // TabelaFDMPorContrato
    // -------------------------------------------------------------------
    // Consolida os indicadores de conformidade de SLA e FDM por contrato.
    //
    // OTIMIZAÇÕES & MELHORIAS APLICADAS:
    //   1. Projeção antecipada de colunas: seleciona apenas as 4 colunas
    //      necessárias de 'Painel T2M', reduzindo drasticamente o consumo
    //      de memória RAM durante as agregações de TSPE/TSPR.
    //   2. Pré-cálculo de parâmetros estáticos fora do pipeline linha a linha:
    //      mês numérico, períodos anteriores e validação histórica são
    //      computados uma única vez no bloco 'let'.
    //   3. Fallback seguro de FDM: caso os períodos anteriores não estejam
    //      disponíveis na base histórica, o FDM final assume o FDM Bruto
    //      calculado para o próprio mês (em vez de resultar em null),
    //      preservando o comportamento fiel do VBA e a continuidade da fatura.
    // -------------------------------------------------------------------

    // 1) Constantes e parâmetros estáticos calculados uma única vez
    MesNumero = fnConverterMesTextoParaNumero(PeriodoMes),
    PeriodoVigenteConst = Text.From(PeriodoAno) & "#" & Text.PadStart(Text.From(MesNumero), 2, "0"),
    PeriodosAnterioresRec = fnCalcularPeriodosAnteriores(MesNumero, PeriodoAno),
    PerAnt1 = PeriodosAnterioresRec[PerAnt1],
    PerAnt2 = PeriodosAnterioresRec[PerAnt2],
    PeriodosValidosConst = fnValidarPeriodoAnterior(PerAnt1) and fnValidarPeriodoAnterior(PerAnt2),

    // 2) Projeção otimizada da base analítica (somente colunas essenciais para FDM)
    BaseParaCalculo = Table.SelectColumns(
        #"Painel T2M",
        {"Contrato", "Situação", "Ref. FDM", "Status Prazo"}
    ),

    // 3) Lista de contratos distintos presentes na base
    ContratosDistintos = Table.Distinct(Table.SelectColumns(BaseParaCalculo, {"Contrato"})),

    // 4) Para cada contrato, calcula TSPE1/TSPR1/TSPE2/TSPR2
    ComTSPE_TSPR = Table.AddColumn(ContratosDistintos, "_TSPE_TSPR", each
        fnCalcularTSPE_TSPR(BaseParaCalculo, [Contrato]),
        type record
    ),
    ExpandirTSPE_TSPR = Table.ExpandRecordColumn(
        ComTSPE_TSPR, "_TSPE_TSPR",
        {"TSPE1", "TSPR1", "TSPE2", "TSPR2"},
        {"TSPE1", "TSPR1", "TSPE2", "TSPR2"}
    ),

    // 5) Calcula os índices IAPFARQ1 e IAPFARQ2 com precisão decimal
    ComIAPFARQ1 = Table.AddColumn(ExpandirTSPE_TSPR, "IAPFARQ1", each
        fnCalcularIAPFARQ([TSPE1], [TSPR1]),
        type number
    ),
    ComIAPFARQ2 = Table.AddColumn(ComIAPFARQ1, "IAPFARQ2", each
        fnCalcularIAPFARQ([TSPE2], [TSPR2]),
        type number
    ),

    // 6) FDM mensal "bruto" (classificação direta do IAPFARQ)
    ComFDM1Bruto = Table.AddColumn(ComIAPFARQ2, "FDM1_Bruto", each
        fnCalcularFDMMensal([IAPFARQ1]),
        type number
    ),
    ComFDM2Bruto = Table.AddColumn(ComFDM1Bruto, "FDM2_Bruto", each
        fnCalcularFDMMensal([IAPFARQ2]),
        type number
    ),

    // 7) Agrega o flag de auditoria de períodos válidos
    ComValidacaoPeriodos = Table.AddColumn(ComFDM2Bruto, "PeriodosValidos", each
        PeriodosValidosConst,
        type logical
    ),

    // 8) Recálculo final do FDM com degradação em cascata (item 8.4.3)
    //    Se não houver histórico válido, assume o FDM Bruto mensal como fallback
    ComFDM1Final = Table.AddColumn(ComValidacaoPeriodos, "FDM1", each
        if not PeriodosValidosConst then
            [FDM1_Bruto]
        else
            try fnRecalcularFDM(PeriodoVigenteConst, PerAnt1, PerAnt2, [Contrato], [FDM1_Bruto], 1)
            otherwise [FDM1_Bruto],
        type number
    ),
    ComFDM2Final = Table.AddColumn(ComFDM1Final, "FDM2", each
        if not PeriodosValidosConst then
            [FDM2_Bruto]
        else
            try fnRecalcularFDM(PeriodoVigenteConst, PerAnt1, PerAnt2, [Contrato], [FDM2_Bruto], 2)
            otherwise [FDM2_Bruto],
        type number
    ),

    // 9) Seleção final e ordenação das colunas
    TabelaFinal = Table.SelectColumns(ComFDM2Final, {
        "Contrato",
        "TSPE1", "TSPR1", "IAPFARQ1", "FDM1_Bruto", "FDM1",
        "TSPE2", "TSPR2", "IAPFARQ2", "FDM2_Bruto", "FDM2",
        "PeriodosValidos"
    })
in
    TabelaFinal

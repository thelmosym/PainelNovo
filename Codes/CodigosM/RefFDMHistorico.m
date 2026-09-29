let
    // -------------------------------------------------------------------
    // RefFDMHistorico
    // -------------------------------------------------------------------
    // Consulta de referencia (nao e uma funcao - e uma tabela) que
    // carrega e bufferiza o historico de FDM/IAPFARQ por periodo,
    // contrato e item, a partir da aba "FDM_IAPFARQ" da tabelaB.
    //
    // O QUE FAZ:
    //   1. Abre o arquivo fisico da tabelaB (localTabelaB) e extrai a
    //      tabela nomeada "FDM_IAPFARQ".
    //   2. Define explicitamente o tipo de cada coluna:
    //        - "Medicao"        ? texto (chave de periodo "AAAA#MM")
    //        - "Contrato"       ? texto
    //        - "Item contrato"  ? numero inteiro (1 ou 2)
    //        - "IAPFARQ (%)"    ? numero decimal
    //        - "FDM"            ? numero decimal
    //   3. Aplica Table.Buffer - carrega o resultado em memoria UMA
    //      UNICA VEZ, evitando que o arquivo fisico seja reaberto a
    //      cada chamada de fnValidarPeriodoAnterior ou
    //      fnBuscarFDMAnterior (que sao invocadas repetidamente, uma
    //      vez por contrato, dentro de TabelaFDMPorContrato).
    //
    // USADA POR:
    //   - fnValidarPeriodoAnterior (consulta RefFDMHistorico[Medicao])
    //   - fnBuscarFDMAnterior (Table.SelectRows sobre RefFDMHistorico)
    //
    // ?? IMPORTANTE:
    //   Esta consulta deve ser referenciada pelo nome exato
    //   "RefFDMHistorico" dentro das duas funcoes acima. Se voce
    //   renomear esta consulta, ajuste tambem as referencias internas
    //   de fnValidarPeriodoAnterior e fnBuscarFDMAnterior.
    // -------------------------------------------------------------------
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    FDM_IAPFARQ_Table = Fonte{[Item = "FDM_IAPFARQ", Kind = "Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(FDM_IAPFARQ_Table, {
        {"Medição", type text},
        {"Contrato", type text},
        {"Item contrato", Int64.Type},
        {"IAPFARQ (%)", type number},
        {"FDM", type number}
    }),
    // Bufferiza o resultado final - evita reabrir o arquivo fisico
    // a cada linha/contrato processado nas funcoes que dependem
    // desta consulta
    RefFDMHistorico = Table.Buffer(#"Tipo Alterado")
in
    RefFDMHistorico

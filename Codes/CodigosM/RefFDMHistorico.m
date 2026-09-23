let
    // -------------------------------------------------------------------
    // RefFDMHistorico
    // -----------------------------------------------------------------
    // Consulta de referência (não é uma função — é uma tabela) que
    // carrega e bufferiza o histórico de FDM/IAPFARQ por período,
    // contrato e item, a partir da aba "FDM_IAPFARQ" da tabelaB.
    //
    // O QUE FAZ:
    //   1. Abre o arquivo físico da tabelaB (localTabelaB) e extrai a
    //      tabela nomeada "FDM_IAPFARQ".
    //   2. Define explicitamente o tipo de cada coluna:
    //        - "Medição"        ? texto (chave de período "AAAA#MM")
    //        - "Contrato"       ? texto
    //        - "Item contrato"  ? número inteiro (1 ou 2)
    //        - "IAPFARQ (%)"    ? número decimal
    //        - "FDM"            ? número decimal
    //   3. Aplica Table.Buffer — carrega o resultado em memória UMA
    //      ÚNICA VEZ, evitando que o arquivo físico seja reaberto a
    //      cada chamada de fnValidarPeriodoAnterior ou
    //      fnBuscarFDMAnterior (que são invocadas repetidamente, uma
    //      vez por contrato, dentro de TabelaFDMPorContrato).
    //
    // USADA POR:
    //   - fnValidarPeriodoAnterior (consulta RefFDMHistorico[Medição])
    //   - fnBuscarFDMAnterior (Table.SelectRows sobre RefFDMHistorico)
    //
    // ?? IMPORTANTE:
    //   Esta consulta deve ser referenciada pelo nome exato
    //   "RefFDMHistorico" dentro das duas funções acima. Se você
    //   renomear esta consulta, ajuste também as referências internas
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

    // Bufferiza o resultado final — evita reabrir o arquivo físico
    // a cada linha/contrato processado nas funções que dependem
    // desta consulta
    RefFDMHistorico = Table.Buffer(#"Tipo Alterado")
in
    RefFDMHistorico

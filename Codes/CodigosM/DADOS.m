let
    Fonte = #"Painel T2M",

    // Table.SelectColumns seleciona E reordena em uma única operação —
    // substitui os 3 passos de Table.ReorderColumns + os 2 passos de
    // Table.RemoveColumns do código original. "Obs. Sigla" já é
    // excluída aqui diretamente — no original ela era renomeada para
    // "Obs." e só depois removida, desperdiçando processamento.
    ColunasSelecionadas = Table.SelectColumns(Fonte, {
        "Contrato.1",
        "Descrição da atividade", "Item", "Aplicação",
        "Código da#(lf)solicitação", "Data solicitação", "Gerência solicitante",
        "Qtd.#(lf)Solicitada", "Código OS", "D. abertura", "D. fechamento",
        "Prazo combinado", "Qtd.#(lf)Atendida",
        "Classificação FDM", "Status Prazo",
        "Centro", "Município", "UF",
        "Galpão", "Galpãp_Cidade", "Galpãp_UF",
        "KM Adicional", "FC", "QExecAgrupado","QExec",
        "Unidade", "Localidade", "Código Centro", "Linha de serviço PPU"
    }),

    // As 10 renomeações combinadas em UMA única chamada (o original
    // fazia isso em 6 chamadas separadas de RenameColumns, cada uma
    // intercalada com uma reordenação completa desnecessária)
    ColunasRenomeadas = Table.RenameColumns(ColunasSelecionadas, {
        {"Centro", "Localidade - Origem"},
        {"Município", "Município - Origem"},
        {"UF", "UF - Origem"},
        {"Galpão", "Localidade - Destino"},
        {"Galpãp_Cidade", "Município - Destino"},
        {"Galpãp_UF", "UF - Destino"},
        {"Localidade", "Centro"},
        {"Classificação FDM", "Classificação"},
        {"Status Prazo", "FDM"}
    }),

    // Colunas de preenchimento manual pelo analista (Agrupamento, FA,
    // Obs. Isenção) — usando null com tipo explícito, em vez de ""
    // (texto vazio) do original. Isso é semanticamente mais correto
    // para "ainda não preenchido" e evita erro de tipagem quando a
    // coluna FA (numérica) for usada em cálculos futuros.
    AdicionarAgrupamento = Table.AddColumn(ColunasRenomeadas, "Agrupamento", each null, type text),
    AdicionarFA = Table.AddColumn(AdicionarAgrupamento, "FA", each null, type number),
    AdicionarObsIsencao = Table.AddColumn(AdicionarFA, "Obs. Isenção", each null, type text),

    // Reordenação final ÚNICA — intercala as 3 colunas manuais nas
    // posições corretas em relação às colunas já selecionadas.
    // ? "Contrato" agora posicionada como PRIMEIRA coluna da tabela,
    // conforme solicitado.
    ColunasFinal = Table.ReorderColumns(AdicionarObsIsencao, {
        "Contrato.1",
        "Descrição da atividade", "Item", "Aplicação",
        "Código da#(lf)solicitação", "Data solicitação", "Gerência solicitante",
        "Qtd.#(lf)Solicitada", "Código OS", "D. abertura", "D. fechamento",
        "Prazo combinado", "Qtd.#(lf)Atendida",
        "Agrupamento", "Classificação", "FDM", "Obs. Isenção", "FA",
        "Localidade - Origem", "Município - Origem", "UF - Origem",
        "Localidade - Destino", "Município - Destino", "UF - Destino",
        "KM Adicional", "FC", "QExec",
        "Unidade", "Centro", "Código Centro", "Linha de serviço PPU"
    }),

    // ? Table.Buffer no resultado final — materializa a tabela em
    // memória uma única vez, evitando reprocessar toda a cadeia acima
    // (que já depende do buffer final de #"Painel T2M") caso esta
    // consulta "DADOS" seja referenciada por outras consultas (ex.:
    // SaidaMEDICAO) mais de uma vez.
    ResultadoFinalBuffer = Table.Buffer(ColunasFinal),
    #"Colunas Reordenadas" = Table.ReorderColumns(ResultadoFinalBuffer,{"Contrato.1", "Descrição da atividade", "Item", "Aplicação", "Código da#(lf)solicitação", "Data solicitação", "Gerência solicitante", "Qtd.#(lf)Solicitada", "Código OS", "D. abertura", "D. fechamento", "Prazo combinado", "Qtd.#(lf)Atendida", "Agrupamento", "Classificação", "FDM", "Obs. Isenção", "FA", "Localidade - Origem", "Município - Origem", "UF - Origem", "Localidade - Destino", "Município - Destino", "UF - Destino", "KM Adicional", "FC", "Unidade", "Centro", "Código Centro", "Linha de serviço PPU", "QExec", "QExecAgrupado"})
in
    #"Colunas Reordenadas"

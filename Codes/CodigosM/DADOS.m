let
    Fonte = #"Painel T2M",
    // Table.SelectColumns seleciona E reordena em uma unica operacao -
    // substitui os 3 passos de Table.ReorderColumns + os 2 passos de
    // Table.RemoveColumns do codigo original. "Obs. Sigla" ja e
    // excluida aqui diretamente - no original ela era renomeada para
    // "Obs." e so depois removida, desperdicando processamento.
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
    // As 10 renomeacoes combinadas em UMA unica chamada (o original
    // fazia isso em 6 chamadas separadas de RenameColumns, cada uma
    // intercalada com uma reordenacao completa desnecessaria)
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
    // Obs. Isencao) - usando null com tipo explicito, em vez de ""
    // (texto vazio) do original. Isso e semanticamente mais correto
    // para "ainda nao preenchido" e evita erro de tipagem quando a
    // coluna FA (numerica) for usada em calculos futuros.
    AdicionarAgrupamento = Table.AddColumn(ColunasRenomeadas, "Agrupamento", each null, type text),
    AdicionarFA = Table.AddColumn(AdicionarAgrupamento, "FA", each null, type number),
    AdicionarObsIsencao = Table.AddColumn(AdicionarFA, "Obs. Isenção", each null, type text),
    // Reordenacao final UNICA - intercala as 3 colunas manuais nas
    // posicoes corretas em relacao as colunas ja selecionadas.
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
    // ? Table.Buffer no resultado final - materializa a tabela em
    // memoria uma unica vez, evitando reprocessar toda a cadeia acima
    // (que ja depende do buffer final de #"Painel T2M") caso esta
    // consulta "DADOS" seja referenciada por outras consultas (ex.:
    // SaidaMEDICAO) mais de uma vez.
    ResultadoFinalBuffer = Table.Buffer(ColunasFinal),
    #"Colunas Reordenadas" = Table.ReorderColumns(ResultadoFinalBuffer,{"Contrato.1", "Descrição da atividade", "Item", "Aplicação", "Código da#(lf)solicitação", "Data solicitação", "Gerência solicitante", "Qtd.#(lf)Solicitada", "Código OS", "D. abertura", "D. fechamento", "Prazo combinado", "Qtd.#(lf)Atendida", "Agrupamento", "Classificação", "FDM", "Obs. Isenção", "FA", "Localidade - Origem", "Município - Origem", "UF - Origem", "Localidade - Destino", "Município - Destino", "UF - Destino", "KM Adicional", "FC", "Unidade", "Centro", "Código Centro", "Linha de serviço PPU", "QExec", "QExecAgrupado"})
in
    #"Colunas Reordenadas"

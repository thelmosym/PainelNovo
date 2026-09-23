let
    // ── 1. CARREGAMENTO E EXPANSÃO ──────────────────────────────────────
    Fonte = Folder.Files("C:\Users\DOK4\OneDrive - PETROBRAS\Área de Trabalho\Thelmo\Monitoramento\2026\07-Julho"),

    ArquivosVisiveis = Table.SelectRows(Fonte, each [Attributes]?[Hidden]? <> true),

    ArquivosTransformados = Table.AddColumn(ArquivosVisiveis, "Transformar Arquivo", each #"Transformar Arquivo"([Content])),

    RenomearOrigem = Table.RenameColumns(ArquivosTransformados, {"Name", "Nome da Origem"}),

    SelecionarColunas = Table.SelectColumns(RenomearOrigem, {"Nome da Origem", "Transformar Arquivo"}),

    ExpandirTabela = Table.ExpandTableColumn(
        SelecionarColunas,
        "Transformar Arquivo",
        Table.ColumnNames(#"Transformar Arquivo"(#"Arquivo de Amostra"))
    ),

    // ✅ BUFFER 1 — após expansão dos arquivos
    // Evita reler e reprocessar os arquivos em cada etapa seguinte
    ExpandirTabelaBuffer = Table.Buffer(ExpandirTabela),

    // ── 2. LIMPEZA INICIAL ──────────────────────────────────────────────
    RemoverColunas = Table.RemoveColumns(ExpandirTabelaBuffer, {
        "Nome da Origem", "Atendente", "UF", "Nome Solicitante"
    }),

    // ── 3. FILTROS ──────────────────────────────────────────────────────
    FiltrarContratos = Table.SelectRows(RemoverColunas, each
        ([#"Período#(lf)faturamento"] <> null) and
        ([Contrato] = "IRON-LT1" or [Contrato] = "PA-LT2")
    ),

    // ✅ BUFFER 2 — após primeiro filtro significativo
    // Cacheia o subconjunto reduzido antes das transformações de período
    FiltrarContratosBuffer = Table.Buffer(FiltrarContratos),

    // ── 4. VALIDAR PERÍODO ──────────────────────────────────────────────
    ValidarPeriodo = Table.AddColumn(FiltrarContratosBuffer, "ValidarPeriodo", each
        let
            sPeriodo   = gerar_periodo(PeriodoMes, PeriodoAno),
            partes     = Text.Split(sPeriodo, " à "),
            dataInicio = Date.FromText(partes{0}, [Format="dd/MM/yyyy"]),
            dataFim    = Date.FromText(partes{1}, [Format="dd/MM/yyyy"]),
            dataFat    = Date.From([#"Período#(lf)faturamento"])
        in
            if dataFat >= dataInicio and dataFat <= dataFim then "S" else "N"
    ),

    FiltrarPeriodoValido = Table.SelectRows(ValidarPeriodo, each [ValidarPeriodo] = "S"),

    // ✅ BUFFER 3 — após filtro de período
    // Cacheia antes das transformações de tipo e texto
    FiltrarPeriodoValidoBuffer = Table.Buffer(FiltrarPeriodoValido),

    // ── 5. RENOMEAR, REORDENAR E REMOVER COLUNAS FINAIS ─────────────────
    RenomearColunas = Table.RenameColumns(FiltrarPeriodoValidoBuffer, {{"D. solicitação", "Data solicitação"}}),

    RemoverColunasFinais = Table.RemoveColumns(RenomearColunas, {
        "OS#(lf)disponibilizada?", "LOG",
        "Caixa#(lf)(20kg)", "Caixa#(lf)(Mídia)",
        "Caixa Tubo#(lf)(perfil de poço)", "Caixa Tubo#(lf)(Engenharia)",
        "Lacre", "Etiqueta (20kg)", "Etiqueta (Mídia)",
        "Etiqueta Tubo#(lf)(perfil de poço)", "Etiqueta Tubo#(lf)(Engenharia)",
        "Lançamento de etiqueta/Registro OS", "Observação"
    }),

    ReordenarColunas = Table.ReorderColumns(RemoverColunasFinais, {
        "Período#(lf)faturamento", "Situação", "Contrato",
        "Descrição da atividade", "Item", "Aplicação",
        "Código da#(lf)solicitação", "Chave#(lf)solicitante",
        "Gerência solicitante", "Qtd.#(lf)Solicitada", "Qtd.#(lf)Atendida",
        "Código OS", "Data solicitação", "Prazo#(lf)aplicação",
        "D. abertura", "D. fechamento", "Prazo combinado",
        "Obs. Sigla", "Localidade", "Comentários", "Observações"
    }),

    // ── 6. TIPOS — definição única ──────────────────────────────────────
    DefinirTipos = Table.TransformColumnTypes(ReordenarColunas, {
        {"Período#(lf)faturamento",   type date},
        {"Situação",                  type text},
        {"Contrato",                  type text},
        {"Descrição da atividade",    type text},
        {"Item",                      type text},
        {"Aplicação",                 type text},
        {"Código da#(lf)solicitação", type text},
        {"Chave#(lf)solicitante",     type text},
        {"Gerência solicitante",      type text},
        {"Código OS",                 type text},
        {"Qtd.#(lf)Atendida",        Int64.Type},
        {"Qtd.#(lf)Solicitada",      Int64.Type},
        {"Data solicitação",          type datetime},
        {"Prazo#(lf)aplicação",      type datetime},
        {"D. abertura",               type datetime},
        {"D. fechamento",             type datetime},
        {"Prazo combinado",           type datetime},
        {"Obs. Sigla",                type text},
        {"Localidade",                type text},
        {"Comentários",               type text},
        {"Observações",               type text}
    }),

    // ── 7. TRANSFORMAÇÕES DE TEXTO ──────────────────────────────────────
    TransformarTexto = Table.TransformColumns(DefinirTipos, {
        {"Descrição da atividade", Text.Upper, type text},
        {"Localidade", each Text.Trim(Text.Clean(Text.Upper(_))), type text}
    }),

    // ── 8. FILTRAR SITUAÇÃO E ORDENAR ───────────────────────────────────
    FiltrarSituacao = Table.SelectRows(TransformarTexto, each [Situação] = "CO"),

    // ✅ BUFFER 4 — após filtro de situação
    // Cacheia o conjunto final de dados antes das mesclagens
    FiltrarSituacaoBuffer = Table.Buffer(FiltrarSituacao),

    OrdenarContrato = Table.Sort(FiltrarSituacaoBuffer, {{"Contrato", Order.Ascending}}),

    // ── 9. ENRIQUECER COM LOCALIDADE E GALPÃO ───────────────────────────
    MesclarLocalidade = Table.NestedJoin(
        OrdenarContrato, {"Localidade"},
        #"Localidade Com Endereço", {"Localidade"},
        "Localidade Com Endereço", JoinKind.LeftOuter
    ),

    ExpandirLocalidade = Table.ExpandTableColumn(
        MesclarLocalidade, "Localidade Com Endereço",
        {"Centro", "Município", "UF"},
        {"Centro", "Município", "UF"}
    ),

    AdicionarValidarGalpao = Table.AddColumn(ExpandirLocalidade, "ValidarGalpao", each
        if List.Contains({
            "DEVOLUCAO DE EMPRESTIMO",
            "COLETA DE EMBALAGEM",
            "COLETA DE ITEM AVULSO",
            "ENTREGA DE EMBALAGEM EXPRESSO",
            "ENTREGA DE EMBALAGEM NORMAL",
            "ENTREGA DE ITEM EXPRESSO",
            "ENTREGA DE ITEM NORMAL"
        }, Text.Upper(Text.Trim([Descrição da atividade])))
        then "Verdadeiro"
        else "Falso"
    ),

    AdicionarPersonalizar = Table.AddColumn(AdicionarValidarGalpao, "Personalizar", each
        if [Obs. Sigla] = "GRJ" then "IRON-RJ"
        else if [Obs. Sigla] = "GES" then "IRON-ES"
        else if [Obs. Sigla] = "GSP" then "IRON-SP"
        else if [Obs. Sigla] = "GBA" then "PA-BA"
        else null
    ),

    MesclarUFGalpao = Table.NestedJoin(
        AdicionarPersonalizar, {"UF"},
        validUFGalpao, {"UF"},
        "validUFGalpao", JoinKind.LeftOuter
    ),

    ExpandirUFGalpao = Table.ExpandTableColumn(
        MesclarUFGalpao, "validUFGalpao",
        {"Contrato"}, {"Contrato.1"}
    ),

    AdicionarGalpao = Table.AddColumn(ExpandirUFGalpao, "Galpão", each
        if [ValidarGalpao] = "Falso"           then "N/A"
        else if [ValidarGalpao] = "Verdadeiro" then [Contrato.1]
        else null
    ),

    MesclarEndGalpao = Table.NestedJoin(
        AdicionarGalpao, {"Galpão"},
        valEnderecoGalpao, {"Galpãp"},
        "valEnderecoGalpao", JoinKind.LeftOuter
    ),

    ExpandirEndGalpao = Table.ExpandTableColumn(
        MesclarEndGalpao, "valEnderecoGalpao",
        {"Galpãp_Cidade", "Galpãp_UF"},
        {"Galpãp_Cidade", "Galpãp_UF"}
    ),

    // ── 10. COLUNAS DE CÁLCULO ──────────────────────────────────────────
    AdicionarColunaConcatenada = Table.AddColumn(ExpandirEndGalpao, "ColunaConcatenadaCalculoFDN", each
        [Descrição da atividade] & "/" & [#"Código da#(lf)solicitação"] & "/" & [Código OS]
    ),

    // Adiciona índice para rastrear posição original de cada linha
    AdicionarIndice = Table.AddIndexColumn(AdicionarColunaConcatenada, "Indice", 0, 1, Int64.Type),

    // ✅ BUFFER 5 — antes do Group By e do Merge
    // Garante que o índice seja estável e evita dupla avaliação
    AdicionarIndiceBuffer = Table.Buffer(AdicionarIndice),

    // Agrupa para encontrar o menor índice de cada valor (= primeira ocorrência)
    PrimeiraOcorrencia = Table.Group(
        AdicionarIndiceBuffer,
        {"ColunaConcatenadaCalculoFDN"},
        {{"IndiceMinimo", each List.Min([Indice]), Int64.Type}}
    ),

    // ✅ BUFFER 6 — cacheia o resultado do Group By antes do Merge
    // Evita que o agrupamento seja recalculado durante a mesclagem
    PrimeiraOcorrenciaBuffer = Table.Buffer(PrimeiraOcorrencia),

    // Mescla o índice mínimo de volta na tabela principal
    MesclarPrimeiraOcorrencia = Table.NestedJoin(
        AdicionarIndiceBuffer, {"ColunaConcatenadaCalculoFDN"},
        PrimeiraOcorrenciaBuffer, {"ColunaConcatenadaCalculoFDN"},
        "PrimeiraOc", JoinKind.LeftOuter
    ),

    ExpandirIndiceMinimo = Table.ExpandTableColumn(
        MesclarPrimeiraOcorrencia, "PrimeiraOc",
        {"IndiceMinimo"}, {"IndiceMinimo"}
    ),

    // S = primeira ocorrência | N = ocorrências seguintes
    AdicionarCalculoFDM = Table.AddColumn(ExpandirIndiceMinimo, "CalculoFDM", each
        if [Indice] = [IndiceMinimo] then "N" else "S"
    ),

    // Remove colunas auxiliares
    RemoverAuxiliares = Table.RemoveColumns(AdicionarCalculoFDM, {"Indice", "IndiceMinimo"}),
    #"Consultas Mescladas" = Table.NestedJoin(RemoverAuxiliares, {"Descrição da atividade"}, tabRefPrazoFDM, {"Atividade do painel"}, "tabRefPrazoFDM", JoinKind.LeftOuter),
    #"tabRefPrazoFDM Expandido" = Table.ExpandTableColumn(#"Consultas Mescladas", "tabRefPrazoFDM", {"Ref. FDM", "Ref. QExec", "Ref. prazo"}, {"Ref. FDM", "Ref. QExec", "Ref. prazo"}),
    #"Personalização Adicionada" = Table.AddColumn(#"tabRefPrazoFDM Expandido", "codContDataMuniUF", each [Município]&"\"&[UF]),
    #"Consultas Mescladas1" = Table.NestedJoin(#"Personalização Adicionada", {"codContDataMuniUF"}, Grupo, {"codGrupo"}, "Grupo", JoinKind.LeftOuter),
    #"Grupo Expandido" = Table.ExpandTableColumn(#"Consultas Mescladas1", "Grupo", {"Grupo"}, {"Grupo.1"}),
    #"Personalização Adicionada1" = Table.AddColumn(#"Grupo Expandido", "codPrazoContrato", each [Contrato]
&"\"&
[Descrição da atividade]
&"\"&
[Grupo.1]),
    #"Consultas Mescladas2" = Table.NestedJoin(#"Personalização Adicionada1", {"codPrazoContrato"}, tabelaPrazoContrato, {"codPrazoContrato"}, "tabelaPrazoContrato", JoinKind.LeftOuter),
    #"tabelaPrazoContrato Expandido" = Table.ExpandTableColumn(#"Consultas Mescladas2", "tabelaPrazoContrato", {"Prazo"}, {"tabelaPrazoContrato.Prazo"}),
    #"Data Inserida" = Table.AddColumn(#"tabelaPrazoContrato Expandido", "Data", each DateTime.Date([D. abertura]), type date),
    #"Colunas Renomeadas" = Table.RenameColumns(#"Data Inserida",{{"Data", "Data Abertura"}}),
    #"Hora Inserida" = Table.AddColumn(#"Colunas Renomeadas", "Hora", each DateTime.Time([D. abertura]), type time),
    #"Colunas Renomeadas1" = Table.RenameColumns(#"Hora Inserida",{{"Hora", "Hora Abertura"}}),
    #"Data Inserida1" = Table.AddColumn(#"Colunas Renomeadas1", "Data", each DateTime.Date([D. fechamento]), type date),
    #"Colunas Renomeadas2" = Table.RenameColumns(#"Data Inserida1",{{"Data", "Data Fechamento"}}),
    #"Hora Inserida1" = Table.AddColumn(#"Colunas Renomeadas2", "Hora", each DateTime.Time([D. fechamento]), type time),
    #"Colunas Renomeadas3" = Table.RenameColumns(#"Hora Inserida1",{{"Hora", "Hora Fechamento"}})
in
    #"Colunas Renomeadas3"
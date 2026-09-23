let

// -------------------------------------------------------------------
// SEÇÃO 0 — TABELAS DE REFERÊNCIA (BUFFER AQUI, NÃO NA TABELA PRINCIPAL)
// -------------------------------------------------------------------
RefLocalidadeComEndereco = Table.Buffer(#"Localidade Com Endereço"),
RefValidUFGalpao = Table.Buffer(validUFGalpao),
RefValEnderecoGalpao = Table.Buffer(valEnderecoGalpao),
RefTabRefPrazoFDM = Table.Buffer(tabRefPrazoFDM),
RefGrupo = Table.Buffer(Grupo),
RefTabelaPrazoContrato = Table.Buffer(
    Table.TransformColumnTypes(tabelaPrazoContrato, {{"Prazo", type number}})
),

// -------------------------------------------------------------------
// SEÇÃO 1 — CARREGAMENTO E EXPANSÃO
// -------------------------------------------------------------------
Fonte = Folder.Files(localMonitIndividualT2M),
ArquivosVisiveis = Table.SelectRows(Fonte, each [Attributes]?[Hidden]? <> true),
ArquivosTransformados = Table.AddColumn(ArquivosVisiveis, "Transformar Arquivo", each #"Transformar Arquivo"([Content])),
RenomearOrigem = Table.RenameColumns(ArquivosTransformados, {"Name", "Nome da Origem"}),
SelecionarColunas = Table.SelectColumns(RenomearOrigem, {"Nome da Origem", "Transformar Arquivo"}),
ExpandirTabela = Table.ExpandTableColumn(
    SelecionarColunas,
    "Transformar Arquivo",
    Table.ColumnNames(#"Transformar Arquivo"(#"Arquivo de Amostra"))
),

// BUFFER 1 — evita reler e reprocessar os arquivos físicos
ExpandirTabelaBuffer = Table.Buffer(ExpandirTabela),

// -------------------------------------------------------------------
// SEÇÃO 2 — LIMPEZA, TIPAGEM E FILTRO DE SITUAÇÃO
// -------------------------------------------------------------------
RemoverColunas = Table.RemoveColumns(ExpandirTabelaBuffer, {
    "Nome da Origem", "Atendente", "UF", "Nome Solicitante"
}),

DefinirTiposIniciais = Table.TransformColumnTypes(RemoverColunas, {
    {"Situação", type text},
    {"Contrato", type text},
    {"Período#(lf)faturamento", type date}
}),

FiltrarSituacao = Table.SelectRows(DefinirTiposIniciais, each [Situação] = "CO"),

FiltrarContratos = Table.SelectRows(FiltrarSituacao, each
    ([#"Período#(lf)faturamento"] <> null) and
    ([Contrato] = "IRON-LT1" or [Contrato] = "PA-LT2")
),

// BUFFER 2 — cacheia o subconjunto já bem reduzido (Situação + Contrato)
FiltrarContratosBuffer = Table.Buffer(FiltrarContratos),

// -------------------------------------------------------------------
// SEÇÃO 3 — VALIDAÇÃO DE PERÍODO
// -------------------------------------------------------------------
ValidarPeriodo = Table.AddColumn(FiltrarContratosBuffer, "ValidarPeriodo", each
    let
        sPeriodo = fnGerarPeriodo(PeriodoMes, PeriodoAno),
        partes = Text.Split(sPeriodo, " à "),
        dataInicio = Date.FromText(partes{0}, [Format="dd/MM/yyyy"]),
        dataFim = Date.FromText(partes{1}, [Format="dd/MM/yyyy"]),
        dataFat = Date.From([#"Período#(lf)faturamento"])
    in
        if dataFat >= dataInicio and dataFat <= dataFim then "S" else "N"
),

FiltrarPeriodoValido = Table.SelectRows(ValidarPeriodo, each [ValidarPeriodo] = "S"),
RemoverColunaValidacao = Table.RemoveColumns(FiltrarPeriodoValido, {"ValidarPeriodo"}),

// BUFFER 3 — cacheia antes das transformações de texto e reordenação
FiltrarPeriodoValidoBuffer = Table.Buffer(RemoverColunaValidacao),

// -------------------------------------------------------------------
// SEÇÃO 4 — RENOMEAR, REORDENAR, REMOVER COLUNAS FINAIS E TIPAR O RESTO
// -------------------------------------------------------------------
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

DefinirTiposRestantes = Table.TransformColumnTypes(ReordenarColunas, {
    {"Descrição da atividade", type text},
    {"Item", type text},
    {"Aplicação", type text},
    {"Código da#(lf)solicitação", type text},
    {"Chave#(lf)solicitante", type text},
    {"Gerência solicitante", type text},
    {"Código OS", type text},
    {"Qtd.#(lf)Atendida", Int64.Type},
    {"Qtd.#(lf)Solicitada", Int64.Type},
    {"Data solicitação", type datetime},
    {"Prazo#(lf)aplicação", type datetime},
    {"D. abertura", type datetime},
    {"D. fechamento", type datetime},
    {"Prazo combinado", type datetime},
    {"Obs. Sigla", type text},
    {"Localidade", type text},
    {"Comentários", type text},
    {"Observações", type text}
}),

// -------------------------------------------------------------------
// SEÇÃO 5 — TRANSFORMAÇÕES DE TEXTO E ORDENAÇÃO
// -------------------------------------------------------------------
TransformarTexto = Table.TransformColumns(DefinirTiposRestantes, {
    {"Descrição da atividade", Text.Upper, type text},
    {"Localidade", each Text.Trim(Text.Clean(Text.Upper(_))), type text}
}),

OrdenarContrato = Table.Sort(TransformarTexto, {{"Contrato", Order.Ascending}}),

// BUFFER 4 — conjunto final antes das mesclagens
OrdenarContratoBuffer = Table.Buffer(OrdenarContrato),

// -------------------------------------------------------------------
// SEÇÃO 6 — ENRIQUECER COM LOCALIDADE E GALPÃO
// -------------------------------------------------------------------
MesclarLocalidade = Table.NestedJoin(
    OrdenarContratoBuffer, {"Localidade"},
    RefLocalidadeComEndereco, {"Localidade"},
    "Localidade Com Endereço", JoinKind.LeftOuter
),

ExpandirLocalidade = Table.ExpandTableColumn(MesclarLocalidade, "Localidade Com Endereço", {"Centro", "Código Centro", "Município", "UF"}, {"Centro", "Código Centro", "Município", "UF"}),

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

AdicionarObsSigla = Table.AddColumn(AdicionarValidarGalpao, "ObsSiglaGalpao", each
    if [Obs. Sigla] = "GRJ" then "IRON-RJ"
    else if [Obs. Sigla] = "GES" then "IRON-ES"
    else if [Obs. Sigla] = "GSP" then "IRON-SP"
    else if [Obs. Sigla] = "GBA" then "PA-BA"
    else null
),

MesclarUFGalpao = Table.NestedJoin(
    AdicionarObsSigla, {"UF"},
    RefValidUFGalpao, {"UF"},
    "validUFGalpao", JoinKind.LeftOuter
),

ExpandirUFGalpao = Table.ExpandTableColumn(
    MesclarUFGalpao, "validUFGalpao",
    {"Contrato"}, {"Contrato.1"}
),

AdicionarGalpao = Table.AddColumn(ExpandirUFGalpao, "Galpão", each
    if [ValidarGalpao] = "Falso" then "N/A"
    else if [ValidarGalpao] = "Verdadeiro" then [Contrato.1]
    else null
),

MesclarEndGalpao = Table.NestedJoin(
    AdicionarGalpao, {"Galpão"},
    RefValEnderecoGalpao, {"Galpãp"},
    "valEnderecoGalpao", JoinKind.LeftOuter
),

ExpandirEndGalpao = Table.ExpandTableColumn(
    MesclarEndGalpao, "valEnderecoGalpao",
    {"Galpãp_Cidade", "Galpãp_UF"},
    {"Galpãp_Cidade", "Galpãp_UF"}
),

// -------------------------------------------------------------------
// SEÇÃO 7 — CÁLCULO DE DUPLICATAS (FDM)
// -------------------------------------------------------------------
AdicionarColunaConcatenada = Table.AddColumn(ExpandirEndGalpao, "ColunaConcatenadaCalculoFDN", each
    [Descrição da atividade] & "/" & [#"Código da#(lf)solicitação"] & "/" & [Código OS]
),

AdicionarIndice = Table.AddIndexColumn(AdicionarColunaConcatenada, "Indice", 0, 1, Int64.Type),

// BUFFER 5 — garante índice estável antes do Group By / Merge
AdicionarIndiceBuffer = Table.Buffer(AdicionarIndice),

PrimeiraOcorrencia = Table.Group(
    AdicionarIndiceBuffer,
    {"ColunaConcatenadaCalculoFDN"},
    {{"IndiceMinimo", each List.Min([Indice]), Int64.Type}}
),

// BUFFER 6 — cacheia o resultado do agrupamento antes do merge
PrimeiraOcorrenciaBuffer = Table.Buffer(PrimeiraOcorrencia),

MesclarPrimeiraOcorrencia = Table.NestedJoin(
    AdicionarIndiceBuffer, {"ColunaConcatenadaCalculoFDN"},
    PrimeiraOcorrenciaBuffer, {"ColunaConcatenadaCalculoFDN"},
    "PrimeiraOc", JoinKind.LeftOuter
),

ExpandirIndiceMinimo = Table.ExpandTableColumn(
    MesclarPrimeiraOcorrencia, "PrimeiraOc",
    {"IndiceMinimo"}, {"IndiceMinimo"}
),

// S = duplicata (ocorrência seguinte) | N = primeira ocorrência
AdicionarCalculoFDM = Table.AddColumn(ExpandirIndiceMinimo, "CalculoFDM", each
    if [Indice] = [IndiceMinimo] then "N" else "S"
),

RemoverAuxiliares = Table.RemoveColumns(AdicionarCalculoFDM, {"Indice", "IndiceMinimo"}),

// -------------------------------------------------------------------
// SEÇÃO 8 — REFERÊNCIAS DE PRAZO, GRUPO E CONTRATO
// -------------------------------------------------------------------
MesclarRefPrazoFDM = Table.NestedJoin(
    RemoverAuxiliares, {"Descrição da atividade"},
    RefTabRefPrazoFDM, {"Atividade do painel"},
    "tabRefPrazoFDM", JoinKind.LeftOuter
),

ExpandirRefPrazoFDM = Table.ExpandTableColumn(
    MesclarRefPrazoFDM, "tabRefPrazoFDM",
    {"Ref. FDM", "Ref. QExec", "Ref. prazo"},
    {"Ref. FDM", "Ref. QExec", "Ref. prazo"}
),

AdicionarChaveMunicipioUF = Table.AddColumn(ExpandirRefPrazoFDM, "codContDataMuniUF", each
    [Município] & "\" & [UF]
),

MesclarGrupoMunicipio = Table.NestedJoin(
    AdicionarChaveMunicipioUF, {"codContDataMuniUF"},
    RefGrupo, {"codGrupo"},
    "Grupo", JoinKind.LeftOuter
),

ExpandirGrupoMunicipio = Table.ExpandTableColumn(
    MesclarGrupoMunicipio, "Grupo",
    {"Grupo"}, {"Grupo.1"}
),

AdicionarChavePrazoContrato = Table.AddColumn(ExpandirGrupoMunicipio, "codPrazoContrato", each
    [Contrato] & "\" & [Descrição da atividade] & "\" & [Grupo.1]
),

MesclarPrazoContrato = Table.NestedJoin(
    AdicionarChavePrazoContrato, {"codPrazoContrato"},
    RefTabelaPrazoContrato, {"codPrazoContrato"},
    "tabelaPrazoContrato", JoinKind.LeftOuter
),

ExpandirPrazoContrato = Table.ExpandTableColumn(
    MesclarPrazoContrato, "tabelaPrazoContrato",
    {"Prazo"}, {"tabelaPrazoContrato.Prazo"}
),

// -------------------------------------------------------------------
// SEÇÃO 9 — SEPARAÇÃO DE DATA/HORA
// -------------------------------------------------------------------
AdicionarDataHoraSeparadas = Table.AddColumn(ExpandirPrazoContrato, "DataHoraAux", each
    [
        #"Data Abertura" = DateTime.Date([D. abertura]),
        #"Hora Abertura" = DateTime.Time([D. abertura]),
        #"Data Fechamento" = DateTime.Date([D. fechamento]),
        #"Hora Fechamento" = DateTime.Time([D. fechamento])
    ]
),

ExpandirDataHoraSeparadas = Table.ExpandRecordColumn(
    AdicionarDataHoraSeparadas, "DataHoraAux",
    {"Data Abertura", "Hora Abertura", "Data Fechamento", "Hora Fechamento"},
    {"Data Abertura", "Hora Abertura", "Data Fechamento", "Hora Fechamento"}
),

// -------------------------------------------------------------------
// SEÇÃO 10 — CLASSIFICAÇÃO FDM E SLA (com tratamento de erro)
// -------------------------------------------------------------------
AdicionarClassificacaoFDM = Table.AddColumn(ExpandirDataHoraSeparadas, "ResultadoFDM", each
    try
        fnClassificarFDM(
            [CalculoFDM], [#"Ref. FDM"], [#"Ref. prazo"],
            [#"Descrição da atividade"], [Contrato],
            [Data Abertura], [Hora Abertura],
            [D. fechamento], [Prazo combinado],
            [Município], [UF], [Grupo.1],
            [#"tabelaPrazoContrato.Prazo"], [Aplicação]
        )
    otherwise
        [ Classificacao = null, StatusPrazo = null, Log = "Erro no cálculo — verificar dados da linha; " ],
    type record
),

ExpandirClassificacaoFDM = Table.ExpandRecordColumn(
    AdicionarClassificacaoFDM, "ResultadoFDM",
    {"Classificacao", "StatusPrazo", "Log"},
    {"Classificação FDM", "Status Prazo", "LOG SLA"}
),
    cod_Tabela_FC = Table.AddColumn(ExpandirClassificacaoFDM, "cod_Tabela_FC", each [Descrição da atividade]&[Item]),
    #"Consultas Mescladas" = Table.NestedJoin(cod_Tabela_FC, {"cod_Tabela_FC"}, tabela_FC, {"cod_Tabela_FC"}, "tabela_FC", JoinKind.LeftOuter),
    #"tabela_FC Expandido" = Table.ExpandTableColumn(#"Consultas Mescladas", "tabela_FC", {"FC"}, {"FC"}),
    #"Valor Substituído" = Table.ReplaceValue(#"tabela_FC Expandido",null,0,Replacer.ReplaceValue,{"FC"}),
    #"Função Personalizada Invocada" = Table.AddColumn(#"Valor Substituído", "QExec", each fnCalcularQExec([Descrição da atividade], [Item], [#"Qtd.#(lf)Atendida"], null, [FC])),
    #"Consultas Mescladas1" = Table.NestedJoin(#"Função Personalizada Invocada", {"cod_Tabela_FC"}, tabelaAtividadeItemPPU, {"CodLinhaPPU"}, "tabelaAtividadeItemPPU", JoinKind.LeftOuter),
    #"tabelaAtividadeItemPPU Expandido" = Table.ExpandTableColumn(#"Consultas Mescladas1", "tabelaAtividadeItemPPU", {"Item PPU", "Unidade", "Linha de serviço PPU"}, {"Item PPU", "Unidade", "Linha de serviço PPU"}),
    #"Personalização Adicionada" = Table.AddColumn(#"tabelaAtividadeItemPPU Expandido", "codOrigemDestino", each [Centro]&[Município]&[Galpão]&[Galpãp_Cidade]),
    #"Consultas Mescladas2" = Table.NestedJoin(#"Personalização Adicionada", {"codOrigemDestino"}, Origem_Destino, {"codOrigenDestino"}, "Origem_Destino", JoinKind.LeftOuter),
    #"Origem_Destino Expandido" = Table.ExpandTableColumn(#"Consultas Mescladas2", "Origem_Destino", {"KM"}, {"Origem_Destino.KM"}),
    #"Personalização Adicionada1" = Table.AddColumn(#"Origem_Destino Expandido", "KM Adicional", each if [Origem_Destino.KM] = null then null else if [Origem_Destino.KM] <= valKmAdicional
                        then 0
                        else [Origem_Destino.KM] - valKmAdicional),
    #"QExec Agrupado Calculado" = fnCalcularQExecAgrupado(
        #"Personalização Adicionada1",
        "Data Abertura",
        "Localidade",
        "Galpãp_Cidade",
        "QExec"
    )

in
    #"QExec Agrupado Calculado"

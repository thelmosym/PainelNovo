let

// -------------------------------------------------------------------
// SEÇÃO 0 — TABELAS DE REFERÊNCIA (BUFFER AQUI, NÃO NA TABELA PRINCIPAL)
// -------------------------------------------------------------------
// Cada tabela de lookup usada em Table.NestedJoin mais adiante é
// carregada e "bufferizada" (materializada em memória) UMA ÚNICA VEZ
// aqui. Isso evita reabrir os arquivos de origem repetidamente durante
// os merges e garante estabilidade caso a fonte externa seja alterada
// no meio da execução da consulta.
RefLocalidadeComEndereco = Table.Buffer(#"Localidade Com Endereço"), // Localidade ? Centro/Código Centro/Município/UF
RefValidUFGalpao = Table.Buffer(validUFGalpao), // Lista de UFs válidas por contrato de galpão
RefValEnderecoGalpao = Table.Buffer(valEnderecoGalpao), // Código do galpão ? Cidade/UF do galpão
RefTabRefPrazoFDM = Table.Buffer(tabRefPrazoFDM), // Atividade ? Ref. FDM / Ref. QExec / Ref. prazo
RefGrupo = Table.Buffer(Grupo), // Município+UF ? Grupo (referência de prazo)
RefTabelaPrazoContrato = Table.Buffer(
    Table.TransformColumnTypes(tabelaPrazoContrato, {{"Prazo", type number}})
), // Contrato+Atividade+Grupo ? Prazo contratual

// -------------------------------------------------------------------
// SEÇÃO 1 — CARREGAMENTO E EXPANSÃO DOS ARQUIVOS DE MONITORAMENTO
// -------------------------------------------------------------------
// Lê todos os arquivos da pasta de monitoramento individual (T2M),
// aplica a transformação padrão de cada arquivo (função "Transformar
// Arquivo") e expande o conteúdo de todos eles em uma única tabela.
Fonte = Folder.Files(localMonitIndividualT2M),

// Ignora arquivos ocultos/temporários (ex.: "~$arquivo.xlsx" gerados
// pelo Excel enquanto um arquivo está aberto)
ArquivosVisiveis = Table.SelectRows(Fonte, each [Attributes]?[Hidden]? <> true),

// Aplica a função de transformação (parsing) a cada arquivo
ArquivosTransformados = Table.AddColumn(ArquivosVisiveis, "Transformar Arquivo", each #"Transformar Arquivo"([Content])),

RenomearOrigem = Table.RenameColumns(ArquivosTransformados, {"Name", "Nome da Origem"}),
SelecionarColunas = Table.SelectColumns(RenomearOrigem, {"Nome da Origem", "Transformar Arquivo"}),

// Table.ColumnNames sobre "Arquivo de Amostra" apenas obtém o nome
// das colunas resultantes da transformação, para poder expandir a
// coluna de tabelas aninhadas — padrão comum em consultas de
// "combinar arquivos de uma pasta"
ExpandirTabela = Table.ExpandTableColumn(
    SelecionarColunas,
    "Transformar Arquivo",
    Table.ColumnNames(#"Transformar Arquivo"(#"Arquivo de Amostra"))
),

// BUFFER 1 — evita reler e reprocessar os arquivos físicos a cada
// passo seguinte (carregamento de arquivo é a operação mais custosa
// de toda a consulta)
ExpandirTabelaBuffer = Table.Buffer(ExpandirTabela),

// -------------------------------------------------------------------
// SEÇÃO 2 — LIMPEZA, TIPAGEM E FILTRO DE SITUAÇÃO + CONTRATO
// -------------------------------------------------------------------
RemoverColunas = Table.RemoveColumns(ExpandirTabelaBuffer, {
    "Nome da Origem", "Atendente", "UF", "Nome Solicitante"
}),

DefinirTiposIniciais = Table.TransformColumnTypes(RemoverColunas, {
    {"Situação", type text},
    {"Contrato", type text},
    {"Período#(lf)faturamento", type date}
}),

// ?? OTIMIZAÇÃO APLICADA: as duas filtragens originais (Situação="CO"
// e Contrato IRON-LT1/PA-LT2 + período não nulo) faziam DUAS
// passagens completas pela tabela. Combinadas em uma única chamada
// Table.SelectRows com "and", o motor do Power Query avalia todas as
// condições em UMA única passagem por linha.
FiltrarSituacaoEContrato = Table.SelectRows(DefinirTiposIniciais, each
    [Situação] = "CO"
    and [#"Período#(lf)faturamento"] <> null
    and ([Contrato] = "IRON-LT1" or [Contrato] = "PA-LT2")
),

// BUFFER 2 — cacheia o subconjunto já bem reduzido (Situação + Contrato + período preenchido)
FiltrarContratosBuffer = Table.Buffer(FiltrarSituacaoEContrato),

// -------------------------------------------------------------------
// SEÇÃO 3 — VALIDAÇÃO DE PERÍODO
// -------------------------------------------------------------------
// ?? OTIMIZAÇÃO CRÍTICA APLICADA:
// No código original, fnGerarPeriodo(PeriodoMes, PeriodoAno),
// Text.Split(...) e Date.FromText(...) eram recalculados A CADA LINHA
// da tabela (dentro do "each" de Table.AddColumn) — mas PeriodoMes e
// PeriodoAno são PARÂMETROS FIXOS para toda a execução da consulta.
// Ou seja, o mesmo cálculo de texto era refeito milhares de vezes sem
// necessidade. Aqui, a data de início e fim do período de referência
// é calculada UMA ÚNICA VEZ, fora do "each", e reutilizada no filtro.
//
// Além disso, isso elimina a necessidade de: (1) criar uma coluna
// auxiliar "ValidarPeriodo", (2) filtrar por ela, e (3) removê-la
// depois — reduzindo 3 passos do pipeline original para apenas 1.
sPeriodoDefinido = fnGerarPeriodo(PeriodoMes, PeriodoAno),
PartesPeriodo = Text.Split(sPeriodoDefinido, " à "),
DataInicioPeriodo = Date.FromText(PartesPeriodo{0}, [Format = "dd/MM/yyyy"]),
DataFimPeriodo = Date.FromText(PartesPeriodo{1}, [Format = "dd/MM/yyyy"]),

// Filtra diretamente pelo intervalo já calculado (apenas a conversão
// de data da linha em si — Date.From — precisa ocorrer por linha,
// já que depende do valor específico daquela linha)
FiltrarPeriodoValido = Table.SelectRows(FiltrarContratosBuffer, each
    Date.From([#"Período#(lf)faturamento"]) >= DataInicioPeriodo
    and Date.From([#"Período#(lf)faturamento"]) <= DataFimPeriodo
),

// BUFFER 3 — cacheia antes das transformações de texto e reordenação
FiltrarPeriodoValidoBuffer = Table.Buffer(FiltrarPeriodoValido),

// -------------------------------------------------------------------
// SEÇÃO 4 — RENOMEAR, REORDENAR, REMOVER COLUNAS FINAIS E TIPAR O RESTO
// -------------------------------------------------------------------
RenomearColunas = Table.RenameColumns(FiltrarPeriodoValidoBuffer, {{"D. solicitação", "Data solicitação"}}),

// Remove colunas que não são utilizadas em nenhuma etapa posterior
// do cálculo de FDM/SLA/QExec (campos de embalagem/etiqueta física,
// LOG antigo, observações livres, etc.)
RemoverColunasFinais = Table.RemoveColumns(RenomearColunas, {
    "OS#(lf)disponibilizada?", "LOG",
    "Caixa#(lf)(20kg)", "Caixa#(lf)(Mídia)",
    "Caixa Tubo#(lf)(perfil de poço)", "Caixa Tubo#(lf)(Engenharia)",
    "Lacre", "Etiqueta (20kg)", "Etiqueta (Mídia)",
    "Etiqueta Tubo#(lf)(perfil de poço)", "Etiqueta Tubo#(lf)(Engenharia)",
    "Lançamento de etiqueta/Registro OS", "Observação"
}),

// Padroniza a ordem das colunas para facilitar leitura/depuração
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
// Padroniza a Descrição da atividade e a Localidade em CAIXA ALTA,
// removendo espaços/caracteres não-imprimíveis (Text.Clean), para
// garantir correspondência exata nos merges das Seções 6-8
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
// Traz Centro/Código Centro/Município/UF a partir da Localidade
// PETROBRAS informada na solicitação
MesclarLocalidade = Table.NestedJoin(
    OrdenarContratoBuffer, {"Localidade"},
    RefLocalidadeComEndereco, {"Localidade"},
    "Localidade Com Endereço", JoinKind.LeftOuter
),

ExpandirLocalidade = Table.ExpandTableColumn(MesclarLocalidade, "Localidade Com Endereço",
    {"Centro", "Código Centro", "Município", "UF"}, {"Centro", "Código Centro", "Município", "UF"}),

// Marca se a atividade é uma das 7 que permitem cálculo de KM
// adicional/validação de galpão (mesma lista usada em
// fnValidarLocalizacao)
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

// Converte a sigla informada (Obs. Sigla) no código de galpão
// correspondente (ex.: "GRJ" ? "IRON-RJ")
AdicionarObsSigla = Table.AddColumn(AdicionarValidarGalpao, "ObsSiglaGalpao", each
    if [Obs. Sigla] = "GRJ" then "IRON-RJ"
    else if [Obs. Sigla] = "GES" then "IRON-ES"
    else if [Obs. Sigla] = "GSP" then "IRON-SP"
    else if [Obs. Sigla] = "GBA" then "PA-BA"
    else null
),

// Valida se a UF de destino é atendida por algum galpão cadastrado
MesclarUFGalpao = Table.NestedJoin(
    AdicionarObsSigla, {"UF"},
    RefValidUFGalpao, {"UF"},
    "validUFGalpao", JoinKind.LeftOuter
),

ExpandirUFGalpao = Table.ExpandTableColumn(
    MesclarUFGalpao, "validUFGalpao",
    {"Contrato"}, {"Contrato.1"}
),

// Determina o código do Galpão final: "N/A" se a atividade não
// permite (ValidarGalpao="Falso"), ou o galpão da UF de destino
// (ValidarGalpao="Verdadeiro")
AdicionarGalpao = Table.AddColumn(ExpandirUFGalpao, "Galpão", each
    if [ValidarGalpao] = "Falso" then "N/A"
    else if [ValidarGalpao] = "Verdadeiro" then [Contrato.1]
    else null
),

// Traz Cidade/UF do galpão a partir do código do Galpão
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
// SEÇÃO 7 — CÁLCULO DE DUPLICATAS PARA FDM
// -------------------------------------------------------------------
// A mesma solicitação (Atividade + Código da solicitação + Código OS)
// pode gerar múltiplas linhas no monitoramento (ex.: atendimentos de
// digitalização separados por item A4/A3/A2). Para o cálculo de FDM,
// apenas a PRIMEIRA ocorrência de cada combinação deve ser
// considerada "nova" (N); as demais são tratadas como duplicata (S).
//
// TÉCNICA: em vez de um loop percorrendo linha a linha com CountIf
// (como no VBA legado), usa-se Table.Group para encontrar o índice
// mínimo de cada chave, e depois um merge para trazer esse índice de
// volta a cada linha — abordagem muito mais performática, pois evita
// N² comparações.
AdicionarColunaConcatenada = Table.AddColumn(ExpandirEndGalpao, "ColunaConcatenadaCalculoFDN", each
    [Descrição da atividade] & "/" & [#"Código da#(lf)solicitação"] & "/" & [Código OS]
),

AdicionarIndice = Table.AddIndexColumn(AdicionarColunaConcatenada, "Indice", 0, 1, Int64.Type),

// BUFFER 5 — garante índice estável antes do Group By / Merge
AdicionarIndiceBuffer = Table.Buffer(AdicionarIndice),

// Para cada chave concatenada, encontra o menor índice (= primeira
// ocorrência daquela combinação na tabela)
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
// Traz da tabelaA (via RefTabRefPrazoFDM) a classificação da
// atividade: se entra no escopo de FDM (Ref. FDM), se conta para
// QExec (Ref. QExec) e qual é a referência de prazo (Tabela x
// Fiscalização)
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

// Monta a chave Município\UF para descobrir o Grupo de referência
// de prazo (usado no cálculo do SLA)
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

// Monta a chave Contrato\Atividade\Grupo para descobrir o Prazo
// contratual específico daquela combinação
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
// Separa Data e Hora de Abertura/Fechamento em colunas independentes
// (necessário para o cálculo de SLA, que trata data e hora
// separadamente — expediente 08h-17h, almoço 12h-13h)
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
// SEÇÃO 10 — CLASSIFICAÇÃO FDM/SLA, QEXEC, KM ADICIONAL
// -------------------------------------------------------------------
// Classifica cada linha em Normal/Expresso/Isento/FDM definido e
// NP/FP, calculando o SLA internamente (fnClassificarFDM). Envolvida
// em try/otherwise para que uma linha com dado inconsistente não
// interrompa o processamento de toda a tabela — o erro é registrado
// na própria coluna de LOG.
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

// ?? CORREÇÃO DE ROBUSTEZ APLICADA: a concatenação original com "&"
// falha com erro se [Item] for null. Usando o operador "??"
// (coalescência de nulo), qualquer campo nulo é tratado como texto
// vazio, preservando o resultado quando ambos os campos existem e
// evitando que a consulta inteira quebre por uma única linha
// incompleta.
AdicionarCodTabelaFC = Table.AddColumn(ExpandirClassificacaoFDM, "cod_Tabela_FC", each
    ([Descrição da atividade] ?? "") & ([Item] ?? "")
),

// Traz o Fator de Conversão (FC) específico daquela Atividade+Item,
// usado por fnCalcularQExec nas atividades de Migração/Digitalização
MesclarTabelaFC = Table.NestedJoin(AdicionarCodTabelaFC, {"cod_Tabela_FC"}, tabela_FC, {"cod_Tabela_FC"}, "tabela_FC", JoinKind.LeftOuter),
ExpandirTabelaFC = Table.ExpandTableColumn(MesclarTabelaFC, "tabela_FC", {"FC"}, {"FC"}),

// Calcula a quantidade padronizada de medição (QExec) por linha,
// conforme a atividade/item/quantidade/fator de conversão
AdicionarQExec = Table.AddColumn(ExpandirTabelaFC, "QExec", each
    fnCalcularQExec([Descrição da atividade], [Item], [#"Qtd.#(lf)Atendida"], null, [FC])
),

// Traz Item PPU, Unidade e Linha de serviço PPU associados à
// combinação Atividade+Item (mesma chave "cod_Tabela_FC" já criada)
MesclarAtividadeItemPPU = Table.NestedJoin(AdicionarQExec, {"cod_Tabela_FC"}, tabelaAtividadeItemPPU, {"CodLinhaPPU"}, "tabelaAtividadeItemPPU", JoinKind.LeftOuter),
ExpandirAtividadeItemPPU = Table.ExpandTableColumn(MesclarAtividadeItemPPU, "tabelaAtividadeItemPPU", {"Item PPU", "Unidade", "Linha de serviço PPU"}, {"Item PPU", "Unidade", "Linha de serviço PPU"}),

// ?? CORREÇÃO DE ROBUSTEZ APLICADA: mesma proteção contra nulos
// (operador "??") na chave de origem/destino, evitando erro se
// Centro/Município/Galpão/Cidade do Galpão estiverem vazios
AdicionarCodOrigemDestino = Table.AddColumn(ExpandirAtividadeItemPPU, "codOrigemDestino", each
    ([Centro] ?? "") & ([Município] ?? "") & ([Galpão] ?? "") & ([Galpãp_Cidade] ?? "")
),

// Traz a distância (KM) cadastrada para a rota Centro?Galpão
MesclarOrigemDestino = Table.NestedJoin(AdicionarCodOrigemDestino, {"codOrigemDestino"}, Origem_Destino, {"codOrigenDestino"}, "Origem_Destino", JoinKind.LeftOuter),
ExpandirOrigemDestinoKM = Table.ExpandTableColumn(MesclarOrigemDestino, "Origem_Destino", {"KM"}, {"Origem_Destino.KM"}),

// -------------------------------------------------------------------
// AdicionarKMAdicional — ? CORRIGIDA (deduplicação por rota+data+tipo)
// -----------------------------------------------------------------
// Equivalente M de CalcularQuilometragemAdicional (Sistema 2, módulo
// A_Calc_Medicao) — que, diferente do Sistema 1 legado
// (subMCGuardaExterna/btMCGuardaExterna, confirmado sem essa
// deduplicação), aplica deduplicação por chave rota+data+tipo de
// serviço via Collection, evitando cobrar o mesmo KM adicional
// múltiplas vezes para atendimentos da mesma viagem.
//
// A chave de deduplicação combina:
//   - codOrigemDestino (Centro+Município+Galpão+Cidade Galpão)
//   - Data Abertura
//   - Linha de serviço PPU (equivalente ao "linhaServico" do VBA,
//     que verifica FRE-NRM/FRE-EXP)
//
// Apenas a PRIMEIRA ocorrência de cada combinação recebe o KM
// Adicional calculado; as demais linhas do mesmo grupo ficam null —
// replicando fielmente o comportamento do Collection VBA
// (chaveConcatenadaProcessada.Add ... If Err.Number = 0).
// -------------------------------------------------------------------

// Monta a chave de deduplicação (rota + data + tipo de serviço)
AdicionarChaveKM = Table.AddColumn(ExpandirOrigemDestinoKM, "_ChaveKM", each
    Text.From([codOrigemDestino] ?? "") & "|"
    & Text.From([Data Abertura], "pt-BR") & "|"
    & Text.From([#"Linha de serviço PPU"] ?? "")
),

// Índice sequencial estável, para localizar a 1ª ocorrência de cada chave
AdicionarIndiceKM = Table.AddIndexColumn(AdicionarChaveKM, "_IndiceKM", 0, 1, Int64.Type),

// Agrupa por chave: acha o índice mínimo (1ª ocorrência) de cada grupo
AgrupadoKM = Table.Group(AdicionarIndiceKM, {"_ChaveKM"}, {
    {"_IndiceMinimoKM", each List.Min([_IndiceKM]), Int64.Type}
}),

// Traz de volta o índice mínimo do grupo para cada linha original
MesclarIndiceKM = Table.NestedJoin(AdicionarIndiceKM, {"_ChaveKM"},
    AgrupadoKM, {"_ChaveKM"}, "_GrupoKM", JoinKind.LeftOuter),
ExpandirIndiceKM = Table.ExpandTableColumn(MesclarIndiceKM, "_GrupoKM", {"_IndiceMinimoKM"}, {"_IndiceMinimoKM"}),

// Calcula o KM Adicional SOMENTE na 1ª ocorrência do grupo (mesma
// rota+data+tipo); as demais linhas do grupo ficam null — evitando
// cobrança duplicada do mesmo frete adicional em múltiplos
// atendimentos da mesma viagem. Mantém a regra original de raio de
// tolerância (valKmAdicional): null se rota não encontrada; 0 se
// dentro do raio; excedente acima do raio, caso contrário.
AdicionarKMAdicional = Table.AddColumn(ExpandirIndiceKM, "KM Adicional", each
    if [_IndiceKM] <> [_IndiceMinimoKM] then null
    else if [Origem_Destino.KM] = null then null
    else if [Origem_Destino.KM] <= valKmAdicional then 0
    else [Origem_Destino.KM] - valKmAdicional
),

// Remove as colunas auxiliares de deduplicação, mantendo apenas o
// resultado final "KM Adicional"
LimpezaColunasKM = Table.RemoveColumns(AdicionarKMAdicional, {"_ChaveKM", "_IndiceKM", "_IndiceMinimoKM"}),

// -------------------------------------------------------------------
// SEÇÃO 11 — QEXEC AGRUPADO (cálculo adicional/paralelo)
// -------------------------------------------------------------------
// Cria um índice de agrupamento por Data Abertura + Localidade
// PETROBRAS + Cidade do Galpão; soma o QExec de todas as linhas do
// grupo e atribui o resultado da fórmula de frete (soma/10, piso 1)
// somente na linha de primeira ocorrência do grupo (as demais linhas
// do mesmo grupo ficam com null nesta coluna).

AdicionarQExecAgrupado = fnCalcularQExecAgrupado(
    LimpezaColunasKM,
    "Data Fechamento",
    "Localidade",
    "Galpãp_Cidade",
    "QExec"
),

// -------------------------------------------------------------------
// BUFFER FINAL — RECOMENDADO
// -------------------------------------------------------------------
// Esta consulta é referenciada por diversas outras consultas
// independentes (ex.: SaidaMC, SaidaDADOS, SaidaMEDICAO,
// TabelaFDMPorContrato, SaidaFretes). Sem este buffer, CADA UMA
// dessas consultas reexecuta TODO o pipeline acima (todos os merges,
// chamadas de função customizada, cálculo de FDM/SLA/QExec) do zero,
// de forma independente. Bufferizar o resultado final materializa a
// tabela em memória uma única vez, reduzindo o reprocessamento
// repetido nas consultas dependentes — o custo é maior uso de memória
// durante a atualização.
ResultadoFinalBuffer = Table.Buffer(AdicionarQExecAgrupado)

in
    ResultadoFinalBuffer

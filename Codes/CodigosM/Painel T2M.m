let
// -------------------------------------------------------------------
// SECAO 0 - TABELAS DE REFERENCIA (BUFFER AQUI, NAO NA TABELA PRINCIPAL)
// -------------------------------------------------------------------
// Cada tabela de lookup usada em Table.NestedJoin mais adiante e
// carregada e "bufferizada" (materializada em memoria) UMA UNICA VEZ
// aqui. Isso evita reabrir os arquivos de origem repetidamente durante
// os merges e garante estabilidade caso a fonte externa seja alterada
// no meio da execucao da consulta.
RefLocalidadeComEndereco = Table.Buffer(#"Localidade Com Endereço"), // Localidade -> Centro/Codigo Centro/Municipio/UF
RefValidUFGalpao = Table.Buffer(validUFGalpao), // Lista de UFs validas por contrato de galpao
RefValEnderecoGalpao = Table.Buffer(valEnderecoGalpao), // Codigo do galpao -> Cidade/UF do galpao
RefTabRefPrazoFDM = Table.Buffer(tabRefPrazoFDM), // Atividade -> Ref. FDM / Ref. QExec / Ref. prazo
RefGrupo = Table.Buffer(Grupo),
RefCalendario = Table.Buffer(Calendario), // Municipio+UF -> Grupo (referencia de prazo)
RefTabelaPrazoContrato = Table.Buffer(
    Table.TransformColumnTypes(tabelaPrazoContrato, {{"Prazo", type number}})
), // Contrato+Atividade+Grupo -> Prazo contratual
// -------------------------------------------------------------------
// SECAO 1 - CARREGAMENTO E EXPANSAO DOS ARQUIVOS DE MONITORAMENTO
// -------------------------------------------------------------------
// Le todos os arquivos da pasta de monitoramento individual (T2M),
// aplica a transformacao padrao de cada arquivo (funcao "Transformar
// Arquivo") e expande o conteudo de todos eles em uma unica tabela.
Fonte = Folder.Files(localMonitIndividualT2M),
// Ignora arquivos ocultos/temporarios e garante que apenas pastas de trabalho Excel (.xlsm / .xlsx) sejam processadas
ArquivosVisiveis = Table.SelectRows(Fonte, each [Attributes]?[Hidden]? <> true and not Text.StartsWith([Name], "~$") and ([Extension] = ".xlsm" or [Extension] = ".xlsx")),
// Aplica a funcao de transformacao (parsing) a cada arquivo
ArquivosTransformados = Table.AddColumn(ArquivosVisiveis, "Transformar Arquivo", each #"Transformar Arquivo"([Content])),
RenomearOrigem = Table.RenameColumns(ArquivosTransformados, {"Name", "Nome da Origem"}),
SelecionarColunas = Table.SelectColumns(RenomearOrigem, {"Nome da Origem", "Transformar Arquivo"}),
// Table.ColumnNames sobre "Arquivo de Amostra" apenas obtem o nome
// das colunas resultantes da transformacao, para poder expandir a
// coluna de tabelas aninhadas - padrao comum em consultas de
// "combinar arquivos de uma pasta"
ExpandirTabela = Table.ExpandTableColumn(
    SelecionarColunas,
    "Transformar Arquivo",
    Table.ColumnNames(#"Transformar Arquivo"(#"Arquivo de Amostra"))
),
// BUFFER 1 - evita reler e reprocessar os arquivos fisicos a cada
// passo seguinte (carregamento de arquivo e a operacao mais custosa
// de toda a consulta)
ExpandirTabelaBuffer = Table.Buffer(ExpandirTabela),
// -------------------------------------------------------------------
// SECAO 2 - LIMPEZA, TIPAGEM E FILTRO DE SITUACAO + CONTRATO
// -------------------------------------------------------------------
RemoverColunas = Table.RemoveColumns(ExpandirTabelaBuffer, {
    "Nome da Origem", "Atendente", "UF", "Nome Solicitante"
}),
DefinirTiposIniciais = Table.TransformColumnTypes(RemoverColunas, {
    {"Situação", type text},
    {"Contrato", type text},
    {"Período#(lf)faturamento", type date}
}),
//  OTIMIZACAO APLICADA: as duas filtragens originais (Situacao="CO"
// e Contrato IRON-LT1/PA-LT2 + periodo nao nulo) faziam DUAS
// passagens completas pela tabela. Combinadas em uma unica chamada
// Table.SelectRows com "and", o motor do Power Query avalia todas as
// condicoes em UMA unica passagem por linha.
FiltrarSituacaoEContrato = Table.SelectRows(DefinirTiposIniciais, each
    [Situação] = "CO"
    and [#"Período#(lf)faturamento"] <> null
    and ([Contrato] = "IRON-LT1" or [Contrato] = "PA-LT2")
),
// BUFFER 2 - cacheia o subconjunto ja bem reduzido (Situacao + Contrato + periodo preenchido)
FiltrarContratosBuffer = Table.Buffer(FiltrarSituacaoEContrato),
// -------------------------------------------------------------------
// SECAO 3 - VALIDACAO DE PERIODO
// -------------------------------------------------------------------
//  OTIMIZACAO CRITICA APLICADA:
// No codigo original, fnGerarPeriodo(PeriodoMes, PeriodoAno),
// Text.Split(...) e Date.FromText(...) eram recalculados A CADA LINHA
// da tabela (dentro do "each" de Table.AddColumn) - mas PeriodoMes e
// PeriodoAno sao PARAMETROS FIXOS para toda a execucao da consulta.
// Ou seja, o mesmo calculo de texto era refeito milhares de vezes sem
// necessidade. Aqui, a data de inicio e fim do periodo de referencia
// e calculada UMA UNICA VEZ, fora do "each", e reutilizada no filtro.
//
// Alem disso, isso elimina a necessidade de: (1) criar uma coluna
// auxiliar "ValidarPeriodo", (2) filtrar por ela, e (3) remove-la
// depois - reduzindo 3 passos do pipeline original para apenas 1.
sPeriodoDefinido = fnGerarPeriodo(PeriodoMes, PeriodoAno),
PartesPeriodo = Text.Split(sPeriodoDefinido, " à "),
DataInicioPeriodo = Date.FromText(PartesPeriodo{0}, [Format = "dd/MM/yyyy"]),
DataFimPeriodo = Date.FromText(PartesPeriodo{1}, [Format = "dd/MM/yyyy"]),
// Filtra diretamente pelo intervalo ja calculado (apenas a conversao
// de data da linha em si - Date.From - precisa ocorrer por linha,
// ja que depende do valor especifico daquela linha)
FiltrarPeriodoValido = Table.SelectRows(FiltrarContratosBuffer, each
    Date.From([#"Período#(lf)faturamento"]) >= DataInicioPeriodo
    and Date.From([#"Período#(lf)faturamento"]) <= DataFimPeriodo
),
// BUFFER 3 - cacheia antes das transformacoes de texto e reordenacao
FiltrarPeriodoValidoBuffer = Table.Buffer(FiltrarPeriodoValido),
// -------------------------------------------------------------------
// SECAO 4 - RENOMEAR, REORDENAR, REMOVER COLUNAS FINAIS E TIPAR O RESTO
// -------------------------------------------------------------------
RenomearColunas = Table.RenameColumns(FiltrarPeriodoValidoBuffer, {{"D. solicitação", "Data solicitação"}}),
// Remove colunas que nao sao utilizadas em nenhuma etapa posterior
// do calculo de FDM/SLA/QExec (campos de embalagem/etiqueta fisica,
// LOG antigo, observacoes livres, etc.)
RemoverColunasFinais = Table.RemoveColumns(RenomearColunas, {
    "OS#(lf)disponibilizada?", "LOG",
    "Caixa#(lf)(20kg)", "Caixa#(lf)(Mídia)",
    "Caixa Tubo#(lf)(perfil de poço)", "Caixa Tubo#(lf)(Engenharia)",
    "Lacre", "Etiqueta (20kg)", "Etiqueta (Mídia)",
    "Etiqueta Tubo#(lf)(perfil de poço)", "Etiqueta Tubo#(lf)(Engenharia)",
    "Lançamento de etiqueta/Registro OS", "Observação"
}),
// Padroniza a ordem das colunas para facilitar leitura/depuracao
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
// SECAO 5 - TRANSFORMACOES DE TEXTO E ORDENACAO
// -------------------------------------------------------------------
// Padroniza a Descricao da atividade e a Localidade em CAIXA ALTA,
// removendo espacos/caracteres nao-imprimiveis (Text.Clean), para
// garantir correspondencia exata nos merges das Secoes 6-8
TransformarTexto = Table.TransformColumns(DefinirTiposRestantes, {
    {"Descrição da atividade", Text.Upper, type text},
    {"Localidade", each Text.Trim(Text.Clean(Text.Upper(_))), type text}
}),
OrdenarContrato = Table.Sort(TransformarTexto, {{"Contrato", Order.Ascending}}),
// BUFFER 4 - conjunto final antes das mesclagens
OrdenarContratoBuffer = Table.Buffer(OrdenarContrato),
// -------------------------------------------------------------------
// SECAO 6 - ENRIQUECER COM LOCALIDADE E GALPAO
// -------------------------------------------------------------------
// Traz Centro/Codigo Centro/Municipio/UF a partir da Localidade
// PETROBRAS informada na solicitacao
MesclarLocalidade = Table.NestedJoin(
    OrdenarContratoBuffer, {"Localidade"},
    RefLocalidadeComEndereco, {"Localidade"},
    "Localidade Com Endereço", JoinKind.LeftOuter
),
ExpandirLocalidade = Table.ExpandTableColumn(MesclarLocalidade, "Localidade Com Endereço",
    {"Centro", "Código Centro", "Município", "UF"}, {"Centro", "Código Centro", "Município", "UF"}),
// Marca se a atividade e uma das 7 que permitem calculo de KM
// adicional/validacao de galpao (mesma lista usada em
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
// Converte a sigla informada (Obs. Sigla) no codigo de galpao
// correspondente (ex.: "GRJ" -> "IRON-RJ")
AdicionarObsSigla = Table.AddColumn(AdicionarValidarGalpao, "ObsSiglaGalpao", each
    if [Obs. Sigla] = "GRJ" then "IRON-RJ"
    else if [Obs. Sigla] = "GES" then "IRON-ES"
    else if [Obs. Sigla] = "GSP" then "IRON-SP"
    else if [Obs. Sigla] = "GBA" then "PA-BA"
    else null
),
// Valida se a UF de destino e atendida por algum galpao cadastrado
MesclarUFGalpao = Table.NestedJoin(
    AdicionarObsSigla, {"UF"},
    RefValidUFGalpao, {"UF"},
    "validUFGalpao", JoinKind.LeftOuter
),
ExpandirUFGalpao = Table.ExpandTableColumn(
    MesclarUFGalpao, "validUFGalpao",
    {"Contrato"}, {"Contrato.1"}
),
// Determina o codigo do Galpao final: "N/A" se a atividade nao
// permite (ValidarGalpao="Falso"), ou o galpao da UF de destino
// (ValidarGalpao="Verdadeiro")
AdicionarGalpao = Table.AddColumn(ExpandirUFGalpao, "Galpão", each
    if [ValidarGalpao] = "Falso" then "N/A"
    else if [ValidarGalpao] = "Verdadeiro" then [Contrato.1]
    else null
),
// Traz Cidade/UF do galpao a partir do codigo do Galpao
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
// SECAO 7 - CALCULO DE DUPLICATAS PARA FDM
// -------------------------------------------------------------------
// A mesma solicitacao (Atividade + Codigo da solicitacao + Codigo OS)
// pode gerar multiplas linhas no monitoramento (ex.: atendimentos de
// digitalizacao separados por item A4/A3/A2). Para o calculo de FDM,
// apenas a PRIMEIRA ocorrencia de cada combinacao deve ser
// considerada "nova" (N); as demais sao tratadas como duplicata (S).
//
// TECNICA: em vez de um loop percorrendo linha a linha com CountIf
// (como no VBA legado), usa-se Table.Group para encontrar o indice
// minimo de cada chave, e depois um merge para trazer esse indice de
// volta a cada linha - abordagem muito mais performatica, pois evita
// N2 comparacoes.
AdicionarColunaConcatenada = Table.AddColumn(ExpandirEndGalpao, "ColunaConcatenadaCalculoFDN", each
    [Descrição da atividade] & "/" & [#"Código da#(lf)solicitação"] & "/" & [Código OS]
),
AdicionarIndice = Table.AddIndexColumn(AdicionarColunaConcatenada, "Indice", 0, 1, Int64.Type),
// BUFFER 5 - garante indice estavel antes do Group By / Merge
AdicionarIndiceBuffer = Table.Buffer(AdicionarIndice),
// Para cada chave concatenada, encontra o menor indice (= primeira
// ocorrencia daquela combinacao na tabela)
PrimeiraOcorrencia = Table.Group(
    AdicionarIndiceBuffer,
    {"ColunaConcatenadaCalculoFDN"},
    {{"IndiceMinimo", each List.Min([Indice]), Int64.Type}}
),
// BUFFER 6 - cacheia o resultado do agrupamento antes do merge
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
// S = duplicata (ocorrencia seguinte) | N = primeira ocorrencia
AdicionarCalculoFDM = Table.AddColumn(ExpandirIndiceMinimo, "CalculoFDM", each
    if [Indice] = [IndiceMinimo] then "N" else "S"
),
RemoverAuxiliares = Table.RemoveColumns(AdicionarCalculoFDM, {"Indice", "IndiceMinimo"}),
// -------------------------------------------------------------------
// SECAO 8 - REFERENCIAS DE PRAZO, GRUPO E CONTRATO
// -------------------------------------------------------------------
// Traz da tabelaA (via RefTabRefPrazoFDM) a classificacao da
// atividade: se entra no escopo de FDM (Ref. FDM), se conta para
// QExec (Ref. QExec) e qual e a referencia de prazo (Tabela x
// Fiscalizacao)
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
// Monta a chave Municipio\UF para descobrir o Grupo de referencia
// de prazo (usado no calculo do SLA)
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
// contratual especifico daquela combinacao
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
// SECAO 9 - SEPARACAO DE DATA/HORA
// -------------------------------------------------------------------
// Separa Data e Hora de Abertura/Fechamento em colunas independentes
// (necessario para o calculo de SLA, que trata data e hora
// separadamente - expediente 08h-17h, almoco 12h-13h)
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
// SECAO 10 - CLASSIFICACAO FDM/SLA, QEXEC, KM ADICIONAL
// -------------------------------------------------------------------
// Classifica cada linha em Normal/Expresso/Isento/FDM definido e
// NP/FP, calculando o SLA internamente (fnClassificarFDM). Envolvida
// em try/otherwise para que uma linha com dado inconsistente nao
// interrompa o processamento de toda a tabela - o erro e registrado
// na propria coluna de LOG.
AdicionarClassificacaoFDM = Table.AddColumn(ExpandirDataHoraSeparadas, "ResultadoFDM", each
    try
        fnClassificarFDM(
            [CalculoFDM], [#"Ref. FDM"], [#"Ref. prazo"],
            [#"Descrição da atividade"], [Contrato],
            [Data Abertura], [Hora Abertura],
            [D. fechamento], [Prazo combinado],
            [Município], [UF], [Grupo.1],
            [#"tabelaPrazoContrato.Prazo"], [Aplicação], RefCalendario
        )
    otherwise
        [ Classificacao = null, StatusPrazo = null, Log = "Erro no calculo - verificar dados da linha; " ],
    type record
),
ExpandirClassificacaoFDM = Table.ExpandRecordColumn(
    AdicionarClassificacaoFDM, "ResultadoFDM",
    {"Classificacao", "StatusPrazo", "Log"},
    {"Classificação FDM", "Status Prazo", "LOG SLA"}
),
//  CORRECAO DE ROBUSTEZ APLICADA: a concatenacao original com "&"
// falha com erro se [Item] for null. Usando o operador "??"
// (coalescencia de nulo), qualquer campo nulo e tratado como texto
// vazio, preservando o resultado quando ambos os campos existem e
// evitando que a consulta inteira quebre por uma unica linha
// incompleta.
AdicionarCodTabelaFC = Table.AddColumn(ExpandirClassificacaoFDM, "cod_Tabela_FC", each
    ([Descrição da atividade] ?? "") & ([Item] ?? "")
),
// Traz o Fator de Conversao (FC) especifico daquela Atividade+Item,
// usado por fnCalcularQExec nas atividades de Migracao/Digitalizacao
MesclarTabelaFC = Table.NestedJoin(AdicionarCodTabelaFC, {"cod_Tabela_FC"}, tabela_FC, {"cod_Tabela_FC"}, "tabela_FC", JoinKind.LeftOuter),
ExpandirTabelaFC = Table.ExpandTableColumn(MesclarTabelaFC, "tabela_FC", {"FC"}, {"FC"}),
// Calcula a quantidade padronizada de medicao (QExec) por linha,
// conforme a atividade/item/quantidade/fator de conversao
AdicionarQExec = Table.AddColumn(ExpandirTabelaFC, "QExec", each
    fnCalcularQExec([Descrição da atividade], [Item], [#"Qtd.#(lf)Atendida"], null, [FC])
),
// Traz Item PPU, Unidade e Linha de servico PPU associados a
// combinacao Atividade+Item (mesma chave "cod_Tabela_FC" ja criada)
MesclarAtividadeItemPPU = Table.NestedJoin(AdicionarQExec, {"cod_Tabela_FC"}, tabelaAtividadeItemPPU, {"CodLinhaPPU"}, "tabelaAtividadeItemPPU", JoinKind.LeftOuter),
ExpandirAtividadeItemPPU = Table.ExpandTableColumn(MesclarAtividadeItemPPU, "tabelaAtividadeItemPPU", {"Item PPU", "Unidade", "Linha de serviço PPU"}, {"Item PPU", "Unidade", "Linha de serviço PPU"}),
//  CORRECAO DE ROBUSTEZ APLICADA: mesma protecao contra nulos
// (operador "??") na chave de origem/destino, evitando erro se
// Centro/Municipio/Galpao/Cidade do Galpao estiverem vazios
AdicionarCodOrigemDestino = Table.AddColumn(ExpandirAtividadeItemPPU, "codOrigemDestino", each
    ([Centro] ?? "") & ([Município] ?? "") & ([Galpão] ?? "") & ([Galpãp_Cidade] ?? "")
),
// Traz a distancia (KM) cadastrada para a rota Centro<->Galpao
MesclarOrigemDestino = Table.NestedJoin(AdicionarCodOrigemDestino, {"codOrigemDestino"}, Origem_Destino, {"codOrigenDestino"}, "Origem_Destino", JoinKind.LeftOuter),
ExpandirOrigemDestinoKM = Table.ExpandTableColumn(MesclarOrigemDestino, "Origem_Destino", {"KM"}, {"Origem_Destino.KM"}),
// -------------------------------------------------------------------
// AdicionarKMAdicional - [OK] CORRIGIDA (deduplicacao por rota+data+tipo)
// -----------------------------------------------------------------
// Equivalente M de CalcularQuilometragemAdicional (Sistema 2, modulo
// A_Calc_Medicao) - que, diferente do Sistema 1 legado
// (subMCGuardaExterna/btMCGuardaExterna, confirmado sem essa
// deduplicacao), aplica deduplicacao por chave rota+data+tipo de
// servico via Collection, evitando cobrar o mesmo KM adicional
// multiplas vezes para atendimentos da mesma viagem.
//
// A chave de deduplicacao combina:
//   - codOrigemDestino (Centro+Municipio+Galpao+Cidade Galpao)
//   - Data Abertura
//   - Linha de servico PPU (equivalente ao "linhaServico" do VBA,
//     que verifica FRE-NRM/FRE-EXP)
//
// Apenas a PRIMEIRA ocorrencia de cada combinacao recebe o KM
// Adicional calculado; as demais linhas do mesmo grupo ficam null -
// replicando fielmente o comportamento do Collection VBA
// (chaveConcatenadaProcessada.Add ... If Err.Number = 0).
// -------------------------------------------------------------------
// Monta a chave de deduplicacao (rota + data + tipo de servico)
AdicionarChaveKM = Table.AddColumn(ExpandirOrigemDestinoKM, "_ChaveKM", each
    Text.From([codOrigemDestino] ?? "") & "|"
    & Text.From([Data Abertura], "pt-BR") & "|"
    & Text.From([#"Linha de serviço PPU"] ?? "")
),
// Indice sequencial estavel, para localizar a 1a ocorrencia de cada chave
AdicionarIndiceKM = Table.AddIndexColumn(AdicionarChaveKM, "_IndiceKM", 0, 1, Int64.Type),
// Agrupa por chave: acha o indice minimo (1a ocorrencia) de cada grupo
AgrupadoKM = Table.Group(AdicionarIndiceKM, {"_ChaveKM"}, {
    {"_IndiceMinimoKM", each List.Min([_IndiceKM]), Int64.Type}
}),
// Traz de volta o indice minimo do grupo para cada linha original
MesclarIndiceKM = Table.NestedJoin(AdicionarIndiceKM, {"_ChaveKM"},
    AgrupadoKM, {"_ChaveKM"}, "_GrupoKM", JoinKind.LeftOuter),
ExpandirIndiceKM = Table.ExpandTableColumn(MesclarIndiceKM, "_GrupoKM", {"_IndiceMinimoKM"}, {"_IndiceMinimoKM"}),
// Calcula o KM Adicional SOMENTE na 1a ocorrencia do grupo (mesma
// rota+data+tipo); as demais linhas do grupo ficam null - evitando
// cobranca duplicada do mesmo frete adicional em multiplos
// atendimentos da mesma viagem. Mantem a regra original de raio de
// tolerancia (valKmAdicional): null se rota nao encontrada; 0 se
// dentro do raio; excedente acima do raio, caso contrario.
AdicionarKMAdicional = Table.AddColumn(ExpandirIndiceKM, "KM Adicional", each
    if [_IndiceKM] <> [_IndiceMinimoKM] then null
    else if [Origem_Destino.KM] = null then null
    else if [Origem_Destino.KM] <= valKmAdicional then 0
    else [Origem_Destino.KM] - valKmAdicional
),
// Remove as colunas auxiliares de deduplicacao, mantendo apenas o
// resultado final "KM Adicional"
LimpezaColunasKM = Table.RemoveColumns(AdicionarKMAdicional, {"_ChaveKM", "_IndiceKM", "_IndiceMinimoKM"}),
// -------------------------------------------------------------------
// SECAO 11 - QEXEC AGRUPADO (calculo adicional/paralelo)
// -------------------------------------------------------------------
// Cria um indice de agrupamento por Data Abertura + Localidade
// PETROBRAS + Cidade do Galpao; soma o QExec de todas as linhas do
// grupo e atribui o resultado da formula de frete (soma/10, piso 1)
// somente na linha de primeira ocorrencia do grupo (as demais linhas
// do mesmo grupo ficam com null nesta coluna).
AdicionarQExecAgrupado = fnCalcularQExecAgrupado(
    LimpezaColunasKM,
    "Data Fechamento",
    "Localidade",
    "Galpãp_Cidade",
    "QExec"
),
// -------------------------------------------------------------------
// BUFFER FINAL - RECOMENDADO
// -------------------------------------------------------------------
// Esta consulta e referenciada por diversas outras consultas
// independentes (ex.: SaidaMC, SaidaDADOS, SaidaMEDICAO,
// TabelaFDMPorContrato, SaidaFretes). Sem este buffer, CADA UMA
// dessas consultas reexecuta TODO o pipeline acima (todos os merges,
// chamadas de funcao customizada, calculo de FDM/SLA/QExec) do zero,
// de forma independente. Bufferizar o resultado final materializa a
// tabela em memoria uma unica vez, reduzindo o reprocessamento
// repetido nas consultas dependentes - o custo e maior uso de memoria
// durante a atualizacao.
ResultadoFinalBuffer = Table.Buffer(AdicionarQExecAgrupado)
in
    ResultadoFinalBuffer

let
    Fonte = Folder.Files("C:\Users\DOK4\OneDrive - PETROBRAS\Área de Trabalho\Thelmo\Monitoramento\2026\07-Julho"),
    #"Arquivos Ocultos Filtrados1" = Table.SelectRows(Fonte, each [Attributes]?[Hidden]? <> true),
    #"Invocar Função Personalizada1" = Table.AddColumn(#"Arquivos Ocultos Filtrados1", "Transformar Arquivo", each #"Transformar Arquivo"([Content])),
    #"Colunas Renomeadas1" = Table.RenameColumns(#"Invocar Função Personalizada1", {"Name", "Nome da Origem"}),
    #"Outras Colunas Removidas1" = Table.SelectColumns(#"Colunas Renomeadas1", {"Nome da Origem", "Transformar Arquivo"}),
    #"Coluna de Tabela Expandida1" = Table.ExpandTableColumn(#"Outras Colunas Removidas1", "Transformar Arquivo", Table.ColumnNames(#"Transformar Arquivo"(#"Arquivo de Amostra"))),
    #"Colunas Removidas" = Table.RemoveColumns(#"Coluna de Tabela Expandida1",{"Nome da Origem", "Atendente"}),
    #"Linhas Filtradas" = Table.SelectRows(#"Colunas Removidas", each ([#"Período#(lf)faturamento"] <> null) and ([Contrato] = "IRON-LT1" or [Contrato] = "PA-LT2")),
    ValidarPeriodo = Table.AddColumn(#"Linhas Filtradas", "ValidarPeriodo", each let
    // Gera o período em texto: "26/07/2026 à 25/08/2026"
    sPeriodo = gerar_periodo(PeriodoMes, PeriodoAno),

    // Quebra o texto em duas partes pela palavra " à "
    partes = Text.Split(sPeriodo, " à "),

    // Converte cada parte para tipo Date
    dataInicio = Date.FromText(partes{0}, [Format="dd/MM/yyyy"]),
    dataFim    = Date.FromText(partes{1}, [Format="dd/MM/yyyy"]),

    // Garante que a coluna de faturamento também é Date
    dataFaturamento = Date.From([#"Período#(lf)faturamento"]),

    // Valida se está dentro do período
    resultado = if dataFaturamento >= dataInicio 
                   and dataFaturamento <= dataFim 
                then "S" 
                else "N"
in
    resultado),
    #"Filtrar ValidarPeriodo" = Table.SelectRows(ValidarPeriodo, each ([ValidarPeriodo] = "S")),
    #"Colunas Reordenadas" = Table.ReorderColumns(#"Filtrar ValidarPeriodo",{"Período#(lf)faturamento", "Situação", "Contrato", "Descrição da atividade", "Item", "Aplicação", "Código da#(lf)solicitação", "Chave#(lf)solicitante", "D. solicitação", "UF", "Nome Solicitante", "Prazo#(lf)aplicação", "Gerência solicitante", "Qtd.#(lf)Solicitada", "Código OS", "D. abertura", "D. fechamento", "Prazo combinado", "Qtd.#(lf)Atendida", "Obs. Sigla", "Localidade", "Comentários", "Observações", "OS#(lf)disponibilizada?", "LOG", "Caixa#(lf)(20kg)", "Caixa#(lf)(Mídia)", "Caixa Tubo#(lf)(perfil de poço)", "Caixa Tubo#(lf)(Engenharia)", "Lacre", "Etiqueta (20kg)", "Etiqueta (Mídia)", "Etiqueta Tubo#(lf)(perfil de poço)", "Etiqueta Tubo#(lf)(Engenharia)", "Lançamento de etiqueta/Registro OS", "Observação"}),
    #"Colunas Renomeadas" = Table.RenameColumns(#"Colunas Reordenadas",{{"D. solicitação", "Data solicitação"}}),
    #"Colunas Removidas1" = Table.RemoveColumns(#"Colunas Renomeadas",{"UF", "Nome Solicitante"}),
    #"Colunas Reordenadas1" = Table.ReorderColumns(#"Colunas Removidas1",{"Período#(lf)faturamento", "Situação", "Contrato", "Descrição da atividade", "Item", "Aplicação", "Código da#(lf)solicitação", "Chave#(lf)solicitante", "Gerência solicitante", "Qtd.#(lf)Solicitada", "Qtd.#(lf)Atendida", "Código OS", "Data solicitação", "Prazo#(lf)aplicação", "D. abertura", "D. fechamento", "Prazo combinado", "Obs. Sigla", "Localidade", "Comentários", "Observações", "OS#(lf)disponibilizada?", "LOG", "Caixa#(lf)(20kg)", "Caixa#(lf)(Mídia)", "Caixa Tubo#(lf)(perfil de poço)", "Caixa Tubo#(lf)(Engenharia)", "Lacre", "Etiqueta (20kg)", "Etiqueta (Mídia)", "Etiqueta Tubo#(lf)(perfil de poço)", "Etiqueta Tubo#(lf)(Engenharia)", "Lançamento de etiqueta/Registro OS", "Observação"}),
    #"Colunas Removidas2" = Table.RemoveColumns(#"Colunas Reordenadas1",{"OS#(lf)disponibilizada?", "LOG", "Caixa#(lf)(20kg)", "Caixa#(lf)(Mídia)", "Caixa Tubo#(lf)(perfil de poço)", "Caixa Tubo#(lf)(Engenharia)", "Lacre", "Etiqueta (20kg)", "Etiqueta (Mídia)", "Etiqueta Tubo#(lf)(perfil de poço)", "Etiqueta Tubo#(lf)(Engenharia)", "Lançamento de etiqueta/Registro OS", "Observação"}),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Colunas Removidas2",{{"Período#(lf)faturamento", type date}, {"Situação", type text}, {"Contrato", type text}, {"Descrição da atividade", type text}, {"Item", type text}, {"Aplicação", type text}, {"Código da#(lf)solicitação", type text}, {"Chave#(lf)solicitante", type text}, {"Gerência solicitante", type text}, {"Código OS", type text}, {"Qtd.#(lf)Atendida", Int64.Type}, {"Qtd.#(lf)Solicitada", Int64.Type}, {"Data solicitação", type datetime}, {"Prazo#(lf)aplicação", type datetime}, {"D. abertura", type datetime}, {"D. fechamento", type datetime}, {"Prazo combinado", type datetime}, {"Obs. Sigla", type text}, {"Localidade", type text}, {"Comentários", type text}, {"Observações", type text}}),
    #"Texto em Maiúscula" = Table.TransformColumns(#"Tipo Alterado",{{"Descrição da atividade", Text.Upper, type text}, {"Localidade", Text.Upper, type text}}),
    #"Texto Limpo" = Table.TransformColumns(#"Texto em Maiúscula",{{"Localidade", Text.Clean, type text}}),
    #"Texto Aparado" = Table.TransformColumns(#"Texto Limpo",{{"Localidade", Text.Trim, type text}}),
    #"Linhas Classificadas" = Table.Sort(#"Texto Aparado",{{"Período#(lf)faturamento", Order.Ascending}}),
    #"Linhas Filtradas1" = Table.SelectRows(#"Linhas Classificadas", each ([Situação] = "CO")),
    #"Linhas Classificadas1" = Table.Sort(#"Linhas Filtradas1",{{"Contrato", Order.Ascending}}),
    #"Tipo Alterado1" = Table.TransformColumnTypes(#"Linhas Classificadas1",{{"Prazo combinado", type datetime}, {"D. fechamento", type datetime}, {"D. abertura", type datetime}, {"Prazo#(lf)aplicação", type datetime}, {"Data solicitação", type datetime}}),
    #"Colunas Reordenadas2" = Table.ReorderColumns(#"Tipo Alterado1",{"Período#(lf)faturamento", "Situação", "Contrato", "Descrição da atividade", "Item", "Aplicação", "Código da#(lf)solicitação", "Chave#(lf)solicitante", "Gerência solicitante", "Qtd.#(lf)Solicitada", "Qtd.#(lf)Atendida", "Código OS", "Data solicitação", "Prazo#(lf)aplicação", "D. abertura", "D. fechamento", "Prazo combinado", "Comentários", "Observações", "Obs. Sigla", "Localidade"}),
    #"Consultas Mescladas" = Table.NestedJoin(#"Colunas Reordenadas2", {"Localidade"}, #"Localidade Com Endereço", {"Localidade"}, "Localidade Com Endereço", JoinKind.LeftOuter),
    #"Localidade Com Endereço Expandido" = Table.ExpandTableColumn(#"Consultas Mescladas", "Localidade Com Endereço", {"Centro", "Município", "UF"}, {"Centro", "Município", "UF"}),
    #"Personalização Adicionada" = Table.AddColumn(#"Localidade Com Endereço Expandido", "ValidarGalpao", each if List.Contains(
    {
        "DEVOLUCAO DE EMPRESTIMO",
        "COLETA DE EMBALAGEM",
        "COLETA DE ITEM AVULSO",
        "ENTREGA DE EMBALAGEM EXPRESSO",
        "ENTREGA DE EMBALAGEM NORMAL",
        "ENTREGA DE ITEM EXPRESSO",
        "ENTREGA DE ITEM NORMAL"
    },
    Text.Upper(Text.Trim([Descrição da atividade]))
)
then "Verdadeiro"
else "Falso"),
    #"Coluna Condicional Adicionada" = Table.AddColumn(#"Personalização Adicionada", "Personalizar", each if [Obs. Sigla] = "GRJ" then "IRON-RJ" else if [Obs. Sigla] = "GES" then "IRON-ES" else if [Obs. Sigla] = "GSP" then "IRON-SP" else if [Obs. Sigla] = "GBA" then "PA-BA" else null),
    #"Consultas Mescladas1" = Table.NestedJoin(#"Coluna Condicional Adicionada", {"UF"}, validUFGalpao, {"UF"}, "validUFGalpao", JoinKind.LeftOuter),
    #"validUFGalpao Expandido" = Table.ExpandTableColumn(#"Consultas Mescladas1", "validUFGalpao", {"Contrato"}, {"Contrato.1"}),
    #"Coluna Condicional Adicionada1" = Table.AddColumn(#"validUFGalpao Expandido", "Galpão", each if [ValidarGalpao] = "Falso" then "N/A" else if [ValidarGalpao] = "Verdadeiro" then [Contrato.1] else null),
    #"Consultas Mescladas2" = Table.NestedJoin(#"Coluna Condicional Adicionada1", {"Galpão"}, valEnderecoGalpao, {"Galpãp"}, "valEnderecoGalpao", JoinKind.LeftOuter),
    #"valEnderecoGalpao Expandido" = Table.ExpandTableColumn(#"Consultas Mescladas2", "valEnderecoGalpao", {"Galpãp_Cidade", "Galpãp_UF"}, {"Galpãp_Cidade", "Galpãp_UF"}),
    #"Personalização Adicionada1" = Table.AddColumn(#"valEnderecoGalpao Expandido", "ColunaConcatenadaCalculoFDN", each [Descrição da atividade]&"/"&[#"Código da#(lf)solicitação"]&"/"&[Código OS]),
    #"Personalização Adicionada2" = Table.AddColumn(#"Personalização Adicionada1", "CalculoFDM", each List.Count(
    List.Select(
        Painel[ColunaConcatenadaCalculoFDN],
        (x) => x = [ColunaConcatenadaCalculoFDN]
      )
))
in
    #"Personalização Adicionada2"
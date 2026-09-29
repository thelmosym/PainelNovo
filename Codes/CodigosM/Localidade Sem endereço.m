let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    Localidade_Table = Fonte{[Item="Localidade",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(Localidade_Table,{{"Localidade", type text}, {"Centro", type text}}),
    #"Consultas Mescladas" = Table.NestedJoin(#"Tipo Alterado", {"Centro"}, Centro, {"Centro"}, "Centro.1", JoinKind.LeftOuter),
    #"Centro.1 Expandido" = Table.ExpandTableColumn(#"Consultas Mescladas", "Centro.1", {"Código Centro", "CNPJ", "Endereço", "CEP", "Bairro", "Município", "UF"}, {"Código Centro", "CNPJ", "Endereço", "CEP", "Bairro", "Município", "UF"}),
    #"Linhas Filtradas" = Table.SelectRows(#"Centro.1 Expandido", each ([Endereço] = null))
in
    #"Linhas Filtradas"

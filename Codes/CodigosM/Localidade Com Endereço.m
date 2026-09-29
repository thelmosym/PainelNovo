let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    Localidade_Table = Fonte{[Item="Localidade",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(Localidade_Table,{{"Localidade", type text}, {"Centro", type text}}),
    #"Texto em Maiúscula" = Table.TransformColumns(#"Tipo Alterado",{{"Localidade", Text.Upper, type text}, {"Centro", Text.Upper, type text}}),
    #"Texto Aparado" = Table.TransformColumns(#"Texto em Maiúscula",{{"Localidade", Text.Trim, type text}, {"Centro", Text.Trim, type text}}),
    #"Texto Limpo" = Table.TransformColumns(#"Texto Aparado",{{"Localidade", Text.Clean, type text}, {"Centro", Text.Clean, type text}}),
    #"Consultas Mescladas" = Table.NestedJoin(#"Texto Limpo", {"Centro"}, Centro, {"Centro"}, "Centro.1", JoinKind.LeftOuter),
    #"Centro.1 Expandido" = Table.ExpandTableColumn(#"Consultas Mescladas", "Centro.1", {"Código Centro", "CNPJ", "Endereço", "CEP", "Bairro", "Município", "UF"}, {"Código Centro", "CNPJ", "Endereço", "CEP", "Bairro", "Município", "UF"}),
    #"Linhas Filtradas" = Table.SelectRows(#"Centro.1 Expandido", each ([Endereço] <> null))
in
    #"Linhas Filtradas"

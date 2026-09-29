let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    Localidade_Table = Fonte{[Item="Localidade",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(Localidade_Table,{{"Localidade", type text}, {"Centro", type text}}),
    #"Texto em Maiúscula" = Table.TransformColumns(#"Tipo Alterado",{{"Localidade", Text.Upper, type text}, {"Centro", Text.Upper, type text}}),
    #"Texto Aparado" = Table.TransformColumns(#"Texto em Maiúscula",{{"Localidade", Text.Trim, type text}, {"Centro", Text.Trim, type text}}),
    #"Texto Limpo" = Table.TransformColumns(#"Texto Aparado",{{"Localidade", Text.Clean, type text}, {"Centro", Text.Clean, type text}})
in
    #"Texto Limpo"

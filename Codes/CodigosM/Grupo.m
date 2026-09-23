let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    Grupo_Table = Fonte{[Item="Grupo",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(Grupo_Table,{{"Município", type text}, {"UF", type text}, {"RMT", type text}, {"Região", type text}, {"Grupo", type text}, {"Contrato", type text}}),
    #"Colunas Reordenadas" = Table.ReorderColumns(#"Tipo Alterado",{"Contrato", "Município", "UF", "RMT", "Região", "Grupo"}),
    #"Personalização Adicionada" = Table.AddColumn(#"Colunas Reordenadas", "codGrupo", each [Município]&"\"&[UF]),
    #"Colunas Reordenadas1" = Table.ReorderColumns(#"Personalização Adicionada",{"codGrupo", "Contrato", "Município", "UF", "RMT", "Região", "Grupo"})
in
    #"Colunas Reordenadas1"

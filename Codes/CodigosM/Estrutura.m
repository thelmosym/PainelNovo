let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    Estrutura_Table = Fonte{[Item="Estrutura",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(Estrutura_Table,{{"Lotação", type text}, {"Diretoria Executiva", type text}, {"Gerência Executiva", type text}, {"Ativo", type text}})
in
    #"Tipo Alterado"

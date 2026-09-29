let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    Calendario_Table = Fonte{[Item="Calendario",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(Calendario_Table,{{"Data", type date}, {"Descrição do feriado", type text}, {"Tipo de feriado", type text}, {"Município", type text}, {"UF", type text}})
in
    #"Tipo Alterado"

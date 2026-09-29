let
    Fonte = (Parâmetro1) => let
        Fonte = Excel.Workbook(Parâmetro1, null, true),
        Monitoramento_Table = Fonte{[Item="Monitoramento",Kind="Table"]}[Data]
    in
        Monitoramento_Table
in
    Fonte

let
    Fonte = Excel.Workbook(localTabelaB2, null, true),
    Monitoramento_Table = Fonte{[Item="Monitoramento",Kind="Table"]}[Data]
in
    Monitoramento_Table

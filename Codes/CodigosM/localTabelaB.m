let
    Fonte = Excel.CurrentWorkbook(){[Name="localTabelaB"]}[Content],
    Column1 = Fonte{0}[Column1]
in
    Column1

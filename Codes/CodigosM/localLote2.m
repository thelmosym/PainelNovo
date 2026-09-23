let
    Fonte = Excel.CurrentWorkbook(){[Name="localLote2"]}[Content],
    Column1 = Fonte{0}[Column1]
in
    Column1

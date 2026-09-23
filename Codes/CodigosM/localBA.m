let
    Fonte = Excel.CurrentWorkbook(){[Name="localBA"]}[Content],
    Column1 = Fonte{0}[Column1]
in
    Column1

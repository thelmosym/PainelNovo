let
    Fonte = Excel.CurrentWorkbook(){[Name="localRJ"]}[Content],
    Column1 = Fonte{0}[Column1]
in
    Column1

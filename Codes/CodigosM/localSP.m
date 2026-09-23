let
    Fonte = Excel.CurrentWorkbook(){[Name="localSP"]}[Content],
    Column1 = Fonte{0}[Column1]
in
    Column1

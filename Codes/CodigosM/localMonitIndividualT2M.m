let
    Fonte = Excel.CurrentWorkbook(){[Name="localMonitIndividualT2M"]}[Content],
    Column1 = Fonte{0}[Column1]
in
    Column1

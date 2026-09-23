let
    Fonte = Excel.CurrentWorkbook(){[Name="PeriodoMes"]}[Content],
    Column1 = Fonte{0}[Column1]
in
    Column1

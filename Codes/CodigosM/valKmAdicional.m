let
    Fonte = Excel.CurrentWorkbook(){[Name="valKmAdicional"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Km a sere usados no calclo de KM adicional", Int64.Type}}),
    #"Km a sere usados no calclo de KM adicional" = #"Tipo Alterado"{0}[Km a sere usados no calclo de KM adicional]
in
    #"Km a sere usados no calclo de KM adicional"

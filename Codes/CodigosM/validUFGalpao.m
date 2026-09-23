let
    Fonte = Excel.CurrentWorkbook(){[Name="validUFGalpao"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"UF", type text}, {"Contrato", type text}})
in
    #"Tipo Alterado"
